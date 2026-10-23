r"""
batch_audit.py
Kiểm định danh mục tài liệu tham khảo và phân tích tự trích dẫn của tác giả từ một hoặc nhiều bản thảo PDF.
Hỗ trợ kiểm tra 1 tệp, nhiều tệp hoặc toàn bộ thư mục.
Xuất kết quả tập trung vào 2 tệp duy nhất:
  - citation_check_result.json (phục vụ máy tính, tự động hóa)
  - citation_check_result.xlsx (phục vụ người dùng, biên tập viên, giao diện trực quan 5 sheets)

Cách sử dụng:
    # 1. Quét toàn bộ thư mục:
    python -X utf8 batch_audit.py "g:\My Drive\60_Bai_Bao_PDF_Tu_Tieu_Ban"

    # 2. Quét một bản thảo duy nhất:
    python -X utf8 batch_audit.py "path/to/paper.pdf"

    # 3. Quét nhiều bản thảo cụ thể:
    python -X utf8 batch_audit.py "paper1.pdf" "paper2.pdf" "paper3.pdf"
"""

import os
import sys
import json
import time
import argparse
import datetime
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

from extract_references import extract_ref_block, parse_references_from_block, extract_manuscript_authors
from verifier import CitationAuditor, verify_single_reference, build_author_matchers
from excel_reporter import export_citation_audit_to_excel

STATUS_ICON = {
    "VERIFIED": "✅ Xác thực (Tài liệu thật khớp CSDL)",
    "VERIFIED_WITH_METADATA_ERROR": "⚠️ Tài liệu thật nhưng lỗi thông tin/DOI/năm",
    "FABRICATED_CITATION": "❌ Nghi ngờ giả mạo trích dẫn / AI Hallucination",
    "BOOK_PATENT_DOMESTIC_DOC": "📘 Sách / Sáng chế / Văn bản / Nội địa",
    "UNVERIFIED_NEEDS_CHECK": "🔍 Chưa xác thực tự động (Cần kiểm tra thủ công)"
}

def resolve_pdf_inputs(inputs_list: list, input_dir: str = None) -> list:
    """
    Nhận vào danh sách đường dẫn (tệp lẻ, thư mục, glob pattern)
    và trả về danh sách đường dẫn tệp PDF tuyệt đối, duy nhất, đã sắp xếp.
    """
    targets = []
    if input_dir:
        targets.append(input_dir)
    if inputs_list:
        if isinstance(inputs_list, str):
            targets.append(inputs_list)
        else:
            targets.extend(inputs_list)

    found_pdfs = []
    for item in targets:
        item = os.path.abspath(item)
        if os.path.isdir(item):
            for f in sorted(os.listdir(item)):
                if f.lower().endswith(".pdf"):
                    found_pdfs.append(os.path.join(item, f))
        elif os.path.isfile(item):
            if item.lower().endswith(".pdf"):
                found_pdfs.append(item)
        else:
            import glob
            for m in glob.glob(item):
                if os.path.isfile(m) and m.lower().endswith(".pdf"):
                    found_pdfs.append(os.path.abspath(m))

    seen = set()
    unique_pdfs = []
    for p in sorted(found_pdfs, key=lambda x: os.path.basename(x).lower()):
        if p not in seen:
            seen.add(p)
            unique_pdfs.append(p)
    return unique_pdfs

def generate_legacy_folder_reports(workspace: str, fname: str, authors: list, refs: list, results: list):
    """
    Tùy chọn tương thích ngược: tạo các thư mục con Report_<stem> khi có cờ --create-folders.
    """
    stem = os.path.splitext(fname)[0].rstrip(" .")
    report_dir = os.path.join(workspace, f"Report_{stem}")
    os.makedirs(report_dir, exist_ok=True)

    # 1. Danh sách TLTK trích xuất nguyên văn
    ref_list_path = os.path.join(report_dir, "Danh_Sach_Tai_Lieu_Tham_Khao.md")
    with open(ref_list_path, "w", encoding="utf-8") as f:
        f.write(f"# Danh sách Tài liệu tham khảo — {stem}\n\n")
        f.write(f"- **Tệp bản thảo gốc:** `{fname}`\n")
        f.write(f"- **Tác giả bản thảo:** {', '.join(authors) if authors else 'Không xác định'}\n")
        f.write(f"- **Tổng số tài liệu tham khảo trích xuất:** {len(refs)}\n\n---\n\n")
        for r in refs:
            f.write(f"**[{r['id']}]** {r['raw']}\n\n")

    # 2. Báo cáo chi tiết Markdown
    md_path = os.path.join(report_dir, "Bao_Cao_Kiem_Tra_Trich_Dan.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(f"# BÁO CÁO KIỂM TRA & XÁC THỰC TÀI LIỆU THAM KHẢO\n\n")
        f.write(f"## 1. Thông tin chung\n")
        f.write(f"- **Bản thảo:** `{fname}`\n")
        f.write(f"- **Tác giả bản thảo:** {', '.join(authors) if authors else 'Chưa xác định / Tập thể'}\n")
        f.write(f"- **Tổng số tài liệu tham khảo:** **{len(refs)}**\n\n")
        f.write(f"## 2. Bảng kết quả kiểm tra chi tiết\n\n")
        f.write(f"| TT | Trích dẫn gốc | Tự trích? | Trạng thái xác thực | Ghi chú CSDL |\n")
        f.write(f"| :---: | :--- | :---: | :--- | :--- |\n")
        for r in results:
            raw_cell = r["raw"].replace("|", "\\|")
            det_cell = r["details"].replace("|", "\\|")
            st_cell = STATUS_ICON.get(r["status"], r["status"])
            self_cell = f"🔄 Có ({', '.join(r.get('self_citation_authors', []))})" if r.get("is_self_citation") else "-"
            f.write(f"| [{r['id']}] | {raw_cell} | {self_cell} | {st_cell} | {det_cell} |\n")

def main():
    parser = argparse.ArgumentParser(
        description="Kiểm tra tính xác thực và tự trích dẫn danh mục tài liệu tham khảo từ một hoặc nhiều bản thảo PDF. Xuất tập trung vào citation_check_result.json và .xlsx."
    )
    parser.add_argument("inputs", nargs="*", help="Đường dẫn một hoặc nhiều tệp PDF hoặc thư mục chứa PDF")
    parser.add_argument("--input", "-i", nargs="+", dest="input_flags", help="Tệp PDF hoặc thư mục chứa PDF")
    parser.add_argument("--input-dir", help="Thư mục chứa các tệp PDF (hỗ trợ tương thích ngược)")
    parser.add_argument("--output-dir", "-o", default=None, help="Thư mục xuất báo cáo (mặc định cùng thư mục với tệp đầu vào)")
    parser.add_argument("--output-prefix", default="citation_check_result", help="Tên tiền tố tệp xuất (mặc định: citation_check_result)")
    parser.add_argument("--email", default="academic_audit@vnu.edu.vn", help="Email đăng ký CrossRef Polite Pool")
    parser.add_argument("--cache-file", default=None, help="Đường dẫn tệp api_cache.json")
    parser.add_argument("--workers", type=int, default=16, help="Số luồng xử lý đồng thời (mặc định 16)")
    parser.add_argument("--create-folders", action="store_true", help="Vẫn tạo các thư mục Report_<filename> riêng lẻ (mặc định: False)")
    parser.add_argument("--md-report", action="store_true", default=True, help="Xuất kèm tệp Markdown tổng hợp (mặc định: True)")
    args = parser.parse_args()

    # Thu thập đường dẫn các tệp PDF đầu vào
    raw_inputs = (args.inputs or []) + (args.input_flags or [])
    pdf_paths = resolve_pdf_inputs(raw_inputs, args.input_dir)

    if not pdf_paths:
        print("Lỗi: Không tìm thấy tệp .pdf nào hợp lệ từ đầu vào!")
        print("Ví dụ sử dụng:")
        print("  python batch_audit.py \"duong_dan_thu_muc\"")
        print("  python batch_audit.py paper1.pdf paper2.pdf")
        return

    # Xác định thư mục đầu ra (output_dir)
    if args.output_dir:
        out_dir = os.path.abspath(args.output_dir)
    elif args.input_dir and os.path.isdir(args.input_dir):
        out_dir = os.path.abspath(args.input_dir)
    elif raw_inputs and os.path.isdir(raw_inputs[0]):
        out_dir = os.path.abspath(raw_inputs[0])
    elif pdf_paths:
        parent_dirs = set(os.path.dirname(p) for p in pdf_paths)
        if len(parent_dirs) == 1:
            out_dir = list(parent_dirs)[0]
        else:
            out_dir = os.getcwd()
    else:
        out_dir = os.getcwd()

    os.makedirs(out_dir, exist_ok=True)

    # Xác định tệp api_cache.json
    if args.cache_file:
        cache_path = os.path.abspath(args.cache_file)
    else:
        potential_caches = [
            os.path.join(out_dir, "api_cache.json"),
            os.path.join(os.getcwd(), "api_cache.json")
        ]
        if pdf_paths:
            potential_caches.insert(0, os.path.join(os.path.dirname(pdf_paths[0]), "api_cache.json"))

        cache_path = potential_caches[0]
        for pc in potential_caches:
            if os.path.exists(pc):
                cache_path = pc
                break

    print(f"================================================================================")
    print(f"BẮT ĐẦU QUY TRÌNH KIỂM ĐỊNH TRÍCH DẪN & PHÂN TÍCH TỰ TRÍCH DẪN")
    print(f"- Số lượng bản thảo PDF tiếp nhận: {len(pdf_paths)}")
    print(f"- Thư mục xuất kết quả: {out_dir}")
    print(f"- Tệp bộ nhớ đệm (Cache): {cache_path}")
    print(f"- Tạo thư mục con từng bài: {'CÓ' if args.create_folders else 'KHÔNG (quản lý tập trung)'}")
    print(f"================================================================================")
    t0 = time.time()

    auditor = CitationAuditor(cache_file=cache_path, user_email=args.email)

    all_tasks = []
    paper_refs = {}
    paper_authors = {}
    paper_matchers = {}
    paper_stems = {}

    for path in pdf_paths:
        fname = os.path.basename(path)
        stem = os.path.splitext(fname)[0].rstrip(" .")
        paper_stems[fname] = stem

        authors = extract_manuscript_authors(path)
        paper_authors[fname] = authors

        matchers = [build_author_matchers(a) for a in authors if build_author_matchers(a)]
        paper_matchers[fname] = matchers

        h, txt = extract_ref_block(path)
        refs = parse_references_from_block(fname, h, txt)
        paper_refs[fname] = refs

        for r in refs:
            all_tasks.append((fname, r, matchers))

    total_refs = len(all_tasks)
    print(f"Trích xuất thành công {total_refs} tài liệu tham khảo từ {len(pdf_paths)} bản thảo.")

    results_by_paper = {os.path.basename(p): {} for p in pdf_paths}
    lock = threading.Lock()
    completed = 0

    def worker(item):
        fname, ref, matchers = item
        res = verify_single_reference(auditor, fname, ref, author_matchers=matchers)
        return fname, ref["id"], res

    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futures = [ex.submit(worker, item) for item in all_tasks]
        for fut in as_completed(futures):
            fname, rid, res = fut.result()
            with lock:
                results_by_paper[fname][rid] = res
                completed += 1
                if completed % 250 == 0 or completed == total_refs:
                    print(f"Tiến độ đối chiếu: {completed}/{total_refs} mục ({time.time()-t0:.1f}s)", flush=True)
                    auditor.save_cache()

    auditor.save_cache()

    # Tổng hợp dữ liệu đa tầng
    master_list = []
    tot_counts = {
        "VERIFIED": 0,
        "VERIFIED_WITH_METADATA_ERROR": 0,
        "FABRICATED_CITATION": 0,
        "BOOK_PATENT_DOMESTIC_DOC": 0,
        "UNVERIFIED_NEEDS_CHECK": 0
    }
    tot_self_cit = 0
    fab_items = []
    meta_items = []

    for path in pdf_paths:
        fname = os.path.basename(path)
        stem = paper_stems[fname]
        authors = paper_authors[fname]
        refs = paper_refs[fname]
        ordered_results = [results_by_paper[fname][r["id"]] for r in refs]

        # Thống kê cho bản thảo này
        p_counts = {k: 0 for k in tot_counts}
        p_self_cnt = 0
        p_self_authors = {}

        for r in ordered_results:
            st = r["status"]
            p_counts[st] = p_counts.get(st, 0) + 1
            tot_counts[st] = tot_counts.get(st, 0) + 1

            if r.get("is_self_citation"):
                p_self_cnt += 1
                tot_self_cit += 1
                for a in r.get("self_citation_authors", []):
                    p_self_authors[a] = p_self_authors.get(a, 0) + 1

            if st == "FABRICATED_CITATION":
                fab_items.append({
                    "filename": fname,
                    "ref_id": r["id"],
                    "raw": r["raw"],
                    "status": st,
                    "label": r["label"],
                    "details": r["details"]
                })
            elif st == "VERIFIED_WITH_METADATA_ERROR":
                meta_items.append({
                    "filename": fname,
                    "ref_id": r["id"],
                    "raw": r["raw"],
                    "status": st,
                    "label": r["label"],
                    "details": r["details"]
                })

        p_self_rate = (p_self_cnt * 100.0 / len(refs)) if refs else 0.0

        paper_entry = {
            "filename": fname,
            "filepath": path,
            "stem": stem,
            "authors": authors,
            "total_references": len(refs),
            "counts": p_counts,
            "self_citation": {
                "count": p_self_cnt,
                "rate": round(p_self_rate, 2),
                "authors_breakdown": p_self_authors
            },
            "references": ordered_results
        }
        master_list.append(paper_entry)

        # Nếu người dùng yêu cầu tạo thư mục riêng:
        if args.create_folders:
            generate_legacy_folder_reports(out_dir, fname, authors, refs, ordered_results)

    # Cấu trúc dữ liệu toàn diện
    overall_self_rate = (tot_self_cit * 100.0 / total_refs) if total_refs else 0.0
    master_data = {
        "audit_version": "1.2",
        "audit_timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": {
            "total_manuscripts": len(pdf_paths),
            "total_references": total_refs,
            "total_self_citations": tot_self_cit,
            "overall_self_citation_rate": round(overall_self_rate, 2),
            "counts": tot_counts
        },
        "flagged_issues": {
            "fabricated_citations": fab_items,
            "metadata_errors": meta_items
        },
        "manuscripts": master_list
    }

    # =========================================================================
    # 1. XUẤT TỆP JSON TẬP TRUNG (CHO MÁY TÍNH / TỰ ĐỘNG HÓA)
    # =========================================================================
    json_path = os.path.join(out_dir, f"{args.output_prefix}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(master_data, f, ensure_ascii=False, indent=2)
    print(f"\n[1/2] Đã xuất thành công tệp JSON: {json_path}")

    # =========================================================================
    # 2. XUẤT TỆP EXCEL TẬP TRUNG (CHO NGƯỜI DÙNG / BAN BIÊN TẬP)
    # =========================================================================
    xlsx_path = os.path.join(out_dir, f"{args.output_prefix}.xlsx")
    export_citation_audit_to_excel(xlsx_path, master_data)
    print(f"[2/2] Đã xuất thành công tệp Excel: {xlsx_path}")

    # =========================================================================
    # 3. XUẤT BÁO CÁO MARKDOWN TỔNG HỢP (NẾU ĐƯỢC BẬT)
    # =========================================================================
    if args.md_report:
        md_path = os.path.join(out_dir, f"{args.output_prefix}.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(f"# BÁO CÁO TỔNG HỢP KIỂM TRA, XÁC THỰC & PHÂN TÍCH TỰ TRÍCH DẪN ({len(pdf_paths)} BẢN THẢO)\n\n")
            f.write(f"## 1. Tổng quan kết quả kiểm tra\n\n")
            f.write(f"- **Tổng số bản thảo đã kiểm tra:** **{len(pdf_paths)} bản thảo**\n")
            f.write(f"- **Tổng số tài liệu tham khảo đã đối chiếu:** **{total_refs} mục**\n")
            f.write(f"- **Tổng số tài liệu tự trích dẫn:** **{tot_self_cit} / {total_refs} ({overall_self_rate:.2f}%)**\n")
            f.write(f"- **Thời gian thực hiện:** {time.time()-t0:.1f} giây\n\n")

            f.write("### Bảng thống kê phân loại tổng thể\n\n")
            f.write("| Trạng thái xác thực | Số lượng | Tỷ lệ (%) | Ý nghĩa |\n")
            f.write("| :--- | :---: | :---: | :--- |\n")
            f.write(f"| ✅ Tài liệu thật (Khớp CSDL quốc tế) | {tot_counts['VERIFIED']} | {tot_counts['VERIFIED']*100/total_refs:.2f}% | Khớp chính xác bài báo khoa học trên CrossRef, DataCite, arXiv, Europe PMC |\n")
            f.write(f"| 📘 Sách / Sáng chế / Pháp quy / Nội địa | {tot_counts['BOOK_PATENT_DOMESTIC_DOC']} | {tot_counts['BOOK_PATENT_DOMESTIC_DOC']*100/total_refs:.2f}% | Sách, giáo trình, bằng sáng chế, quyết định, đề tài, tạp chí nội địa |\n")
            f.write(f"| ⚠️ Tài liệu thật nhưng sai lệch Metadata / DOI | {tot_counts['VERIFIED_WITH_METADATA_ERROR']} | {tot_counts['VERIFIED_WITH_METADATA_ERROR']*100/total_refs:.2f}% | Nhầm DOI bài khác, thừa số 0, sai năm, lỗi dán nhầm văn bản thân bài |\n")
            f.write(f"| ❌ Giả mạo trích dẫn / AI Hallucination | {tot_counts['FABRICATED_CITATION']} | {tot_counts['FABRICATED_CITATION']*100/total_refs:.2f}% | Bài báo bịa đặt hoàn toàn và gắn kèm mã DOI giả (404) hoặc DOI của bài khác |\n")
            f.write(f"| 🔍 Chưa xác thực tự động (Cần kiểm tra thủ công) | {tot_counts['UNVERIFIED_NEEDS_CHECK']} | {tot_counts['UNVERIFIED_NEEDS_CHECK']*100/total_refs:.2f}% | Trích dẫn rút gọn thiếu tiêu đề/DOI hoặc công trình mới đang chờ lập chỉ mục |\n")
            f.write(f"| **TỔNG CỘNG** | **{total_refs}** | **100.00%** | |\n\n---\n\n")

            if fab_items:
                f.write(f"## 2. Chi tiết các trích dẫn giả mạo / AI Hallucination ({len(fab_items)} mục)\n\n")
                for it in fab_items:
                    f.write(f"### ❌ Bản thảo `{it['filename'][:6]}` — Mục `[{it['ref_id']}]`\n")
                    f.write(f"- **Tệp bản thảo:** `{it['filename']}`\n")
                    f.write(f"- **Nội dung trích dẫn:** `{it['raw']}`\n")
                    f.write(f"- **Kết luận xác thực:** **{it['label']}**\n")
                    f.write(f"- **Bằng chứng đối chiếu CSDL:** {it['details']}\n\n")
                f.write("---\n\n")

            if meta_items:
                f.write(f"## 3. Danh sách các tài liệu lỗi Metadata / Nhầm DOI ({len(meta_items)} mục)\n\n")
                f.write("| TT | Mã bài | Mục | Phân loại lỗi | Trích dẫn trong bản thảo | Chi tiết đối chiếu & Hướng khắc phục |\n")
                f.write("| :---: | :---: | :---: | :--- | :--- | :--- |\n")
                for idx, it in enumerate(meta_items, 1):
                    raw_esc = it["raw"].replace("|", "\\|")
                    det_esc = it["details"].replace("|", "\\|")
                    f.write(f"| {idx} | `{it['filename'][:6]}` | `[{it['ref_id']}]` | {it['label']} | {raw_esc} | {det_esc} |\n")
                f.write("\n---\n\n")

            f.write("## 4. Bảng tổng hợp theo từng bản thảo\n\n")
            f.write("| STT | Mã bài | Tác giả bản thảo | Tổng TLTK | 🔄 Tự trích (%) | ✅ Xác thực | 📘 Sách/Nội địa | ⚠️ Lỗi Meta | ❌ Giả mạo | 🔍 Cần KT |\n")
            f.write("| :---: | :---: | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |\n")
            for i, d in enumerate(master_list, 1):
                c = d["counts"]
                code = d["filename"][:6]
                auth_str = ', '.join(d["authors"][:2]) + ('...' if len(d["authors"]) > 2 else '') if d["authors"] else "Tập thể"
                self_str = f"{d['self_citation']['count']} ({d['self_citation']['rate']:.0f}%)"
                f.write(f"| {i} | `{code}` | {auth_str} | {d['total_references']} | {self_str} | {c['VERIFIED']} | {c['BOOK_PATENT_DOMESTIC_DOC']} | {c['VERIFIED_WITH_METADATA_ERROR']} | {c['FABRICATED_CITATION']} | {c['UNVERIFIED_NEEDS_CHECK']} |\n")

    print(f"\nHoàn tất toàn bộ quy trình kiểm định sau {time.time()-t0:.1f} giây!")
    print(f"- File JSON: {json_path}")
    print(f"- File Excel: {xlsx_path}")

if __name__ == "__main__":
    main()
