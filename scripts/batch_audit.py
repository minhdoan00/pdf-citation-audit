r"""
batch_audit.py
Automated reference audit & author self-citation analysis for one or multiple scientific PDF manuscripts.
Supports single-file, multi-file, or recursive directory scanning.
Consolidates all verification findings into 2 central files:
  - citation_check_result.json (Machine-readable structured format for automation/APIs)
  - citation_check_result.xlsx (Human-readable 5-sheet interactive Excel workbook for editors/reviewers)

Usage:
    # 1. Scan an entire directory:
    python -X utf8 batch_audit.py "path/to/manuscripts_folder"

    # 2. Audit a single PDF paper:
    python -X utf8 batch_audit.py "path/to/paper.pdf"

    # 3. Audit multiple explicit PDF files:
    python -X utf8 batch_audit.py "paper1.pdf" "paper2.pdf" "paper3.pdf" -o "output_folder"
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

STATUS_ICON_EN = {
    "VERIFIED": "✅ Verified (Matches Scholarly Database)",
    "VERIFIED_WITH_METADATA_ERROR": "⚠️ Verified with Metadata Error (DOI/Year Discrepancy)",
    "FABRICATED_CITATION": "❌ Fabricated Citation / AI Hallucination",
    "BOOK_PATENT_DOMESTIC_DOC": "📘 Book / Patent / Standard / Regional Document",
    "UNVERIFIED_NEEDS_CHECK": "🔍 Unverified (Requires Manual Check)"
}

STATUS_ICON_VI = {
    "VERIFIED": "✅ Xác thực (Tài liệu thật khớp CSDL)",
    "VERIFIED_WITH_METADATA_ERROR": "⚠️ Tài liệu thật nhưng lỗi thông tin/DOI/năm",
    "FABRICATED_CITATION": "❌ Nghi ngờ giả mạo trích dẫn / AI Hallucination",
    "BOOK_PATENT_DOMESTIC_DOC": "📘 Sách / Sáng chế / Văn bản / Nội địa",
    "UNVERIFIED_NEEDS_CHECK": "🔍 Chưa xác thực tự động (Cần kiểm tra thủ công)"
}

def resolve_pdf_inputs(inputs_list: list, input_dir: str = None) -> list:
    """
    Resolves targets (individual files, directories, glob patterns)
    into a sorted, deduplicated list of absolute PDF file paths.
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

def generate_legacy_folder_reports(workspace: str, fname: str, authors: list, refs: list, results: list, status_icon: dict):
    """
    Backwards compatibility helper: creates Report_<stem> subdirectories when --create-folders is specified.
    """
    stem = os.path.splitext(fname)[0].rstrip(" .")
    report_dir = os.path.join(workspace, f"Report_{stem}")
    os.makedirs(report_dir, exist_ok=True)

    ref_list_path = os.path.join(report_dir, "Danh_Sach_Tai_Lieu_Tham_Khao.md")
    with open(ref_list_path, "w", encoding="utf-8") as f:
        f.write(f"# Extracted References List — {stem}\n\n")
        f.write(f"- **Manuscript:** `{fname}`\n")
        f.write(f"- **Authors:** {', '.join(authors) if authors else 'Unstated / Collective'}\n")
        f.write(f"- **Total References:** {len(refs)}\n\n---\n\n")
        for r in refs:
            f.write(f"**[{r['id']}]** {r['raw']}\n\n")

    md_path = os.path.join(report_dir, "Bao_Cao_Kiem_Tra_Trich_Dan.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(f"# CITATION VERIFICATION REPORT\n\n")
        f.write(f"- **Manuscript:** `{fname}`\n")
        f.write(f"- **Authors:** {', '.join(authors) if authors else 'Unstated'}\n")
        f.write(f"- **Total References:** **{len(refs)}**\n\n")
        f.write(f"| ID | Raw Citation | Self-Citation? | Verification Status | Cross-Check Details |\n")
        f.write(f"| :---: | :--- | :---: | :--- | :--- |\n")
        for r in results:
            raw_cell = r["raw"].replace("|", "\\|")
            det_cell = r["details"].replace("|", "\\|")
            st_cell = status_icon.get(r["status"], r["status"])
            self_cell = f"🔄 Yes ({', '.join(r.get('self_citation_authors', []))})" if r.get("is_self_citation") else "-"
            f.write(f"| [{r['id']}] | {raw_cell} | {self_cell} | {st_cell} | {det_cell} |\n")

def main():
    parser = argparse.ArgumentParser(
        description="Verify citation authenticity and analyze author self-citations from one or multiple PDF manuscripts. Outputs consolidated citation_check_result.json and .xlsx."
    )
    parser.add_argument("inputs", nargs="*", help="Paths to one or more PDF files or directories containing PDFs")
    parser.add_argument("--input", "-i", nargs="+", dest="input_flags", help="PDF file(s) or directory")
    parser.add_argument("--input-dir", help="Directory containing PDF manuscripts (backwards-compatible)")
    parser.add_argument("--output-dir", "-o", default=None, help="Output directory for reports (default: same as input)")
    parser.add_argument("--output-prefix", default="citation_check_result", help="Output filename prefix (default: citation_check_result)")
    parser.add_argument("--lang", default="en", choices=["en", "vi"], help="Language for outputs and terminal logs (en: English, vi: Vietnamese, default: en)")
    parser.add_argument("--email", default="academic_audit@vnu.edu.vn", help="Email for CrossRef Polite Pool priority queue")
    parser.add_argument("--cache-file", default=None, help="Path to api_cache.json")
    parser.add_argument("--workers", type=int, default=16, help="Concurrent verification threads (default: 16)")
    parser.add_argument("--create-folders", action="store_true", help="Also generate individual Report_<stem> folders (default: False)")
    parser.add_argument("--md-report", action="store_true", default=True, help="Generate companion Markdown summary (default: True)")
    args = parser.parse_args()

    is_vi = (args.lang.lower() == "vi")
    status_icon = STATUS_ICON_VI if is_vi else STATUS_ICON_EN

    # Resolve input PDF files
    raw_inputs = (args.inputs or []) + (args.input_flags or [])
    pdf_paths = resolve_pdf_inputs(raw_inputs, args.input_dir)

    if not pdf_paths:
        msg = "Lỗi: Không tìm thấy tệp .pdf nào hợp lệ!" if is_vi else "Error: No valid PDF files found in specified inputs!"
        print(msg)
        print("Example usage:")
        print("  python batch_audit.py \"path/to/folder\"")
        print("  python batch_audit.py paper1.pdf paper2.pdf")
        return

    # Determine output directory
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

    # Determine cache file location
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
    header_msg = "BẮT ĐẦU QUY TRÌNH KIỂM ĐỊNH TRÍCH DẪN & PHÂN TÍCH TỰ TRÍCH DẪN" if is_vi else "STARTING CITATION AUDIT & AUTHOR SELF-CITATION VERIFICATION"
    print(header_msg)
    print(f"- Target manuscripts: {len(pdf_paths)} PDF(s)")
    print(f"- Output directory: {out_dir}")
    print(f"- Persistent cache: {cache_path}")
    print(f"- Output language: {'Vietnamese (vi)' if is_vi else 'English (en)'}")
    print(f"- Legacy folder export: {'ENABLED' if args.create_folders else 'DISABLED (consolidated output)'}")
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
    print(f"Successfully extracted {total_refs} references across {len(pdf_paths)} manuscript(s).")

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
                    print(f"Progress: {completed}/{total_refs} references verified ({time.time()-t0:.1f}s)", flush=True)
                    auditor.save_cache()

    auditor.save_cache()

    # Aggregate structured results
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

        if args.create_folders:
            generate_legacy_folder_reports(out_dir, fname, authors, refs, ordered_results, status_icon)

    overall_self_rate = (tot_self_cit * 100.0 / total_refs) if total_refs else 0.0
    master_data = {
        "audit_version": "1.2",
        "audit_timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "language": args.lang,
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
    # 1. EXPORT MACHINE-READABLE JSON
    # =========================================================================
    json_path = os.path.join(out_dir, f"{args.output_prefix}.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(master_data, f, ensure_ascii=False, indent=2)
    print(f"\n[1/2] JSON report exported: {json_path}")

    # =========================================================================
    # 2. EXPORT HUMAN-READABLE EXCEL WORKBOOK (5 SHEETS)
    # =========================================================================
    xlsx_path = os.path.join(out_dir, f"{args.output_prefix}.xlsx")
    export_citation_audit_to_excel(xlsx_path, master_data, lang=args.lang)
    print(f"[2/2] Excel workbook exported: {xlsx_path}")

    # =========================================================================
    # 3. EXPORT COMPANION MARKDOWN SUMMARY
    # =========================================================================
    if args.md_report:
        md_path = os.path.join(out_dir, f"{args.output_prefix}.md")
        with open(md_path, "w", encoding="utf-8") as f:
            if is_vi:
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
            else:
                f.write(f"# CITATION AUDIT & AUTHOR SELF-CITATION REPORT ({len(pdf_paths)} MANUSCRIPTS)\n\n")
                f.write(f"## 1. Executive Summary\n\n")
                f.write(f"- **Total Manuscripts Audited:** **{len(pdf_paths)}**\n")
                f.write(f"- **Total References Cross-Checked:** **{total_refs}**\n")
                f.write(f"- **Author Self-Citations Identified:** **{tot_self_cit} / {total_refs} ({overall_self_rate:.2f}%)**\n")
                f.write(f"- **Verification Runtime:** {time.time()-t0:.1f} seconds\n\n")

                f.write("### Database Cross-Check Classification Breakdown\n\n")
                f.write("| Verification Category | Count | Ratio (%) | Scholarly Interpretation |\n")
                f.write("| :--- | :---: | :---: | :--- |\n")
                f.write(f"| ✅ Verified (Matches Database) | {tot_counts['VERIFIED']} | {tot_counts['VERIFIED']*100/total_refs:.2f}% | Exact match confirmed in CrossRef, DataCite, arXiv, or Europe PMC |\n")
                f.write(f"| 📘 Book / Patent / Standard / Regional | {tot_counts['BOOK_PATENT_DOMESTIC_DOC']} | {tot_counts['BOOK_PATENT_DOMESTIC_DOC']*100/total_refs:.2f}% | Monograph, textbook, patent, legal regulation, thesis, or domestic journal |\n")
                f.write(f"| ⚠️ Verified with Metadata / DOI Error | {tot_counts['VERIFIED_WITH_METADATA_ERROR']} | {tot_counts['VERIFIED_WITH_METADATA_ERROR']*100/total_refs:.2f}% | Existing paper with mismatched DOI, trailing zeroes, wrong year, or pasted body text |\n")
                f.write(f"| ❌ Fabricated Citation / AI Hallucination | {tot_counts['FABRICATED_CITATION']} | {tot_counts['FABRICATED_CITATION']*100/total_refs:.2f}% | Non-existent work; fabricated title/authors with fake (404) or hijacked DOI |\n")
                f.write(f"| 🔍 Unverified (Requires Manual Review) | {tot_counts['UNVERIFIED_NEEDS_CHECK']} | {tot_counts['UNVERIFIED_NEEDS_CHECK']*100/total_refs:.2f}% | Short citation lacking title/DOI or newly published work awaiting indexing |\n")
                f.write(f"| **TOTAL** | **{total_refs}** | **100.00%** | |\n\n---\n\n")

                if fab_items:
                    f.write(f"## 2. Fabricated Citations & AI Hallucinations ({len(fab_items)} item(s))\n\n")
                    for it in fab_items:
                        f.write(f"### ❌ Manuscript `{it['filename'][:6]}` — Ref `[{it['ref_id']}]`\n")
                        f.write(f"- **File:** `{it['filename']}`\n")
                        f.write(f"- **Citation in Paper:** `{it['raw']}`\n")
                        f.write(f"- **Classification:** **{it['label']}**\n")
                        f.write(f"- **Database Evidence:** {it['details']}\n\n")
                    f.write("---\n\n")

                if meta_items:
                    f.write(f"## 3. Actionable Metadata & Publishing Discrepancies ({len(meta_items)} item(s))\n\n")
                    f.write("| No. | Code | Ref [ID] | Issue Classification | Raw Citation in Paper | Cross-Check Details & Recommended Fix |\n")
                    f.write("| :---: | :---: | :---: | :--- | :--- | :--- |\n")
                    for idx, it in enumerate(meta_items, 1):
                        raw_esc = it["raw"].replace("|", "\\|")
                        det_esc = it["details"].replace("|", "\\|")
                        f.write(f"| {idx} | `{it['filename'][:6]}` | `[{it['ref_id']}]` | {it['label']} | {raw_esc} | {det_esc} |\n")
                    f.write("\n---\n\n")

                f.write("## 4. Manuscript Summary Table\n\n")
                f.write("| No. | Code | Authors | Total Refs | 🔄 Self-Cit (%) | ✅ Verified | 📘 Book/Reg. | ⚠️ Meta Error | ❌ Fabricated | 🔍 Manual Check |\n")
                f.write("| :---: | :---: | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |\n")
                for i, d in enumerate(master_list, 1):
                    c = d["counts"]
                    code = d["filename"][:6]
                    auth_str = ', '.join(d["authors"][:2]) + ('...' if len(d["authors"]) > 2 else '') if d["authors"] else "Collective"
                    self_str = f"{d['self_citation']['count']} ({d['self_citation']['rate']:.0f}%)"
                    f.write(f"| {i} | `{code}` | {auth_str} | {d['total_references']} | {self_str} | {c['VERIFIED']} | {c['BOOK_PATENT_DOMESTIC_DOC']} | {c['VERIFIED_WITH_METADATA_ERROR']} | {c['FABRICATED_CITATION']} | {c['UNVERIFIED_NEEDS_CHECK']} |\n")

    print(f"\nAudit completed in {time.time()-t0:.1f} seconds!")
    print(f"- JSON data: {json_path}")
    print(f"- Excel workbook: {xlsx_path}")

if __name__ == "__main__":
    main()
