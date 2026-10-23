"""
excel_reporter.py
Generates professional, multi-tab Excel (.xlsx) citation verification reports
for Journal Editors, Peer Reviewers, and Academic Committees.
Includes 5 dedicated worksheets:
  1. Overview & Summary
  2. Manuscripts Summary
  3. All References
  4. Flagged Issues
  5. Author Self-Citations
"""

import os
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Professional Palette
COLOR_HEADER_BG = "1F4E78"       # Deep Navy
COLOR_HEADER_TXT = "FFFFFF"      # White
COLOR_SUBHEADER_BG = "2F5597"    # Steel Blue
COLOR_ZEBRA = "F9FAFB"           # Very Light Gray

STATUS_CONFIG_EN = {
    "VERIFIED": {
        "fill": "E2EFDA",        # Soft Green
        "font": "276A3C",        # Dark Green
        "label": "✅ Verified (Matches Scholarly Database)",
        "meaning": "Exact match found in CrossRef, DataCite, arXiv, or Europe PMC."
    },
    "BOOK_PATENT_DOMESTIC_DOC": {
        "fill": "DDEBF7",        # Soft Blue
        "font": "1B4F72",        # Dark Blue
        "label": "📘 Book / Patent / Standard / Regional Document",
        "meaning": "Monograph, textbook, patent, technical standard, thesis, or regional journal."
    },
    "VERIFIED_WITH_METADATA_ERROR": {
        "fill": "FFF2CC",        # Soft Amber
        "font": "B25E00",        # Dark Amber
        "label": "⚠️ Verified with Metadata Error (DOI / Year Discrepancy)",
        "meaning": "Work exists, but manuscript has wrong DOI, extra zeroes, wrong year, or pasted body text."
    },
    "FABRICATED_CITATION": {
        "fill": "FCE4D6",        # Soft Red
        "font": "C00000",        # Dark Red
        "label": "❌ Fabricated Citation / AI Hallucination",
        "meaning": "Non-existent work; fabricated title/authors with fake (404) or hijacked DOI."
    },
    "UNVERIFIED_NEEDS_CHECK": {
        "fill": "EDEDED",        # Soft Gray
        "font": "595959",        # Dark Gray
        "label": "🔍 Unverified (Requires Manual Check)",
        "meaning": "Short citation lacking title/DOI or recently published work awaiting indexing."
    }
}

STATUS_CONFIG_VI = {
    "VERIFIED": {
        "fill": "E2EFDA",
        "font": "276A3C",
        "label": "✅ Xác thực (Tài liệu thật khớp CSDL)",
        "meaning": "Khớp chính xác bài báo khoa học trên CrossRef, DataCite, arXiv, Europe PMC."
    },
    "BOOK_PATENT_DOMESTIC_DOC": {
        "fill": "DDEBF7",
        "font": "1B4F72",
        "label": "📘 Sách / Sáng chế / Pháp quy / Nội địa",
        "meaning": "Sách chuyên khảo, giáo trình, bằng sáng chế, quyết định pháp quy, đề tài, tạp chí nội địa."
    },
    "VERIFIED_WITH_METADATA_ERROR": {
        "fill": "FFF2CC",
        "font": "B25E00",
        "label": "⚠️ Tài liệu thật nhưng lỗi Metadata / DOI",
        "meaning": "Tài liệu có thật nhưng tác giả dán nhầm DOI, thừa số 0, sai năm hoặc dán lẫn văn bản thân bài."
    },
    "FABRICATED_CITATION": {
        "fill": "FCE4D6",
        "font": "C00000",
        "label": "❌ Giả mạo trích dẫn / AI Hallucination",
        "meaning": "Trích dẫn hoàn toàn không tồn tại; tên bài, tác giả và số trang bị bịa đặt (AI Hallucination)."
    },
    "UNVERIFIED_NEEDS_CHECK": {
        "fill": "EDEDED",
        "font": "595959",
        "label": "🔍 Chưa xác thực tự động (Cần KT thủ công)",
        "meaning": "Trích dẫn viết rút gọn thiếu tiêu đề/DOI hoặc công trình mới đang chờ lập chỉ mục."
    }
}

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color=COLOR_HEADER_TXT)
HEADER_FILL = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color=COLOR_SUBHEADER_BG, end_color=COLOR_SUBHEADER_BG, fill_type="solid")
REGULAR_FONT = Font(name="Segoe UI", size=10)
BOLD_FONT = Font(name="Segoe UI", size=10, bold=True)

def apply_header_style(cell, text, subheader=False):
    cell.value = text
    cell.font = HEADER_FONT
    cell.fill = SUBHEADER_FILL if subheader else HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = THIN_BORDER

def auto_fit_columns(ws, min_widths=None, max_widths=None):
    min_w = min_widths or {}
    max_w = max_widths or {}
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row <= 2 and ws.max_column > 3 and cell.column == 1:
                continue
            v = str(cell.value or "")
            if "\n" in v:
                lines = v.split("\n")
                line_max = max(len(l) for l in lines)
                max_len = max(max_len, line_max)
            else:
                max_len = max(max_len, len(v))
        
        computed = max(max_len + 3, min_w.get(col_letter, 10))
        computed = min(computed, max_w.get(col_letter, 85))
        ws.column_dimensions[col_letter].width = computed

def export_citation_audit_to_excel(output_path: str, audit_data: dict, lang: str = "en"):
    """
    Exports citation_check_result.xlsx with 5 comprehensive worksheets.
    Supports lang='en' (default) and lang='vi'.
    """
    is_vi = (lang.lower() == "vi")
    status_cfg = STATUS_CONFIG_VI if is_vi else STATUS_CONFIG_EN

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    summary = audit_data.get("summary", {})
    manuscripts = audit_data.get("manuscripts", [])
    flagged = audit_data.get("flagged_issues", {})
    fab_list = flagged.get("fabricated_citations", [])
    meta_list = flagged.get("metadata_errors", [])

    # =========================================================================
    # SHEET 1: Overview & Summary
    # =========================================================================
    s1_title = "Tổng quan & Thống kê" if is_vi else "Overview & Summary"
    ws1 = wb.create_sheet(title=s1_title)
    ws1.views.sheetView[0].showGridLines = True
    ws1.sheet_properties.tabColor = "1F4E78"

    # Title Banner
    ws1.merge_cells("A1:G1")
    title_cell = ws1["A1"]
    title_cell.value = (
        "BÁO CÁO KIỂM ĐỊNH TÀI LIỆU THAM KHẢO & PHÂN TÍCH TỰ TRÍCH DẪN" if is_vi
        else "CITATION VERIFICATION & AUTHOR SELF-CITATION AUDIT REPORT"
    )
    title_cell.font = Font(name="Segoe UI", size=16, bold=True, color="1F4E78")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 40

    ws1.merge_cells("A2:G2")
    sub_cell = ws1["A2"]
    ts_str = audit_data.get('audit_timestamp', datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))
    sub_cell.value = (
        f"Thời điểm xuất: {ts_str} | Nguồn đối chiếu: CrossRef Polite Pool, DataCite, arXiv, Europe PMC" if is_vi
        else f"Generated: {ts_str} | Verified via CrossRef Polite Pool, DataCite, arXiv, Europe PMC REST APIs"
    )
    sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="595959")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 22

    # KPI Summary Cards
    ws1["A4"] = "CHỈ SỐ TỔNG THỂ" if is_vi else "KEY AUDIT METRICS"
    ws1["A4"].font = Font(name="Segoe UI", size=12, bold=True, color="1F4E78")

    if is_vi:
        kpis = [
            ("Tổng số bản thảo đã kiểm tra", summary.get("total_manuscripts", len(manuscripts)), "bản thảo"),
            ("Tổng số tài liệu tham khảo (TLTK)", summary.get("total_references", 0), "mục"),
            ("Tổng số tài liệu Tự trích dẫn", f"{summary.get('total_self_citations', 0)} ({summary.get('overall_self_citation_rate', 0):.2f}%)", "mục"),
            ("Số trích dẫn Giả mạo (AI Hallucination)", len(fab_list), "mục cảnh báo đỏ"),
            ("Số trích dẫn Lỗi Metadata / Nhầm DOI", len(meta_list), "mục cần sửa")
        ]
    else:
        kpis = [
            ("Total Manuscripts Audited", summary.get("total_manuscripts", len(manuscripts)), "manuscripts"),
            ("Total References Extracted", summary.get("total_references", 0), "references"),
            ("Author Self-Citations", f"{summary.get('total_self_citations', 0)} ({summary.get('overall_self_citation_rate', 0):.2f}%)", "citations"),
            ("Fabricated Citations (AI Hallucination)", len(fab_list), "critical alert"),
            ("Metadata & Publishing Errors", len(meta_list), "actionable issues")
        ]

    for idx, (label, val, note) in enumerate(kpis, 5):
        ws1[f"A{idx}"] = label
        ws1[f"A{idx}"].font = BOLD_FONT
        ws1[f"A{idx}"].border = THIN_BORDER
        
        ws1[f"B{idx}"] = val
        ws1[f"B{idx}"].font = Font(name="Segoe UI", size=11, bold=True, color="C00000" if ("Giả mạo" in label or "Fabricated" in label) and val > 0 else "1F4E78")
        ws1[f"B{idx}"].alignment = Alignment(horizontal="center")
        ws1[f"B{idx}"].border = THIN_BORDER
        
        ws1[f"C{idx}"] = note
        ws1[f"C{idx}"].font = Font(name="Segoe UI", size=9, italic=True, color="7F7F7F")
        ws1[f"C{idx}"].border = THIN_BORDER
        ws1.row_dimensions[idx].height = 24

    # Status Breakdown Table
    row_st = 12
    ws1[f"A{row_st-1}"] = "BẢNG PHÂN LOẠI TRẠNG THÁI XÁC THỰC CSDL QUỐC TẾ" if is_vi else "DATABASE VERIFICATION CLASSIFICATION BREAKDOWN"
    ws1[f"A{row_st-1}"].font = Font(name="Segoe UI", size=12, bold=True, color="1F4E78")

    headers_st = (
        ["Phân loại trạng thái", "Mã trạng thái", "Số lượng", "Tỷ lệ (%)", "Ý nghĩa & Hướng xử lý"] if is_vi
        else ["Verification Category", "Status Code", "Count", "Ratio (%)", "Scholarly Interpretation & Action"]
    )
    for col_idx, h in enumerate(headers_st, 1):
        apply_header_style(ws1.cell(row=row_st, column=col_idx), h)
    ws1.row_dimensions[row_st].height = 26

    counts = summary.get("counts", {})
    tot_refs = summary.get("total_references", 1) or 1
    status_order = [
        "VERIFIED",
        "BOOK_PATENT_DOMESTIC_DOC",
        "VERIFIED_WITH_METADATA_ERROR",
        "FABRICATED_CITATION",
        "UNVERIFIED_NEEDS_CHECK"
    ]

    for idx, skey in enumerate(status_order, row_st + 1):
        cfg = status_cfg.get(skey, {})
        cnt = counts.get(skey, 0)
        pct = (cnt * 100.0 / tot_refs)
        
        c1 = ws1.cell(row=idx, column=1, value=cfg.get("label", skey))
        c1.font = Font(name="Segoe UI", size=10, bold=True, color=cfg.get("font", "000000"))
        c1.fill = PatternFill(start_color=cfg.get("fill", "FFFFFF"), end_color=cfg.get("fill", "FFFFFF"), fill_type="solid")
        c1.border = THIN_BORDER

        c2 = ws1.cell(row=idx, column=2, value=skey)
        c2.font = REGULAR_FONT
        c2.border = THIN_BORDER

        c3 = ws1.cell(row=idx, column=3, value=cnt)
        c3.font = BOLD_FONT
        c3.alignment = Alignment(horizontal="center")
        c3.border = THIN_BORDER

        c4 = ws1.cell(row=idx, column=4, value=f"{pct:.2f}%")
        c4.font = REGULAR_FONT
        c4.alignment = Alignment(horizontal="right")
        c4.border = THIN_BORDER

        c5 = ws1.cell(row=idx, column=5, value=cfg.get("meaning", ""))
        c5.font = REGULAR_FONT
        c5.border = THIN_BORDER
        ws1.row_dimensions[idx].height = 24

    # Self-citation Benchmark Table
    row_self = 20
    ws1[f"A{row_self-1}"] = "QUY CHUẨN ĐÁNH GIÁ TỰ TRÍCH DẪN" if is_vi else "AUTHOR SELF-CITATION EVALUATION BENCHMARK"
    ws1[f"A{row_self-1}"].font = Font(name="Segoe UI", size=12, bold=True, color="1F4E78")

    headers_self = (
        ["Mức độ tự trích", "Tỷ lệ (%)", "Ý nghĩa học thuật & Khuyến nghị phản biện"] if is_vi
        else ["Self-Citation Level", "Threshold (%)", "Academic Significance & Editorial Recommendation"]
    )
    for col_idx, h in enumerate(headers_self, 1):
        apply_header_style(ws1.cell(row=row_self, column=col_idx), h, subheader=True)
    ws1.row_dimensions[row_self].height = 26

    if is_vi:
        self_bench = [
            ("🟢 Bình thường", "<= 25%", "Kế thừa tự nhiên các nghiên cứu trước đây của nhóm tác giả; phù hợp thông lệ quốc tế."),
            ("⚠️ Cần lưu ý", "25% - 50%", "Mức độ tự trích đáng chú ý; ban biên tập cần xem xét sự cân bằng với việc trích dẫn tài liệu quốc tế độc lập."),
            ("🔄 Báo cáo tổng kết nhóm", "> 50% (đến 100%)", "Thường gặp ở Báo cáo tổng kết chặng đường 10–20 năm thành lập phòng thí nghiệm/nhóm nghiên cứu (Retrospective Review). Việc tập trung trích dẫn bài của nhóm là hoàn toàn tự nhiên và mang tính lịch sử.")
        ]
    else:
        self_bench = [
            ("🟢 Normal Level", "<= 25%", "Natural inheritance of prior research by the author/group; conforms to standard international academic practice."),
            ("⚠️ Substantial Level", "25% - 50%", "Notable self-citation rate; editors should evaluate balance against independent literature in the field."),
            ("🔄 Group Retrospective Review", "> 50% (up to 100%)", "Typical for 10-20 year retrospective laboratory/research group review papers. Citing prior group publications is legitimate, historical, and inherent to the paper type.")
        ]

    for idx, (m, r, desc) in enumerate(self_bench, row_self + 1):
        ws1.cell(row=idx, column=1, value=m).font = BOLD_FONT
        ws1.cell(row=idx, column=2, value=r).font = BOLD_FONT
        ws1.cell(row=idx, column=3, value=desc).font = REGULAR_FONT
        for col_idx in range(1, 4):
            ws1.cell(row=idx, column=col_idx).border = THIN_BORDER
        ws1.row_dimensions[idx].height = 24

    auto_fit_columns(ws1, min_widths={"A": 30, "B": 24, "C": 14, "D": 14, "E": 65})

    # =========================================================================
    # SHEET 2: Manuscripts Summary
    # =========================================================================
    s2_title = "Danh sách Bản thảo" if is_vi else "Manuscripts Summary"
    ws2 = wb.create_sheet(title=s2_title)
    ws2.views.sheetView[0].showGridLines = True
    ws2.sheet_properties.tabColor = "2F5597"

    headers_m = (
        [
            "STT", "Mã bài", "Tên tệp bản thảo PDF", "Tác giả bản thảo",
            "Tổng TLTK", "Tự trích (mục)", "Tỷ lệ tự trích (%)", "Đánh giá Tự trích",
            "✅ Xác thực", "📘 Sách/Nội địa", "⚠️ Lỗi Meta", "❌ Giả mạo (AI)", "🔍 Cần KT thủ công",
            "Đánh giá sơ bộ"
        ] if is_vi else [
            "No.", "Code", "PDF Manuscript Filename", "Manuscript Authors",
            "Total Refs", "Self-Citations", "Self-Citation Rate (%)", "Self-Citation Context",
            "✅ Verified", "📘 Book/Regional", "⚠️ Metadata Error", "❌ Fabricated (AI)", "🔍 Manual Check",
            "Overall Status"
        ]
    )
    for col_idx, h in enumerate(headers_m, 1):
        apply_header_style(ws2.cell(row=1, column=col_idx), h)
    ws2.row_dimensions[1].height = 28

    for idx, m in enumerate(manuscripts, 2):
        c_counts = m.get("counts", {})
        sc = m.get("self_citation", {})
        sc_cnt = sc.get("count", 0)
        sc_rate = sc.get("rate", 0.0)
        fab_cnt = c_counts.get("FABRICATED_CITATION", 0)
        meta_cnt = c_counts.get("VERIFIED_WITH_METADATA_ERROR", 0)

        if is_vi:
            sc_eval = "🔄 Tổng kết nhóm" if sc_rate >= 50.0 else ("⚠️ Đáng chú ý" if sc_rate >= 25.0 else "🟢 Bình thường")
            if fab_cnt > 0:
                status_summary = "🔴 Có trích dẫn giả mạo (AI)"
            elif meta_cnt > 0:
                status_summary = "🟠 Có lỗi metadata/DOI"
            elif c_counts.get("UNVERIFIED_NEEDS_CHECK", 0) > 0:
                status_summary = "🟡 Cần rà soát thủ công"
            else:
                status_summary = "🟢 Đạt chuẩn CSDL"
            unknown_auth = "Tập thể / Chưa rõ"
        else:
            sc_eval = "🔄 Group Review" if sc_rate >= 50.0 else ("⚠️ Noteworthy" if sc_rate >= 25.0 else "🟢 Normal")
            if fab_cnt > 0:
                status_summary = "🔴 Contains Fabricated Citations (AI)"
            elif meta_cnt > 0:
                status_summary = "🟠 Contains Metadata / DOI Errors"
            elif c_counts.get("UNVERIFIED_NEEDS_CHECK", 0) > 0:
                status_summary = "🟡 Requires Manual Review"
            else:
                status_summary = "🟢 Fully Compliant"
            unknown_auth = "Collective / Unstated"

        fname = m.get("filename", "")
        code = fname[:6] if len(fname) >= 6 and fname[2] == "." else fname[:8]
        authors_str = ", ".join(m.get("authors", [])) or unknown_auth

        row_vals = [
            idx - 1,
            code,
            fname,
            authors_str,
            m.get("total_references", 0),
            sc_cnt,
            f"{sc_rate:.1f}%",
            sc_eval,
            c_counts.get("VERIFIED", 0),
            c_counts.get("BOOK_PATENT_DOMESTIC_DOC", 0),
            meta_cnt,
            fab_cnt,
            c_counts.get("UNVERIFIED_NEEDS_CHECK", 0),
            status_summary
        ]

        zebra_fill = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid") if (idx % 2 == 0) else None

        for col_idx, v in enumerate(row_vals, 1):
            cell = ws2.cell(row=idx, column=col_idx, value=v)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if zebra_fill:
                cell.fill = zebra_fill
            
            if col_idx in (1, 2, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")

            if col_idx == 12 and fab_cnt > 0:
                cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
            elif col_idx == 11 and meta_cnt > 0:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="B25E00")
            elif col_idx == 14 and fab_cnt > 0:
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

        ws2.row_dimensions[idx].height = 22

    ws2.freeze_panes = "C2"
    ws2.auto_filter.ref = ws2.dimensions
    auto_fit_columns(ws2, min_widths={"A": 6, "B": 10, "C": 35, "D": 28, "E": 12, "F": 12, "G": 14, "H": 18, "I": 12, "J": 14, "K": 12, "L": 14, "M": 16, "N": 24}, max_widths={"C": 50, "D": 40})

    # =========================================================================
    # SHEET 3: All References
    # =========================================================================
    s3_title = "Chi tiết Toàn bộ TLTK" if is_vi else "All References"
    ws3 = wb.create_sheet(title=s3_title)
    ws3.views.sheetView[0].showGridLines = True
    ws3.sheet_properties.tabColor = "385723"

    headers_all = (
        [
            "STT Toàn cục", "Mã bài", "Tên tệp bản thảo PDF", "Mục [ID]",
            "Nội dung trích dẫn trong bản thảo", "🔄 Tự trích?", "Tác giả tự trích",
            "Trạng thái xác thực", "Nhãn phân loại", "DOI xác thực", "Chi tiết đối chiếu CSDL & Ghi chú xử lý"
        ] if is_vi else [
            "Global No.", "Code", "PDF Manuscript Filename", "Ref [ID]",
            "Raw Citation in Manuscript", "🔄 Self-Citation?", "Cited Self-Author(s)",
            "Verification Status", "Status Label", "Verified DOI", "Database Cross-Check Details & Action Notes"
        ]
    )
    for col_idx, h in enumerate(headers_all, 1):
        apply_header_style(ws3.cell(row=1, column=col_idx), h)
    ws3.row_dimensions[1].height = 28

    global_stt = 1
    for m in manuscripts:
        fname = m.get("filename", "")
        code = fname[:6] if len(fname) >= 6 and fname[2] == "." else fname[:8]
        refs = m.get("references", [])

        for r in refs:
            row_idx = global_stt + 1
            st = r.get("status", "UNVERIFIED_NEEDS_CHECK")
            cfg = status_cfg.get(st, {})
            is_self = r.get("is_self_citation", False)
            self_authors = ", ".join(r.get("self_citation_authors", []))

            doi_val = r.get("doi") or ""
            if not doi_val and "DOI:" in r.get("details", ""):
                import re
                m_doi = re.search(r'10\.\d{4,9}/[-._;()/:A-Za-z0-9]+', r.get("details", ""))
                if m_doi:
                    doi_val = m_doi.group(0)

            yes_self = "🔄 Có" if is_vi else "🔄 Yes"
            row_vals = [
                global_stt,
                code,
                fname,
                f"[{r.get('id', '')}]",
                r.get("raw", ""),
                yes_self if is_self else "-",
                self_authors,
                st,
                cfg.get("label", r.get("label", "")),
                doi_val,
                r.get("details", "")
            ]

            fill_color = cfg.get("fill")
            for col_idx, v in enumerate(row_vals, 1):
                cell = ws3.cell(row=row_idx, column=col_idx, value=v)
                cell.font = REGULAR_FONT
                cell.border = THIN_BORDER
                
                if col_idx in (1, 2, 4, 6, 8):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in (5, 11):
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                else:
                    cell.alignment = Alignment(vertical="center")

                if st == "FABRICATED_CITATION":
                    cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                    if col_idx in (8, 9):
                        cell.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
                elif st == "VERIFIED_WITH_METADATA_ERROR":
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    if col_idx in (8, 9):
                        cell.font = Font(name="Segoe UI", size=10, bold=True, color="B25E00")
                elif is_self and col_idx == 6:
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="2F5597")

            ws3.row_dimensions[row_idx].height = 28
            global_stt += 1

    ws3.freeze_panes = "E2"
    ws3.auto_filter.ref = ws3.dimensions
    auto_fit_columns(ws3, min_widths={"A": 8, "B": 10, "C": 28, "D": 10, "E": 55, "F": 12, "G": 22, "H": 22, "I": 30, "J": 22, "K": 60}, max_widths={"E": 70, "K": 80})

    # =========================================================================
    # SHEET 4: Flagged Issues
    # =========================================================================
    s4_title = "Cảnh báo & Lỗi cần xử lý" if is_vi else "Flagged Issues"
    ws4 = wb.create_sheet(title=s4_title)
    ws4.views.sheetView[0].showGridLines = True
    ws4.sheet_properties.tabColor = "C00000"

    headers_fl = (
        [
            "STT", "Mức độ cảnh báo", "Mã bài", "Tên tệp bản thảo PDF",
            "Mục [ID]", "Phân loại lỗi", "Trích dẫn nguyên văn trong bản thảo",
            "Bằng chứng đối chiếu CSDL & Hướng khắc phục đề xuất"
        ] if is_vi else [
            "No.", "Alert Severity", "Code", "PDF Manuscript Filename",
            "Ref [ID]", "Issue Category", "Raw Citation in Manuscript",
            "Database Cross-Check Evidence & Recommended Editorial Action"
        ]
    )
    for col_idx, h in enumerate(headers_fl, 1):
        apply_header_style(ws4.cell(row=1, column=col_idx), h)
    ws4.row_dimensions[1].height = 28

    all_flagged_rows = []
    crit_tag = "❌ NGHIÊM TRỌNG (AI GIẢ MẠO)" if is_vi else "❌ CRITICAL (AI HALLUCINATION)"
    warn_tag = "⚠️ CẦN HIỆU CHỈNH METADATA" if is_vi else "⚠️ METADATA / DOI DISCREPANCY"

    for it in fab_list:
        all_flagged_rows.append((crit_tag, it))
    for it in meta_list:
        all_flagged_rows.append((warn_tag, it))

    for idx, (level, item) in enumerate(all_flagged_rows, 2):
        fname = item.get("filename", "")
        code = fname[:6] if len(fname) >= 6 and fname[2] == "." else fname[:8]
        rid = item.get("ref_id") or item.get("id") or ""

        row_vals = [
            idx - 1,
            level,
            code,
            fname,
            f"[{rid}]",
            item.get("label", ""),
            item.get("raw", ""),
            item.get("details", "")
        ]

        is_fab = ("CRITICAL" in level or "NGHIÊM TRỌNG" in level)
        fill_color = "FCE4D6" if is_fab else "FFF2CC"
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        for col_idx, v in enumerate(row_vals, 1):
            cell = ws4.cell(row=idx, column=col_idx, value=v)
            cell.font = REGULAR_FONT
            cell.fill = row_fill
            cell.border = THIN_BORDER
            
            if col_idx in (1, 2, 3, 5):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (7, 8):
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="center")

            if col_idx == 2:
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="C00000" if is_fab else "B25E00")

        ws4.row_dimensions[idx].height = 36

    ws4.freeze_panes = "D2"
    ws4.auto_filter.ref = ws4.dimensions
    auto_fit_columns(ws4, min_widths={"A": 6, "B": 24, "C": 10, "D": 30, "E": 10, "F": 30, "G": 55, "H": 65}, max_widths={"G": 70, "H": 80})

    # =========================================================================
    # SHEET 5: Author Self-Citations
    # =========================================================================
    s5_title = "Phân tích Tự trích dẫn" if is_vi else "Author Self-Citations"
    ws5 = wb.create_sheet(title=s5_title)
    ws5.views.sheetView[0].showGridLines = True
    ws5.sheet_properties.tabColor = "70AD47"

    headers_sc = (
        [
            "STT", "Mã bài", "Tác giả bản thảo", "Tên tệp bản thảo PDF",
            "Mục [ID]", "Tác giả được tự trích", "Trích dẫn nguyên văn trong bài",
            "Trạng thái CSDL", "Chi tiết xác thực CSDL"
        ] if is_vi else [
            "No.", "Code", "Manuscript Authors", "PDF Manuscript Filename",
            "Ref [ID]", "Matched Self-Author(s)", "Raw Citation in Manuscript",
            "Database Status", "Database Cross-Check Details"
        ]
    )
    for col_idx, h in enumerate(headers_sc, 1):
        apply_header_style(ws5.cell(row=1, column=col_idx), h, subheader=True)
    ws5.row_dimensions[1].height = 28

    sc_stt = 1
    for m in manuscripts:
        fname = m.get("filename", "")
        code = fname[:6] if len(fname) >= 6 and fname[2] == "." else fname[:8]
        m_authors = ", ".join(m.get("authors", [])) or ("Tập thể" if is_vi else "Collective")
        refs = m.get("references", [])

        for r in refs:
            if not r.get("is_self_citation"):
                continue

            row_idx = sc_stt + 1
            self_authors = ", ".join(r.get("self_citation_authors", []))

            row_vals = [
                sc_stt,
                code,
                m_authors,
                fname,
                f"[{r.get('id', '')}]",
                self_authors,
                r.get("raw", ""),
                r.get("status", ""),
                r.get("details", "")
            ]

            zebra_fill = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid") if (sc_stt % 2 == 0) else None

            for col_idx, v in enumerate(row_vals, 1):
                cell = ws5.cell(row=row_idx, column=col_idx, value=v)
                cell.font = REGULAR_FONT
                cell.border = THIN_BORDER
                if zebra_fill:
                    cell.fill = zebra_fill
                
                if col_idx in (1, 2, 5, 8):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in (7, 9):
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                else:
                    cell.alignment = Alignment(vertical="center")

                if col_idx == 6:
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="1F4E78")

            ws5.row_dimensions[row_idx].height = 26
            sc_stt += 1

    ws5.freeze_panes = "C2"
    ws5.auto_filter.ref = ws5.dimensions
    auto_fit_columns(ws5, min_widths={"A": 6, "B": 10, "C": 26, "D": 30, "E": 10, "F": 24, "G": 55, "H": 20, "I": 55}, max_widths={"G": 70, "I": 75})

    wb.save(output_path)
    print(f"Exported Excel report successfully: {output_path}")
