"""
extract_references.py
Mô-đun trích xuất danh mục tài liệu tham khảo và danh sách tác giả từ bản thảo PDF.
Hỗ trợ chuẩn hóa Unicode NFC, lọc header/footer trang an toàn, nối DOI bị ngắt dòng,
xử lý đa định dạng đánh số trích dẫn ([n], n., APA không đánh số, bullet points),
và trích xuất chính xác danh sách tác giả bản thảo từ trang bìa/trang 1.
"""

import os
import re
import unicodedata
import pymupdf

AFFILIATION_KEYWORDS = {
    'khoa', 'viện', 'trường', 'trung tâm', 'bộ môn', 'phòng', 'đại học', 'công ty',
    'university', 'institute', 'department', 'faculty', 'center', 'centre', 'school', 'laboratory', 'lab',
    'email', 'hà nội', 'hồ chí minh', 'đà nẵng', 'việt nam', 'vietnam', 'địa chỉ', 'tòa nhà', 'phường', 'quận',
    'đường', 'số 1', 'hoàng quốc việt', 'đại cồ việt', 'nguyễn trãi', 'bạch mai', 'nghĩa đô', 'thủ dầu một',
    'bách khoa', 'phenikaa', 'duy tân', 'khoa học tự nhiên', 'văn lang', 'huế', 'thái nguyên'
}

def clean_page_headers_footers(page_text: str) -> str:
    lines = [l for l in page_text.split('\n')]
    cleaned = []
    n = len(lines)
    for idx, line in enumerate(lines):
        s = line.strip()
        if re.search(r'Hội thảo [Kk]hoa học kỷ niệm 60 năm thành lập Hội Vật lý Việt Nam(?:\s*\(1966[-\u2013\u2014]2026\))?$', s, re.I):
            continue
        if (idx <= 2 or idx >= n - 3) and re.match(r'^\d{1,3}$', s):
            continue
        cleaned.append(line)
    return '\n'.join(cleaned)

def extract_manuscript_authors(pdf_path: str) -> list:
    """Trích xuất danh sách họ và tên các tác giả của bản thảo từ trang đầu tiên."""
    doc = pymupdf.open(pdf_path)
    p0 = unicodedata.normalize('NFC', doc[0].get_text())
    fname = os.path.basename(pdf_path)
    
    stem = os.path.splitext(fname)[0]
    title_from_fname = re.sub(r'^TB\d+\.\d+_', '', stem)
    title_words = set(re.findall(r'\b[A-ZÀ-Ỹa-zà-ỹ0-9]+\b', title_from_fname.lower()))
    
    m_abs = re.search(r'(?im)^\s*(tóm\s*tắt|abstract)\b', p0)
    header_area = p0[:m_abs.start()] if m_abs else p0[:1200]
    lines = [l.strip() for l in header_area.split('\n') if l.strip()]
    
    cleaned_lines = []
    for l in lines:
        if re.search(r'Hội thảo [Kk]hoa học|kỷ niệm 60 năm|1966[-–—]2026', l, re.I):
            continue
        if re.match(r'^\d{1,3}$', l):
            continue
        cleaned_lines.append(l)

    author_lines = []
    for l in cleaned_lines:
        l_low = l.lower()
        if any(kw in l_low for kw in AFFILIATION_KEYWORDS) or re.match(r'^\*?\s*email\b', l_low):
            continue
        words_l = re.findall(r'\b[A-ZÀ-Ỹa-zà-ỹ0-9]+\b', l_low)
        if words_l and all(w.isupper() for w in re.findall(r'\b\w+\b', l)):
            match_title = sum(1 for w in words_l if w in title_words) / len(words_l)
            if match_title >= 0.5:
                continue
        if any(kw in l_low for kw in ('phố hồ chí minh', 'thành phố', 'việt nam', 'hà nội')):
            continue
        author_lines.append(l)

    raw_author_str = ' '.join(author_lines)
    cleaned_str = re.sub(r'[\*∗\d]+(?=\s|,|;|$)', '', raw_author_str)
    tokens = re.split(r'[,;]|\s+và\s+|\s+and\s+', cleaned_str)
    authors = []
    for tok in tokens:
        tok_clean = ' '.join(tok.split()).strip('.*0123456789 -–:')
        words = tok_clean.split()
        if not words:
            continue
        while words and (words[0].isupper() and words[0].lower() in title_words and len(words) > 2):
            words.pop(0)
        tok_clean = ' '.join(words)
        tok_clean = re.sub(r'^(GS|PGS|TS|ThS|KS|BS|CN|Prof|Dr)\.?\s*', '', tok_clean, flags=re.I).strip()
        if len(tok_clean.split()) >= 2 and len(tok_clean) >= 4:
            if not any(kw in tok_clean.lower() for kw in ('khoa', 'viện', 'trường', 'đại học', 'email', 'phòng')):
                authors.append(tok_clean)
        elif len(tok_clean.split()) == 1 and len(tok_clean) >= 3 and not any(kw in tok_clean.lower() for kw in ('phố', 'hồ', 'chí', 'minh', 'khoa')):
            authors.append(tok_clean)

    if not authors and "bộ môn vật lý hạt nhân" in p0.lower():
        authors = ["Bộ môn Vật lý hạt nhân"]

    return authors

def extract_ref_block(pdf_path: str):
    doc = pymupdf.open(pdf_path)
    pages = []
    for p in doc:
        txt = unicodedata.normalize('NFC', p.get_text())
        txt = clean_page_headers_footers(txt)
        pages.append(txt)
    full_text = '\n'.join(pages)
    
    fname = os.path.basename(pdf_path)
    if fname.startswith("TB1.20"):
        return "EMBEDDED_TB1.20", full_text

    pattern = re.compile(
        r'(?im)^[\s\dIVXLC\.]*(tài\s*liệu\s*tham\s*khảo|references|bibliography|tài\s*liệu)\s*:?\s*$'
    )
    matches = list(pattern.finditer(full_text))
    if not matches:
        pattern2 = re.compile(r'(?i)(tài\s*liệu\s*tham\s*khảo|references|bibliography)')
        matches = list(pattern2.finditer(full_text))
    
    if not matches:
        return None, ""
    
    last_m = matches[-1]
    ref_text = full_text[last_m.end():].strip()
    return last_m.group(0).strip(), ref_text

def join_broken_urls_and_dois(chunk: str) -> str:
    for _ in range(3):
        chunk = re.sub(r'(10\.\d{4,9}/[a-zA-Z0-9_\-\./]*?[-/])\s*\n\s*([a-zA-Z0-9_\-\./]+)', r'\1\2', chunk)
        chunk = re.sub(r'(10\.\d{4,9}/j\.[a-zA-Z0-9_]+\.\d{4}\.\d{3,6})\s*\n\s*(\d{1,4})\b', r'\1\2', chunk)
        chunk = re.sub(r'(10\.\d{4,9}/[a-zA-Z0-9_\-\./]+)\s*\n\s*(\d{1,5})\b', r'\1\2', chunk)
    return chunk

def parse_references_from_block(fname: str, heading: str, ref_text: str):
    ref_text = re.sub(
        r'Hội thảo [Kk]hoa học kỷ niệm 60 năm thành lập Hội Vật lý Việt Nam(?:\s*\(1966[-\u2013\u2014]2026\))?',
        '',
        ref_text
    )
    ref_text = re.sub(r'\b10\.\s*(\d{4,9}/)', r'10.\1', ref_text)
    ref_text = re.sub(r'(10\.29333/ejmste/8500)\.(7\.\s+H\.)', r'\1.\n\2', ref_text)
    ref_text = re.sub(r'(?m)^11\.\s+12\.\s+', '12. ', ref_text)
    ref_text = re.sub(r'(?m)^(\d+)\.([A-Z])', r'\1. \2', ref_text)
    ref_text = re.sub(r'\.\s+53\]\.\s+', '.\n53. ', ref_text)

    if fname.startswith("TB1.20"):
        blocks = re.findall(
            r'-\s*Một số công trình tiêu biểu.*?:(.*?)-\s*Đánh giá chung',
            ref_text, flags=re.S | re.I
        )
        refs = []
        idx = 1
        for b in blocks:
            items = re.split(r'(?m)^\s*\+\s*', b)
            for it in items:
                it = join_broken_urls_and_dois(it)
                it_clean = ' '.join(it.split())
                if not it_clean:
                    continue
                refs.append({"id": str(idx), "raw": it_clean})
                idx += 1
        return refs

    if fname.startswith("TB3.08"):
        lines = [line.strip() for line in ref_text.split('\n') if line.strip()]
        merged = '\n'.join(lines)
        starts = [
            r"Nguyen, N\.-Q\., Kurosawa",
            r"Le, N\.-T\., Nguyen, N\.-Q\.",
            r"Le, N\.-T\. \(2020\)",
            r"Bui, D\.-K\.",
            r"Le, N\.-T\., et al\. \(2025\)\. ISO",
            r"Le, N\.-T\., et al\. \(2025\)\. Determination",
            r"Pham, H\.-L\.",
            r"Le, N\.-T\., et al\. \(2024\)",
            r"Le, N\.-T\. \(2024\)",
            r"Nguyen, N\.-Q\., Hoang, T\.-P\.H\.",
            r"Mai, V\.-D\.",
            r"Cong, M\.",
            r"O'Brien, M\., Lamperti",
            r"O'Brien, M\., Bueermann",
            r"IAEA \(2000\)",
            r"NIST \(n\.d\.\)"
        ]
        comb = "(" + "|".join(starts) + ")"
        parts = re.split(comb, merged)
        refs = []
        idx = 1
        for i in range(1, len(parts), 2):
            raw = ' '.join((parts[i] + parts[i+1]).split())
            refs.append({"id": str(idx), "raw": raw})
            idx += 1
        return refs

    n_bracket = len(re.findall(r'(?m)^\s*\[\d+\]\.?\s*', ref_text))
    n_dot = len(re.findall(r'(?m)^\s*\d+\.\s+', ref_text))
    use_bracket = n_bracket >= n_dot and n_bracket > 0
    
    if use_bracket:
        pattern = re.compile(r'(?m)^\s*\[(\d+)\]\.?\s*')
    else:
        pattern = re.compile(r'(?m)^\s*(\d+)\.\s*(?![0-9]{4,9}/)')
        
    matches = list(pattern.finditer(ref_text))
    valid_matches = []
    expected = 1
    for idx_m, m in enumerate(matches):
        num = int(m.group(1))
        if not valid_matches:
            if num in (1, 0):
                valid_matches.append(m)
                expected = num + 1
        else:
            if num == expected:
                valid_matches.append(m)
                expected = num + 1
            elif expected < num <= expected + 2:
                upcoming_nums = [int(x.group(1)) for x in matches[idx_m+1:idx_m+6]]
                if expected not in upcoming_nums:
                    valid_matches.append(m)
                    expected = num + 1
            elif num == expected - 1:
                upcoming_nums = [int(x.group(1)) for x in matches[idx_m+1:idx_m+4]]
                if expected in upcoming_nums:
                    valid_matches.append(m)
    
    refs = []
    for i, m in enumerate(valid_matches):
        start_pos = m.end()
        end_pos = valid_matches[i+1].start() if i + 1 < len(valid_matches) else len(ref_text)
        chunk = ref_text[start_pos:end_pos]
        chunk = re.sub(r'(?m)^Nhóm [AB]\s*[—\-–].*$', '', chunk)
        chunk = re.sub(r'(?s)\s*Bảng\s+1\s*:\s*Tham số kích thước.*$', '', chunk)
        chunk = join_broken_urls_and_dois(chunk)
        raw = ' '.join(chunk.split())
        refs.append({"id": m.group(1), "raw": raw})
    return refs
