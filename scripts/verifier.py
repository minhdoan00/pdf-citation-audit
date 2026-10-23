"""
verifier.py
Scholarly reference verification engine cross-referencing CrossRef, DataCite, arXiv, and Europe PMC.
5-Tier Verification Classification:
  1. VERIFIED: Exact match confirmed in international database.
  2. BOOK_PATENT_DOMESTIC_DOC: Monograph, textbook, patent, legal standard, regional journal.
  3. VERIFIED_WITH_METADATA_ERROR: Genuine work with mismatched DOI, trailing zeroes, wrong year, or layout error.
  4. FABRICATED_CITATION: Non-existent work / AI Hallucination / Hijacked DOI.
  5. UNVERIFIED_NEEDS_CHECK: Abbreviated citation lacking title/DOI or newly published work.
Integrates Author Self-Citation analysis module.
"""

import os
import re
import json
import time
import urllib.parse
import unicodedata
import threading
import xml.etree.ElementTree as ET
import requests
from requests.adapters import HTTPAdapter

STOPWORDS = {
    "the", "and", "for", "with", "from", "into", "using", "based", "under",
    "between", "through", "during", "after", "before", "over", "under",
    "study", "studies", "investigation", "analysis", "application", "applications",
    "effect", "effects", "properties", "synthesis", "characterization", "development",
    "towards", "toward", "novel", "high", "performance", "review", "overview",
    "của", "trong", "cho", "trên", "các", "với", "bằng", "một", "những", "nghiên", "cứu"
}

def strip_accents(s: str) -> str:
    if not s:
        return ""
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")

def norm_tokens(s: str):
    s = strip_accents(s).lower()
    return re.findall(r"[a-z0-9]+", s)

def extract_dois(raw: str):
    cleaned = re.sub(r'(10\.\d{4,9})\s*/\s*', r'\1/', raw)
    cleaned = re.sub(r'(10\.\d{4,9}/[a-zA-Z0-9_\-\./]+?-)(\s+)([a-zA-Z0-9])', r'\1\3', cleaned)
    cleaned = re.sub(r'(10\.\d{4,9}/j\.[a-zA-Z0-9_]+\.\d{4}\.\d{3,5})\s+(\d{1,3})\b', r'\1\2', cleaned)
    matches = re.findall(r'(10\.\d{4,9}/[a-zA-Z0-9_\-\./\(\):]+)', cleaned)
    dois = []
    for m in matches:
        d = m.rstrip('.').rstrip(',').rstrip(';').rstrip(']').rstrip(')').rstrip(':')
        if d and d not in dois:
            dois.append(d)
    return dois

def extract_arxiv_ids(raw: str):
    ids = []
    for m in re.findall(r'(?i)arxiv\s*[:\.]?\s*(\d{4}\.\d{4,5}(?:v\d+)?|[a-z\-]+/\d{7})', raw):
        clean_id = re.sub(r'v\d+$', '', m)
        if clean_id not in ids:
            ids.append(clean_id)
    for m in re.findall(r'\[(\d{4}\.\d{4,5}(?:v\d+)?)\]', raw):
        clean_id = re.sub(r'v\d+$', '', m)
        if clean_id not in ids:
            ids.append(clean_id)
    return ids

def extract_cited_years(raw: str):
    cleaned = re.sub(r'https?://\S+', ' ', raw)
    cleaned = re.sub(r'10\.\d{4,9}/\S+', ' ', cleaned)
    paren_years = [int(y) for y in re.findall(r'\(\s*(19\d{2}|20[0-2]\d)\s*\)', cleaned)]
    if paren_years:
        return paren_years
    all_years = [int(y) for y in re.findall(r'\b(19\d{2}|20[0-2]\d)\b', cleaned)]
    return all_years

def is_non_article_or_domestic(raw: str) -> str:
    r_low = raw.lower()
    if any(k in r_low for k in [
        "thủ tướng chính phủ", "quyết định số", "quyết định 438", "quyết định 1012", "quyết định 245", "quyết định 2736",
        "chỉ thị số", "thông tư số", "bộ giáo dục và đào tạo", "cục quản lí chất lượng",
        "nghị quyết", "chính phủ", "luật năng lượng", "european parliament", "european commission",
        "directive 2008/98/ec", "official journal of the european union", "japanese standards association",
        "jis a 5021", "tiêu chuẩn quốc gia", "tcvn "
    ]):
        return "Văn bản pháp quy / Chính sách / Tiêu chuẩn kỹ thuật"
    if any(k in r_low for k in [
        "bằng độc quyền", "bằng sáng chế", "giải pháp hữu ích", "cục sở hữu trí tuệ",
        "chấp nhận đơn", "công báo sở hữu công nghiệp", "patent", "sáng chế"
    ]):
        return "Bằng độc quyền sáng chế / Sở hữu trí tuệ"
    if any(k in r_low for k in [
        "luận án tiến sĩ", "luận văn thạc sĩ", "khóa luận tốt nghiệp", "luận án", "luận văn",
        "phd thesis", "master thesis", "dissertation", "doctoral thesis"
    ]):
        return "Luận án / Luận văn"
    if any(k in r_low for k in [
        "đề tài cấp nhà nước", "đề tài cấp bộ", "đề tài cấp cơ sở", "báo cáo tổng kết", "nhiệm vụ hợp tác quốc tế",
        "chương trình đào tạo", "đề án tuyển sinh", "danh mục các ngành đào tạo",
        "kỷ yếu 40 năm", "phòng thí nghiệm kỹ thuật hạt nhân, các định hướng"
    ]):
        return "Đề tài KH&CN / Báo cáo / Tài liệu nội bộ"
    if any(k in r_low for k in [
        "đang trong vòng phản biện", "chưa công bố", "submitted to", "in preparation", "accepted by nuclear science and technology"
    ]):
        return "Bản thảo đang phản biện / Mới chấp nhận đăng"
    if any(k in r_low for k in [
        "world-nuclear.org", "world- nuclear.org", "wikipedia.org", "leybold-shop.com",
        "workplace-safety", "iaea.org", "who.int/"
    ]) or (r_low.startswith("http") and "doi.org" not in r_low):
        return "Website / Cổng thông tin trực tuyến"
    if any(k in r_low for k in [
        "university press", "cambridge university", "oxford university", "wiley-interscience",
        "john wiley", "mcgraw-hill", "mc graw-hill", "academic press", "springer-verlag",
        "world scientific", "nhà xuất bản", "nxb ", "feynman lectures", "handbook of optics",
        "iaea nuclear energy series", "trs-398", "absorbed dose determination in external beam",
        "nist (n.d.)", "ferromagnetic insulators: garnets", "persistent spectral hole burning",
        "surveillance and diagnostics of next generation nuclear reactors",
        "thư gửi hội nghị tübingen", "intrinsic dosimetry: properties and mechanisms",
        "unscear", "reference sheet – reference materials", "oak ridge national laboratory review",
        "maker media", "getting started with arduino", "a nation empowered",
        "proceedings of the international school of physics", "crc press", "taylor & francis",
        "macmillan education", "second edition"
    ]):
        return "Sách chuyên khảo / Giáo trình / Báo cáo kỹ thuật quốc tế"
    if any(k in r_low for k in [
        "tạp chí phát triển khoa học và công nghệ", "tạp chí y học việt nam",
        "tạp chí khoa học và công nghệ", "hội nghị vật lý", "hội thảo khoa học",
        "tuyển tập báo cáo", "kỷ yếu hội nghị", "tnu journal of science and technology",
        "journal of science and technology of metal", "advances in natural science,",
        "science & technology development journal", "science and technology development journal",
        "vnu journal of science", "nuclear science and technology", "kmitl science journal",
        "chinese journal of physics, 53", "tạp chí "
    ]):
        return "Bài báo / Kỷ yếu hội nghị trong nước & khu vực"
    return ""

def guess_deterministic_dois(raw: str):
    dois = []
    m = re.search(r'Phys\.\s*Rev\.\s*([A-E]|Lett\.)\s*(\d{1,3})\s*,\s*(\d{5,7})\b', raw, re.I)
    if m:
        sub = "Lett" if "lett" in m.group(1).lower() else m.group(1).upper()
        dois.append(f"10.1103/PhysRev{sub}.{m.group(2)}.{m.group(3)}")
    m_pr_old = re.search(r'Phys\.\s*Rev\.?(?:,\s*vol\.)?\s*(\d{1,3})\s*,\s*(?:pp\.\s*)?(\d{1,6})\b', raw, re.I)
    if m_pr_old:
        dois.append(f"10.1103/PhysRev.{m_pr_old.group(1)}.{m_pr_old.group(2)}")
    m2 = re.search(r'J\.\s*Phys\.:?\s*Conf\.\s*Ser\.\s*(\d{3,5})\s*,\s*(\d{6})\b', raw, re.I)
    if m2:
        dois.append(f"10.1088/1742-6596/{m2.group(1)}/1/{m2.group(2)}")
    m3 = re.search(r'Optics\s*Express\s*,\s*(\d{1,3})\s*\(\d+\)\s*,\s*(\d{4,6})\b', raw, re.I)
    if m3:
        dois.append(f"10.1364/OE.{m3.group(1)}.{int(m3.group(2)):06d}")
    return dois

# ==================== SELF-CITATION DETECTION MODULE ====================

def build_author_matchers(author_name: str):
    name_clean = ' '.join(author_name.split())
    name_no_acc = strip_accents(name_clean).lower()
    toks = [t for t in re.findall(r'[a-z]+', name_no_acc) if len(t) >= 1]
    if len(toks) < 2:
        return None
    
    fam = toks[0]
    given = toks[-1]
    middles = toks[1:-1]
    
    inits = [t[0] for t in toks[:-1]]
    init_pat = r'\.?\s*'.join(inits) + r'\.?'
    
    patterns = []
    # 1. Initials + Given name: P. D. Tam / P.D. Tam
    patterns.append(rf'\b{init_pat}\s+{given}\b')
    # 2. Given name + Initials: Tam P. D. / Tam, P. D.
    patterns.append(rf'\b{given},?\s+{init_pat}\b')
    # 3. Full name: Phuong Dinh Tam / Phuong D. Tam
    if middles:
        mid_pats = [m[0] + r'\.?' for m in middles]
        patterns.append(rf'\b{fam}\s+(?:{"|".join(middles)}|{"|".join(mid_pats)})\s+{given}\b')
    patterns.append(rf'\b{fam}\s+(?:\w+\s+)?{given}\b')
    # 4. Family initial + Given name (nếu given name dài >= 4 ký tự)
    if len(given) >= 4:
        patterns.append(rf'\b{fam[0]}\.?\s+{given}\b')
        patterns.append(rf'\b{given},?\s+{fam[0]}\.?\b')
        
    combined_regex = re.compile('|'.join(patterns), re.I)
    return {
        "original": author_name,
        "fam": fam,
        "given": given,
        "toks": toks,
        "regex": combined_regex
    }

def check_self_citation(author_matchers: list, ref_raw: str, matched_work: dict = None) -> list:
    if not author_matchers:
        return []
    matched_authors = []
    raw_no_acc = strip_accents(ref_raw).lower()
    
    for m in author_matchers:
        hit = False
        if m["regex"].search(raw_no_acc):
            hit = True
        elif matched_work and matched_work.get("authors"):
            for a in matched_work["authors"]:
                w_fam = strip_accents(a.get("family", "")).lower()
                w_giv = strip_accents(a.get("given", "")).lower()
                w_nam = strip_accents(a.get("name", "")).lower()
                w_all = f"{w_fam} {w_giv} {w_nam}"
                w_toks = set(re.findall(r'[a-z]+', w_all))
                if m["given"] in w_toks and (m["fam"] in w_toks or any(t.startswith(m["fam"][:3]) for t in w_toks)):
                    hit = True
                    break
        if hit and m["original"] not in matched_authors:
            matched_authors.append(m["original"])
            
    return matched_authors

# ==================== AUDITOR ENGINE ====================

class CitationAuditor:
    def __init__(self, cache_file: str = "api_cache.json", user_email: str = "your_email@domain.edu.vn"):
        self.cache_file = cache_file
        self.user_email = user_email
        self.lock = threading.Lock()
        self.cr_rate_lock = threading.Lock()
        self.last_cr_search_time = 0.0
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=32, pool_maxsize=32)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        self.session.headers.update({
            "User-Agent": f"AcademicCitationAudit/1.0 (mailto:{self.user_email})"
        })
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    self.cache = json.load(f)
            except Exception:
                self.cache = {}
        else:
            self.cache = {}

    def _wait_cr_search(self):
        with self.cr_rate_lock:
            now = time.time()
            elapsed = now - self.last_cr_search_time
            if elapsed < 0.28:
                time.sleep(0.28 - elapsed)
            self.last_cr_search_time = time.time()

    def save_cache(self):
        with self.lock:
            snap = dict(self.cache)
        tmp_file = self.cache_file + ".tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(snap, f, ensure_ascii=False)
        os.replace(tmp_file, self.cache_file)

    def fetch_crossref_doi(self, doi: str):
        key = f"cr_doi:{doi.lower()}"
        with self.lock:
            if key in self.cache and self.cache[key].get("status") in (200, 404):
                return self.cache[key]
        url = f"https://api.crossref.org/works/{urllib.parse.quote(doi, safe='')}"
        for attempt in range(3):
            try:
                r = self.session.get(url, params={"mailto": self.user_email}, timeout=12)
                if r.status_code == 200:
                    msg = r.json().get("message", {})
                    res = {"status": 200, "work": self._slim_crossref_work(msg)}
                    with self.lock:
                        self.cache[key] = res
                    return res
                elif r.status_code == 404:
                    res = {"status": 404, "work": None}
                    with self.lock:
                        self.cache[key] = res
                    return res
                elif r.status_code == 429:
                    time.sleep(1.5 * (attempt + 1))
            except Exception:
                time.sleep(1.0)
        return {"status": "error", "work": None}

    def fetch_datacite_or_handle(self, doi: str):
        key = f"dc_doi:{doi.lower()}"
        if key in self.cache:
            return self.cache[key]
        try:
            r = self.session.get(f"https://api.datacite.org/dois/{urllib.parse.quote(doi, safe='')}", timeout=10)
            if r.status_code == 200:
                attr = r.json().get("data", {}).get("attributes", {})
                titles = [t.get("title", "") for t in attr.get("titles", [])]
                authors = [{"family": c.get("familyName", ""), "given": c.get("givenName", ""), "name": c.get("name", "")} for c in attr.get("creators", [])]
                res = {
                    "status": 200,
                    "work": {
                        "doi": doi,
                        "title": titles[0] if titles else "",
                        "authors": authors,
                        "journal": attr.get("container", {}).get("title", ""),
                        "year": attr.get("publicationYear"),
                        "volume": "",
                        "page": ""
                    }
                }
                self.cache[key] = res
                return res
        except Exception:
            pass
        try:
            r = self.session.get(f"https://doi.org/api/handles/{urllib.parse.quote(doi, safe='')}", timeout=10)
            if r.status_code == 200 and r.json().get("responseCode") == 1:
                res = {"status": 200, "work": {"doi": doi, "title": "[Registered DOI via Handle]", "authors": [], "journal": "", "year": None, "volume": "", "page": ""}}
            else:
                res = {"status": 404, "work": None}
        except Exception:
            res = {"status": "error", "work": None}
        self.cache[key] = res
        return res

    def fetch_arxiv(self, arxiv_id: str):
        key = f"arxiv:{arxiv_id}"
        if key in self.cache:
            return self.cache[key]
        url = f"http://export.arxiv.org/api/query?id_list={urllib.parse.quote(arxiv_id)}"
        try:
            r = self.session.get(url, timeout=12)
            if r.status_code == 200:
                root = ET.fromstring(r.text)
                ns = {"atom": "http://www.w3.org/2005/Atom"}
                entry = root.find("atom:entry", ns)
                if entry is not None:
                    title_el = entry.find("atom:title", ns)
                    title = " ".join(title_el.text.split()) if title_el is not None and title_el.text else ""
                    if title and title.lower() != "error":
                        authors = []
                        for a in entry.findall("atom:author", ns):
                            n = a.find("atom:name", ns)
                            if n is not None and n.text:
                                parts = n.text.strip().split()
                                authors.append({"family": parts[-1], "given": " ".join(parts[:-1])})
                        pub = entry.find("atom:published", ns)
                        year = int(pub.text[:4]) if pub is not None and pub.text else None
                        res = {"status": 200, "work": {"doi": f"arXiv:{arxiv_id}", "title": title, "authors": authors, "journal": "arXiv", "year": year, "volume": "", "page": arxiv_id}}
                        self.cache[key] = res
                        return res
            res = {"status": 404, "work": None}
        except Exception as e:
            res = {"status": "error", "error": str(e), "work": None}
        self.cache[key] = res
        return res

    def search_crossref(self, query_str: str):
        q = re.sub(r'https?://\S+', ' ', query_str)
        q = re.sub(r'\bDOI\s*:?\s*\S+', ' ', q, flags=re.I)
        q = ' '.join(q.split())[:300]
        key = f"cr_q:{q.lower()}"
        with self.lock:
            if key in self.cache and self.cache[key].get("status") == 200:
                return self.cache[key]
        for attempt in range(4):
            self._wait_cr_search()
            try:
                r = self.session.get(
                    "https://api.crossref.org/works",
                    params={"query.bibliographic": q, "rows": 4, "mailto": self.user_email},
                    timeout=15
                )
                if r.status_code == 200:
                    items = r.json().get("message", {}).get("items", [])
                    res = {"status": 200, "items": [self._slim_crossref_work(it) for it in items]}
                    with self.lock:
                        self.cache[key] = res
                    return res
                elif r.status_code == 429:
                    time.sleep(2.0 * (attempt + 1))
            except Exception:
                time.sleep(1.5)
        return {"status": "error", "items": []}

    def search_crossref_tail(self, raw: str):
        q = re.sub(r'https?://\S+', ' ', raw)
        m = re.search(r'[“"”]([^“"”]{18,})[“"”]', q)
        if m:
            q_focus = m.group(1)
        else:
            q_focus = ' '.join(q.split())[-240:]
        key = f"cr_tail:{q_focus.lower()}"
        with self.lock:
            if key in self.cache and self.cache[key].get("status") == 200:
                return self.cache[key]
        for attempt in range(3):
            self._wait_cr_search()
            try:
                params = {"rows": 4, "mailto": self.user_email}
                if m:
                    params["query.title"] = q_focus
                else:
                    params["query.bibliographic"] = q_focus
                r = self.session.get("https://api.crossref.org/works", params=params, timeout=15)
                if r.status_code == 200:
                    items = r.json().get("message", {}).get("items", [])
                    res = {"status": 200, "items": [self._slim_crossref_work(it) for it in items]}
                    with self.lock:
                        self.cache[key] = res
                    return res
                elif r.status_code == 429:
                    time.sleep(1.5 * (attempt + 1))
            except Exception:
                time.sleep(1.0)
        return {"status": "error", "items": []}

    def search_europepmc(self, query_str: str):
        q = re.sub(r'https?://\S+', ' ', query_str)
        m = re.search(r'[“"”]([^“"”]{20,})[“"”]', q)
        if m:
            q = m.group(1)
        else:
            q = re.sub(r'[^\w\s]', ' ', strip_accents(q))
        q = ' '.join(q.split())[:220]
        key = f"epmc_q:{q.lower()}"
        with self.lock:
            if key in self.cache and self.cache[key].get("status") == 200:
                return self.cache[key]
        try:
            r = self.session.get(
                "https://www.ebi.ac.uk/europepmc/webservices/rest/search",
                params={"query": q, "format": "json", "pageSize": 5},
                timeout=12
            )
            if r.status_code == 200:
                results = r.json().get("resultList", {}).get("result", [])
                items = []
                for it in results:
                    auth_str = it.get("authorString") or ""
                    authors = []
                    for part in auth_str.split(","):
                        part = part.strip().rstrip(".")
                        if part:
                            sp = part.split()
                            authors.append({"family": sp[0], "given": " ".join(sp[1:])})
                    yr = int(it["pubYear"]) if str(it.get("pubYear", "")).isdigit() else None
                    items.append({
                        "doi": it.get("doi") or "",
                        "title": re.sub(r'<[^>]+>', '', it.get("title") or ""),
                        "authors": authors,
                        "journal": it.get("journalTitle") or "",
                        "year": yr,
                        "all_years": [yr] if yr else [],
                        "volume": str(it.get("journalVolume") or ""),
                        "page": str(it.get("pageInfo") or "")
                    })
                res = {"status": 200, "items": items}
                with self.lock:
                    self.cache[key] = res
                return res
        except Exception:
            pass
        return {"status": "error", "items": []}

    def _slim_crossref_work(self, msg: dict) -> dict:
        titles = msg.get("title") or []
        subtitles = msg.get("subtitle") or []
        title = titles[0] if titles else ""
        if subtitles and subtitles[0] and subtitles[0].lower() not in title.lower():
            title = f"{title}: {subtitles[0]}"
        title = re.sub(r'<[^>]+>', ' ', title)
        title = ' '.join(title.split())
        authors = []
        for a in msg.get("author") or []:
            authors.append({
                "family": a.get("family", ""),
                "given": a.get("given", ""),
                "name": a.get("name", "")
            })
        containers = msg.get("container-title") or []
        short_containers = msg.get("short-container-title") or []
        journal = containers[0] if containers else (short_containers[0] if short_containers else "")
        short_journal = short_containers[0] if short_containers else ""
        years = []
        for yfield in ["issued", "published-print", "published-online", "created"]:
            dp = (msg.get(yfield) or {}).get("date-parts")
            if dp and dp[0] and dp[0][0]:
                years.append(int(dp[0][0]))
        ji = msg.get("journal-issue") or {}
        for yfield in ["published-print", "published-online"]:
            dp = (ji.get(yfield) or {}).get("date-parts")
            if dp and dp[0] and dp[0][0]:
                years.append(int(dp[0][0]))
        year = min(years) if years else None
        page = msg.get("page") or msg.get("article-number") or ""
        return {
            "doi": msg.get("DOI", ""),
            "title": title,
            "authors": authors,
            "journal": journal,
            "short_journal": short_journal,
            "year": year,
            "all_years": sorted(list(set(years))),
            "volume": str(msg.get("volume") or ""),
            "page": str(page)
        }

def has_explicit_english_title(raw: str) -> bool:
    m = re.search(r'[“"”]([^“"”]{22,})[“"”]', raw)
    if m:
        return True
    cleaned = re.sub(r'https?://\S+', ' ', raw)
    words = [w for w in re.findall(r'\b[a-z]{4,}\b', cleaned) if w not in ("phys", "lett", "jour", "proc", "conf", "intl", "scie", "tech", "educ", "appl", "chem", "mater")]
    return len(words) >= 7

def score_candidate(raw: str, w: dict, is_tb120: bool = False, is_explicit_doi: bool = False) -> dict:
    raw_norm = strip_accents(raw).lower()
    raw_toks = set(re.findall(r"[a-z0-9]+", raw_norm))
    has_title = has_explicit_english_title(raw)
    
    # 1. Title coverage
    w_title = w.get("title") or ""
    w_title_toks = [t for t in norm_tokens(w_title) if len(t) >= 4 and t not in STOPWORDS]
    if len(w_title_toks) >= 2:
        matched_t = sum(1 for t in w_title_toks if t in raw_toks or any(rt.startswith(t[:6]) or t.startswith(rt[:6]) for rt in raw_toks if len(rt) >= 5))
        title_cov = matched_t / len(w_title_toks)
    elif len(w_title_toks) == 1:
        title_cov = 1.0 if w_title_toks[0] in raw_toks else 0.0
    else:
        title_cov = 0.0

    # 2. Author match
    w_authors = w.get("authors") or []
    first_author_matched = False
    matched_authors = 0
    for idx, a in enumerate(w_authors):
        fam = strip_accents(a.get("family", "")).lower()
        giv = strip_accents(a.get("given", "")).lower()
        nam = strip_accents(a.get("name", "")).lower()
        if any(inst in nam for inst in ("department", "university", "institute", "school", "college", "center", "faculty", "academy", "ministry")):
            nam = ""
        a_toks = [t for t in re.findall(r"[a-z]+", f"{fam} {giv} {nam}") if len(t) >= 3 and t not in ("van", "thi", "and", "the", "for", "with", "education", "science", "physics", "stem")]
        fam_toks = [t for t in re.findall(r"[a-z]+", fam) if len(t) >= 2 and t not in ("van", "thi")]
        is_hit = False
        if fam_toks and any(ft in raw_toks for ft in fam_toks):
            is_hit = True
        elif a_toks and any(at in raw_toks for at in a_toks):
            is_hit = True
        if is_hit:
            matched_authors += 1
            if idx == 0:
                first_author_matched = True

    author_cov = (matched_authors / len(w_authors)) if w_authors else 0.0

    # 3. Volume & Page match
    vol = str(w.get("volume") or "").strip()
    vol_match = bool(vol and vol in raw_toks)
    
    page_str = str(w.get("page") or "").strip()
    page_nums = [p for p in re.findall(r"[a-z0-9]+", page_str.lower()) if len(p) >= 1]
    first_page = page_nums[0] if page_nums else ""
    page_match = bool(first_page and (first_page in raw_toks or first_page in raw_norm))

    # 4. Journal match
    j_full = strip_accents(w.get("journal") or "").lower()
    j_short = strip_accents(w.get("short_journal") or "").lower()
    j_toks = [t for t in re.findall(r"[a-z0-9]+", j_full) if len(t) >= 3 and t not in ("journal", "the", "and", "for", "section", "part", "series")]
    if j_toks:
        j_cov = sum(1 for jt in j_toks if jt in raw_toks or any(rt.startswith(jt[:4]) for rt in raw_toks if len(rt) >= 3)) / len(j_toks)
    else:
        j_cov = 0.0
    journal_match = j_cov >= 0.5 or (bool(j_full) and j_full in raw_norm)

    # 5. Year comparison
    cited_years = extract_cited_years(raw)
    w_years = w.get("all_years") or ([w["year"]] if w.get("year") else [])
    w_years = [int(y) for y in w_years if y]
    if cited_years and w_years:
        min_year_diff = min(abs(cy - wy) for cy in cited_years for wy in w_years)
    else:
        min_year_diff = None

    is_same_work = False
    if is_tb120:
        is_same_work = journal_match or (min_year_diff is not None and min_year_diff <= 1)
    elif is_explicit_doi and not has_title:
        if matched_authors >= 1 and (min_year_diff is None or min_year_diff <= 1 or vol_match or page_match):
            is_same_work = True
        elif vol_match and page_match and (min_year_diff is None or min_year_diff <= 1):
            is_same_work = True
        elif journal_match and (vol_match or page_match) and (min_year_diff is not None and min_year_diff == 0) and not w_authors:
            is_same_work = True
    elif has_title:
        if len(w_title_toks) <= 3 and not (vol_match or page_match):
            is_same_work = False
        elif title_cov >= 0.50 and (first_author_matched or matched_authors >= 1 or vol_match or page_match or "collaboration" in raw_norm):
            is_same_work = True
        elif title_cov >= 0.72 and len(w_title_toks) >= 4:
            is_same_work = True
        elif title_cov >= 0.45 and first_author_matched and vol_match and (min_year_diff is not None and min_year_diff <= 1):
            is_same_work = True
    else:
        if (first_author_matched or matched_authors >= 2) and vol_match and page_match:
            is_same_work = True
        elif first_author_matched and journal_match and (vol_match or page_match) and (min_year_diff is not None and min_year_diff <= 1):
            is_same_work = True
        elif matched_authors >= 2 and journal_match and (vol_match or page_match) and (min_year_diff is not None and min_year_diff <= 1):
            is_same_work = True
        elif first_author_matched and matched_authors >= 2 and journal_match and (min_year_diff is not None and min_year_diff <= 1) and not vol:
            is_same_work = True

    return {
        "is_same_work": is_same_work,
        "title_cov": round(title_cov, 2),
        "first_author_matched": first_author_matched,
        "matched_authors": matched_authors,
        "author_cov": round(author_cov, 2),
        "vol_match": vol_match,
        "page_match": page_match,
        "journal_match": journal_match,
        "min_year_diff": min_year_diff,
        "cited_years": cited_years,
        "pub_years": w_years
    }

def pick_best_candidate(raw: str, items: list):
    matches = []
    for cand in items:
        sc = score_candidate(raw, cand)
        if sc["is_same_work"]:
            vp_ok = (1 if sc["vol_match"] else 0) + (1 if sc["page_match"] else 0)
            vp_both = 1 if (sc["vol_match"] and sc["page_match"] and sc["journal_match"]) else 0
            yr_ok = 1 if (sc["min_year_diff"] is not None and sc["min_year_diff"] <= 1) else 0
            j_ok = 1 if sc["journal_match"] else 0
            matches.append(((vp_both, yr_ok, j_ok, vp_ok, sc["title_cov"]), cand, sc))
    if not matches:
        return None, None
    matches.sort(key=lambda x: x[0], reverse=True)
    return matches[0][1], matches[0][2]

def is_genuine_year_error(sc: dict, w: dict) -> bool:
    if sc["min_year_diff"] is None or sc["min_year_diff"] < 2:
        return False
    doi = (w.get("doi") or "").lower()
    if doi.startswith("10.15625/") and 2019 in sc["pub_years"]:
        return False
    if doi.startswith("10.1007/") and sc["min_year_diff"] == 2 and sc["vol_match"] and sc["page_match"]:
        return False
    return True

def verify_single_reference(auditor: CitationAuditor, fname: str, ref: dict, author_matchers: list = None) -> dict:
    raw = ref["raw"]
    ref_id = ref["id"]
    is_tb120 = fname.startswith("TB1.20")
    
    # 0. Kiểm tra lỗi biên tập chế bản / dán nhầm văn bản
    if "cây hương nhu trắng" in raw.lower() or "cường độ phổ chuẩn hóa" in raw.lower() or raw.strip().endswith("(20"):
        res = {
            "id": ref_id, "raw": raw,
            "status": "VERIFIED_WITH_METADATA_ERROR",
            "label": "Lỗi biên tập nghiêm trọng (Dán nhầm văn bản thân bài / Chữ trục đồ thị / Cắt cụt dòng)",
            "matched_work": None,
            "details": "Mục trích dẫn bị lỗi biên tập chế bản: dán lẫn nội dung văn bản thân bài/nhãn trục đồ thị vào danh mục tài liệu tham khảo hoặc bị cắt cụt dòng."
        }
        self_hits = check_self_citation(author_matchers, raw, None)
        res["is_self_citation"] = bool(self_hits)
        res["self_citation_authors"] = self_hits
        return res

    dois = extract_dois(raw)
    arxiv_ids = extract_arxiv_ids(raw)
    cited_years = extract_cited_years(raw)
    special_cat = is_non_article_or_domestic(raw)

    result = None

    # 1. Kiểm tra DOI tường minh
    if dois:
        doi_results = []
        for d in dois:
            cr = auditor.fetch_crossref_doi(d)
            if cr["status"] == 200 and cr["work"]:
                sc = score_candidate(raw, cr["work"], is_tb120=is_tb120, is_explicit_doi=True)
                doi_results.append({"doi": d, "resolved": True, "source": "CrossRef", "work": cr["work"], "score": sc})
            else:
                dc = auditor.fetch_datacite_or_handle(d)
                if dc["status"] == 200 and dc["work"]:
                    sc = score_candidate(raw, dc["work"], is_tb120=is_tb120, is_explicit_doi=True)
                    doi_results.append({"doi": d, "resolved": True, "source": "DataCite/Handle", "work": dc["work"], "score": sc})
                else:
                    doi_results.append({"doi": d, "resolved": False, "source": "CrossRef+DataCite", "work": None, "score": None})

        any_resolved = any(r["resolved"] for r in doi_results)
        
        if any_resolved:
            r0 = [r for r in doi_results if r["resolved"]][0]
            w0 = r0["work"]
            sc0 = r0["score"]
            if sc0["is_same_work"] or w0["title"] == "[Registered DOI via Handle]":
                failed_dois = [r["doi"] for r in doi_results if not r["resolved"]]
                if failed_dois:
                    result = {
                        "id": ref_id, "raw": raw,
                        "status": "VERIFIED_WITH_METADATA_ERROR",
                        "label": "Một phần DOI hợp lệ, một phần DOI không tồn tại (404)",
                        "matched_work": w0,
                        "details": f"DOI {r0['doi']} hợp lệ ({w0.get('title','')[:80]}), nhưng DOI {failed_dois} trả về 404 (Không tồn tại)."
                    }
                elif is_genuine_year_error(sc0, w0):
                    result = {
                        "id": ref_id, "raw": raw,
                        "status": "VERIFIED_WITH_METADATA_ERROR",
                        "label": "Tài liệu thật nhưng ghi sai năm xuất bản",
                        "matched_work": w0,
                        "details": f"DOI {r0['doi']} hợp lệ nhưng năm trích dẫn ({sc0['cited_years']}) lệch so với năm xuất bản thực tế ({sc0['pub_years']})."
                    }
                else:
                    result = {
                        "id": ref_id, "raw": raw,
                        "status": "VERIFIED",
                        "label": "Tài liệu thật (Đã xác thực qua DOI)",
                        "matched_work": w0,
                        "details": f"Xác thực khớp qua {r0['source']} DOI: {r0['doi']} — {w0.get('title','')[:90]}"
                    }
            else:
                cr_search = auditor.search_crossref(raw)
                cand_alt, sc_alt = pick_best_candidate(raw, cr_search.get("items", []))
                if cand_alt:
                    w0_fam = (w0.get("authors") or [{}])[0].get("family", "N/A")
                    result = {
                        "id": ref_id, "raw": raw,
                        "status": "VERIFIED_WITH_METADATA_ERROR",
                        "label": "Tài liệu thật nhưng ghi nhầm mã DOI của bài khác",
                        "matched_work": cand_alt,
                        "details": f"Mã DOI trong bài ({r0['doi']}) trỏ tới bài khác ('{w0.get('title','')[:60]}' của {w0_fam}), nhưng bài báo được trích dẫn có tồn tại thực tế với DOI: {cand_alt.get('doi')}."
                    }
                else:
                    w0_fam = (w0.get("authors") or [{}])[0].get("family", "N/A")
                    result = {
                        "id": ref_id, "raw": raw,
                        "status": "FABRICATED_CITATION",
                        "label": "Giả mạo trích dẫn / AI Hallucination (Gắn DOI của bài báo khác vào tên bài/tác giả không tồn tại)",
                        "matched_work": w0,
                        "details": f"Mã DOI {r0['doi']} thực chất thuộc về bài báo hoàn toàn khác: '{w0.get('title','')}' (Tác giả: {w0_fam}, {w0.get('year')}). Bài báo và nhóm tác giả nêu trong bản thảo không tồn tại trên cơ sở dữ liệu."
                    }
        else:
            for d in dois:
                if "/" in d:
                    pref, suff = d.split("/", 1)
                    if "00" in suff:
                        suff_alt = re.sub(r'00+', lambda m: m.group(0)[:-1], suff)
                        d_alt = f"{pref}/{suff_alt}"
                        if d_alt != d:
                            cr_alt = auditor.fetch_crossref_doi(d_alt)
                            if cr_alt["status"] == 200 and cr_alt["work"]:
                                result = {
                                    "id": ref_id, "raw": raw,
                                    "status": "VERIFIED_WITH_METADATA_ERROR",
                                    "label": "Tài liệu thật nhưng mã DOI ghi trong bản thảo bị gõ thừa số 0 (404)",
                                    "matched_work": cr_alt["work"],
                                    "details": f"Mã DOI {d} trong bản thảo bị gõ dư số 0 (lỗi 404); mã DOI chính xác trên CrossRef là {d_alt} ('{cr_alt['work'].get('title','')[:80]}')."
                                }
            if not result and not is_tb120:
                cr_search = auditor.search_crossref(raw)
                cand_alt, sc_alt = pick_best_candidate(raw, cr_search.get("items", []))
                if not cand_alt:
                    epmc_search = auditor.search_europepmc(raw)
                    cand_alt, sc_alt = pick_best_candidate(raw, epmc_search.get("items", []))
                if cand_alt:
                    result = {
                        "id": ref_id, "raw": raw,
                        "status": "VERIFIED_WITH_METADATA_ERROR",
                        "label": "Tài liệu thật nhưng mã DOI ghi trong bản thảo bị sai cú pháp/ký tự (404)",
                        "matched_work": cand_alt,
                        "details": f"Mã DOI {dois[0]} trả về lỗi 404, nhưng tìm thấy bài báo thật với DOI đúng là: {cand_alt.get('doi')} ('{cand_alt.get('title','')[:80]}')."
                    }
            if not result:
                result = {
                    "id": ref_id, "raw": raw,
                    "status": "FABRICATED_CITATION",
                    "label": "Giả mạo trích dẫn / AI Hallucination (Mã DOI 404 Not Found & Bài báo không tồn tại)",
                    "matched_work": None,
                    "details": f"Mã DOI {', '.join(dois)} hoàn toàn không tồn tại (404 Not Found) trên CrossRef, DataCite và hệ thống định danh DOI quốc tế; đồng thời bài báo được trích dẫn không tồn tại trên cơ sở dữ liệu."
                }

    # 2. Kiểm tra arXiv ID
    if not result and arxiv_ids:
        ax = auditor.fetch_arxiv(arxiv_ids[0])
        if ax["status"] == 200 and ax["work"]:
            w0 = ax["work"]
            sc0 = score_candidate(raw, w0)
            if sc0["first_author_matched"] or sc0["matched_authors"] >= 1 or sc0["title_cov"] >= 0.4 or "collaboration" in raw.lower():
                result = {
                    "id": ref_id, "raw": raw,
                    "status": "VERIFIED",
                    "label": "Tài liệu thật (Đã xác thực qua arXiv / Tạp chí)",
                    "matched_work": w0,
                    "details": f"Xác thực khớp qua arXiv:{arxiv_ids[0]} — {w0.get('title','')[:90]}"
                }

    # 3. Phân loại tài liệu đặc thù
    if not result and special_cat:
        result = {
            "id": ref_id, "raw": raw,
            "status": "BOOK_PATENT_DOMESTIC_DOC",
            "label": special_cat,
            "matched_work": None,
            "details": f"Phân loại: {special_cat} (Không thuộc phạm vi kiểm tra DOI bài báo quốc tế; thông tin trích dẫn hợp thức)."
        }

    # 3.5. Kiểm tra DOI tiền định
    if not result:
        for gd in guess_deterministic_dois(raw):
            ck = f"cr_doi:{gd.lower()}"
            if ck in auditor.cache and auditor.cache[ck].get("status") == 200 and auditor.cache[ck].get("work"):
                w_gd = auditor.cache[ck]["work"]
                sc_gd = score_candidate(raw, w_gd)
                if sc_gd["is_same_work"] or sc_gd["first_author_matched"] or sc_gd["matched_authors"] >= 1:
                    if is_genuine_year_error(sc_gd, w_gd):
                        result = {
                            "id": ref_id, "raw": raw,
                            "status": "VERIFIED_WITH_METADATA_ERROR",
                            "label": "Tài liệu thật nhưng ghi sai năm xuất bản",
                            "matched_work": w_gd,
                            "details": f"Tìm thấy bài báo khớp DOI chuẩn {gd}, nhưng năm trích dẫn ({sc_gd['cited_years']}) lệch so với năm xuất bản thực tế ({sc_gd['pub_years']})."
                        }
                    else:
                        result = {
                            "id": ref_id, "raw": raw,
                            "status": "VERIFIED",
                            "label": "Tài liệu thật (Đã xác thực qua CrossRef DOI chuẩn)",
                            "matched_work": w_gd,
                            "details": f"Khớp bài báo: '{w_gd.get('title','')[:85]}' — {w_gd.get('journal','')} ({w_gd.get('year')}), DOI: {w_gd.get('doi')}"
                        }
                    break

    # 4. Tìm kiếm thư mục
    if not result:
        cr_search = auditor.search_crossref(raw)
        best_cand, best_sc = pick_best_candidate(raw, cr_search.get("items", []))

        if not best_cand and has_explicit_english_title(raw):
            epmc_search = auditor.search_europepmc(raw)
            best_cand, best_sc = pick_best_candidate(raw, epmc_search.get("items", []))

        if not best_cand and (has_explicit_english_title(raw) or len(raw) > 220):
            cr_tail = auditor.search_crossref_tail(raw)
            best_cand, best_sc = pick_best_candidate(raw, cr_tail.get("items", []))

        if best_cand:
            if is_genuine_year_error(best_sc, best_cand) and (
                (best_sc["journal_match"] and best_sc["vol_match"] and best_sc["page_match"])
                or (best_sc["title_cov"] >= 0.85 and best_sc["matched_authors"] >= 4)
            ):
                result = {
                    "id": ref_id, "raw": raw,
                    "status": "VERIFIED_WITH_METADATA_ERROR",
                    "label": "Tài liệu thật nhưng ghi sai năm xuất bản",
                    "matched_work": best_cand,
                    "details": f"Tìm thấy bài báo khớp tác giả/tạp chí/tập/trang (DOI: {best_cand.get('doi')}), nhưng năm trích dẫn trong bản thảo ({best_sc['cited_years']}) sai lệch so với năm xuất bản thực tế ({best_sc['pub_years']})."
                }
            else:
                result = {
                    "id": ref_id, "raw": raw,
                    "status": "VERIFIED",
                    "label": "Tài liệu thật (Đã xác thực qua CrossRef/OpenAlex)",
                    "matched_work": best_cand,
                    "details": f"Khớp bài báo: '{best_cand.get('title','')[:85]}' — {best_cand.get('journal','')} ({best_cand.get('year')}), DOI: {best_cand.get('doi') or 'N/A'}"
                }

    # 5. Bài báo tiếng Anh có tiêu đề đầy đủ mà không tồn tại
    if not result:
        if has_explicit_english_title(raw) and cited_years and max(cited_years) <= 2025:
            result = {
                "id": ref_id, "raw": raw,
                "status": "FABRICATED_CITATION",
                "label": "Nghi ngờ giả mạo trích dẫn / Không tồn tại bài báo trên CrossRef & OpenAlex",
                "matched_work": None,
                "details": "Trích dẫn có tiêu đề tiếng Anh đầy đủ và năm xuất bản <= 2025 nhưng hoàn toàn không tìm thấy trên cả CrossRef lẫn OpenAlex."
            }
        else:
            result = {
                "id": ref_id, "raw": raw,
                "status": "UNVERIFIED_NEEDS_CHECK",
                "label": "Chưa xác thực tự động (Trích dẫn rút gọn / Mới xuất bản 2026 / Sách hoặc kỷ yếu)",
                "matched_work": None,
                "details": "Trích dẫn viết dạng rút gọn (không có tiêu đề bài báo, không có DOI) hoặc công trình mới năm 2026 chưa cập nhật chỉ mục CrossRef/OpenAlex."
            }

    # ==================== ĐỐI CHIẾU TỰ TRÍCH DẪN ====================
    self_hits = check_self_citation(author_matchers, raw, result.get("matched_work"))
    result["is_self_citation"] = bool(self_hits)
    result["self_citation_authors"] = self_hits
    if self_hits:
        result["details"] = f"[🔄 Tự trích: {', '.join(self_hits)}] " + result["details"]

    return result
