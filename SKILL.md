---
name: pdf-citation-audit
description: Extracts, cross-checks, and verifies academic references directly from one or multiple scientific PDF manuscripts. Audits against CrossRef, DataCite, arXiv, and Europe PMC to detect AI-hallucinated citations, metadata/DOI errors, and author self-citations. Consolidates all findings into citation_check_result.json (machine-readable) and citation_check_result.xlsx (5-sheet interactive Excel workbook).
allowed-tools: Read Write Edit Bash WebSearch WebFetch
license: MIT License
compatibility: Python 3.9+ with pymupdf (fitz), requests, openpyxl. Runs on Windows, Linux, macOS.
metadata:
  version: "1.2"
  last-reviewed: "2026-10-23"
  skill-author: Antigravity & K-Dense Web Team
---

# PDF Citation Audit & Verification Skill

## 1. Overview

This skill provides an automated, end-to-end scholarly audit pipeline to:
1. **Accept Versatile Inputs**: Audit a **single PDF manuscript**, **multiple explicit PDF files**, or **recursively scan an entire folder** of scientific papers.
2. **Extract Complete Bibliographies**: Reliably parse reference sections from scientific papers, conference proceedings, theses, and journal articles.
3. **Cross-Reference Scholarly Databases**: Query CrossRef (Polite Pool queue), DataCite, Handle.net, arXiv API, and Europe PMC via persistent caching (`api_cache.json`).
4. **Detect AI Hallucinations & Fabricated Citations**: Identify papers invented by generative LLMs (e.g., ChatGPT) featuring non-existent titles, fabricated authors, fake (`404 Not Found`) DOIs, or **hijacked DOIs** belonging to completely unrelated publications.
5. **Diagnose Metadata & Publishing Discrepancies**: Pinpoint genuine works misattributed due to copy-paste DOI mixups, trailing zeroes, publication year discrepancies, or pasted body text.
6. **Analyze Author Self-Citations**: Automatically extract manuscript authors from Page 1, generate multi-variant name matchers (diacritics, reversed order, abbreviated initials), calculate self-citation counts and rates (`self_citation_rate`), and contextualize 10–20 year retrospective laboratory reviews versus isolated papers.
7. **Consolidate Results**: Instead of cluttering the workspace with fragmented subfolders for each paper, all audit data is consolidated into two core deliverables:
   - `citation_check_result.json`: Fully structured machine-readable dataset for automated pipelines and APIs.
   - `citation_check_result.xlsx`: 5-sheet styled Excel workbook with visual color-coding, AutoFilters, and frozen headers for editors and peer reviewers.

---

## 2. Multilingual Interaction Principle (Nguyên tắc Đa ngôn ngữ)

> [!IMPORTANT]
> **Adaptive User Language**: When an AI agent executes this skill or communicates with the user, **always respond in the exact language the user used in their prompt**.
> - **If the user interacts in Vietnamese**: Provide all explanations, analysis summaries, dialogue, and consultation in Vietnamese.
> - **If the user interacts in English**: Respond in English.
> - **If the user interacts in French, Japanese, etc.**: Adapt your response to match their language.
> - **Tool Output**: The underlying scripts default to English (`--lang en`) for international publication standards, but can produce Vietnamese reports if `--lang vi` is explicitly requested.

---

## 3. When to Use This Skill

- Journal editorial offices, conference program committees, or thesis evaluation panels verifying the authenticity of reference lists in scientific manuscripts.
- Investigating suspected AI-generated hallucinations in academic submissions.
- Evaluating author self-citation rates to ensure healthy peer engagement while properly recognizing legitimate retrospective research group reviews.
- Auditing large batches of PDF files (1 to 100+ manuscripts) rapidly without encountering API rate limits.
- Pre-publication galley proof checks to catch broken DOIs, wrong publication years, or publishing copy-paste errors.

---

## 4. Pipeline Architecture

```mermaid
flowchart TD
    A["Inputs: Single PDF / Multiple PDFs / Directory"] --> B["Stage 1: PDF Extraction & Normalization"]
    B --> B1["Unicode NFC Normalization"]
    B --> B2["Extract Manuscript Authors (Page 1)"]
    B --> B3["Filter Headers/Footers & Rejoin Broken DOIs"]
    B --> B4["Detect Reference Numbering: [n], n., APA, Bullet"]
    
    B4 --> C["Stage 2: Indexing & Persistent Cache"]
    C --> C1["Extract DOI / arXiv ID / Journal / Year"]
    C --> C2["Deterministic DOI Guessing (APS, IOP, OSA, Elsevier, Springer)"]
    C --> C3["Query CrossRef Polite Pool & DataCite"]
    
    C3 --> D["Stage 3: 5-Tier Classification & Verification"]
    D --> D1["✅ VERIFIED: Exact Match in International Database"]
    D --> D2["📘 BOOK_PATENT_DOMESTIC: Monograph / Patent / Standard / Regional"]
    D --> D3["⚠️ METADATA_ERROR: DOI Mixup / Year Discrepancy / Typo"]
    D --> D4["❌ FABRICATED_CITATION: AI Hallucination / Hijacked DOI"]
    D --> D5["🔍 UNVERIFIED: Abbreviated Citation / Awaiting Indexing"]
    
    B2 --> S["Author Self-Citation Analysis"]
    D --> S
    S --> S1["Match Citation String & Database Authors"]
    S --> S2["Calculate Self-Citation Rate (%) & Evaluate Context"]
    
    S2 --> E["Stage 4: Consolidated Deliverables"]
    E --> E1["citation_check_result.json (Machine-readable)"]
    E --> E2["citation_check_result.xlsx (5-Sheet Excel Workbook)"]
    E --> E3["citation_check_result.md (Markdown Executive Summary)"]
```

---

## 5. Key Technical Rules

### 5.1. Text Extraction & Author Identification
1. **Unicode Normalization:** Always apply `unicodedata.normalize('NFC', text)` immediately after extracting text via `pymupdf` (`fitz`).
2. **First-Page Author Extraction (`extract_manuscript_authors`):**
   - Strip conference/journal headers.
   - Discard title words matching the filename stem.
   - Filter out affiliation keywords (`department`, `university`, `faculty`, `institute`, `email`, `@`).
   - Strip academic titles and honorifics (`Prof.`, `Assoc. Prof.`, `Dr.`, `GS.`, `PGS.`, `TS.`).
   - Recognize collective authors (e.g., `Department of Nuclear Physics`).
3. **Safe Page-Number Filtering:** Only remove standalone digits `^\d{1,3}$` occurring in the first 2 or last 2 lines of a page. Never strip numbers in mid-text to avoid breaking DOIs with numeric line breaks.
4. **Rejoin Broken DOIs:** Apply regex to rejoin broken sequences such as `10.\n1103/` or `doi.org/10.xxxx/\n12345`.

### 5.2. Network Strategy & Rate Limit Mitigation
1. **CrossRef Polite Pool:** Always send a valid `mailto=your_email@domain.edu` parameter and descriptive User-Agent to enter CrossRef's Polite Pool (up to 10 req/s with prioritized reliability).
2. **Query Throttling:** Maintain a minimum delay of `0.25s - 0.28s` between text search calls (`/works?query.bibliographic=...`) using a thread lock.
3. **Persistent Cache (`api_cache.json`):** Save all HTTP 200 and 404 responses. Re-audits will execute in seconds by loading 100% of data from cache.

### 5.3. Fraud Detection (AI Hallucination vs. Metadata Error)
1. **Fabricated Citation (AI Hallucination - `FABRICATED_CITATION`):**
   - **Criterion 1:** DOI returns `404 Not Found` across CrossRef, DataCite, and Handle.net, AND the title/author query yields zero matching works.
   - **Criterion 2 (Hijacked DOI):** Manuscript cites a specific paper title, but the attached DOI actually belongs to an entirely unrelated paper (different authors, different field, different journal), and the cited paper does not exist.
2. **Genuine Work with DOI Mixup (`VERIFIED_WITH_METADATA_ERROR`):**
   - The manuscript DOI points to a different paper, BUT searching the cited title discovers the authentic paper in CrossRef/Europe PMC under a different, correct DOI (copy-paste error by authors).

### 5.4. Author Self-Citation Benchmark
1. **Matching Engine (`build_author_matchers`):** Match variations across full name, diacritics/non-diacritics, reversed order, and abbreviated initials (`P. D. Tam`, `Tam, P. D.`), checked against both the raw citation string and `matched_work["authors"]`.
2. **Evaluation Thresholds:**
   - 🟢 **$\le 25\%$ (Normal):** Natural succession of research; healthy academic practice.
   - ⚠️ **$25\% - 50\%$ (Substantial):** Noteworthy; editorial review should verify balance against external literature.
   - 🔄 **$> 50\%$ (Review Context):** Differentiate **Retrospective Laboratory/Group Reviews** (where citing prior group publications is legitimate, historical, and inherent to the paper) from standard research articles (which should avoid insular citation bubbles).

---

## 6. CLI Usage Guide

Scripts are located in `scripts/`:
- `batch_audit.py`: Central CLI runner supporting single/multiple PDFs and directory scans.
- `extract_references.py`: PDF text and author extraction.
- `verifier.py`: CrossRef/DataCite/arXiv/Europe PMC verification engine.
- `excel_reporter.py`: 5-sheet styled Excel workbook generator.

### Command Examples:

#### 1. Audit an entire directory of manuscripts:
```bash
python -X utf8 scripts/batch_audit.py "path/to/manuscripts_folder" --email "your_email@domain.edu"
```

#### 2. Audit a single PDF file:
```bash
python -X utf8 scripts/batch_audit.py "path/to/paper.pdf" --email "your_email@domain.edu"
```

#### 3. Audit multiple explicit PDF files:
```bash
python -X utf8 scripts/batch_audit.py "paper1.pdf" "paper2.pdf" "paper3.pdf" -o "output_folder"
```

#### CLI Options:
- `--lang {en,vi}`: Output language for reports and terminal output (default: `en`).
- `--email <email>`: Email for CrossRef Polite Pool priority queue.
- `--output-dir` / `-o`: Directory to save deliverables (defaults to input directory).
- `--output-prefix`: File prefix (default: `citation_check_result`).
- `--cache-file`: Explicit path to `api_cache.json`.
- `--workers`: Number of concurrent threads (default: `16`).
- `--create-folders`: Generate legacy `Report_<stem>` individual subfolders (default: `False`).

---

## 7. Deliverables Structure

Outputs are written to the target directory:
1. **`citation_check_result.json`**: Machine-readable JSON containing full metadata, summary metrics, flagged issues, and per-manuscript reference breakdowns.
2. **`citation_check_result.xlsx`**: 5-sheet interactive Excel workbook:
   - `Overview & Summary`: Executive KPIs, 5-tier status table, self-citation benchmarks.
   - `Manuscripts Summary`: Roster of all audited papers with self-citation metrics and compliance status.
   - `All References`: Complete flat table of all references with verification details, DOIs, and self-citation tags.
   - `Flagged Issues`: Filtered table of AI hallucinations and metadata discrepancies for editorial action.
   - `Author Self-Citations`: Filtered table of all self-citations.
3. **`citation_check_result.md`**: Executive summary formatted in GitHub Markdown.
