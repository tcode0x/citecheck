#!/usr/bin/env python3
"""
citecheck.py — Trình kiểm tra trích dẫn toàn diện cho bài báo khoa học.

Hỗ trợ đầu vào:
  - .docx  (Word)            -> tự trích danh mục tham khảo dạng [n] / n. và trích dẫn [n] trong văn bản
  - .tex + .bib (LaTeX)      -> tương thích chế độ cũ (\\cite{key} + @entry{key,...})
  - .pdf                     -> nếu cài pymupdf/pdfplumber (tùy chọn)
  - .txt / .md               -> văn bản thuần

Phát hiện:
  - Trích dẫn không có trong danh mục (undefined) và mục không được trích (unused)
  - Tài liệu TRÙNG LẶP (cùng DOI hoặc cùng tiêu đề nhưng khác số / khác key)
  - Marker gãy: [?], [??], ??, \\ref{} chưa giải quyết, "citation needed"
  - Số tham chiếu lặp / thiếu / lệch thứ tự
  - Sai khớp TÊN tác giả giữa "Author et al. [n]" trong văn bản và danh mục
  - Mục thiếu tiêu đề / năm / DOI-URL

Đối chiếu online (tùy chọn --api), nhiều nguồn, có cache + retry:
  - Crossref (DOI + bibliographic search), OpenAlex (DOI + search), arXiv (preprint)
  - Khớp mờ bằng Jaccard token + difflib + chồng lấp tác giả + gần năm
  - Luôn trả về ứng viên tốt nhất kèm link để rà tay

Xuất:
  - <out>.json  : báo cáo máy đọc
  - <out>.md    : báo cáo người đọc, có bảng đầy đủ để manual-check
  - <out>.csv   : bảng từng tham chiếu (mở Excel kiểm tra nhanh)

Ví dụ:
  python citecheck.py --docx paper.docx --api --mailto you@example.com --out report
  python citecheck.py --tex main.tex intro.tex --bib refs.bib --api --out report
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import time
import urllib.parse
from dataclasses import dataclass, field, asdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional

try:
    import requests
except ImportError:
    requests = None


# --------------------------------------------------------------------------- #
# Regex & hằng số
# --------------------------------------------------------------------------- #

CITE_RE = re.compile(
    r"\\(?:cite|citep|citet|citealp|citeauthor|citeyear|citeyearpar|parencite"
    r"|textcite|autocite|nocite)(?:\[[^\]]*\]){0,2}\{([^}]*)\}"
)
REF_RE = re.compile(r"\\(?:ref|eqref|autoref|cref|Cref)\{([^}]*)\}")
LABEL_RE = re.compile(r"\\label\{([^}]*)\}")
BIB_ENTRY_RE = re.compile(r"@\w+\s*\{\s*([^,\s]+)\s*,", re.I)

DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", re.I)
ARXIV_RE = re.compile(r"arXiv\s*:?\s*(\d{4}\.\d{4,5})(?:v\d+)?", re.I)
URL_RE = re.compile(r"https?://[^\s)>\]]+", re.I)
YEAR_RE = re.compile(r"(?:19|20)\d{2}")
QUOTE_RE = re.compile(r'["“”]([^"“”]{6,}?)["“”]')

# Marker gãy hay gặp khi build LaTeX lỗi hoặc cross-ref hỏng
BROKEN_RE = re.compile(r"\[\s*\?{1,2}\s*\]|\?\?+|\[\s*citation needed\s*\]", re.I)

# Trích dẫn số trong văn bản: [12], [1, 2], [1-3], [1–3], [1, 4, 9]
INTEXT_NUM_RE = re.compile(r"\[\s*(\d{1,3}(?:\s*[–—-]\s*\d{1,3})?(?:\s*,\s*\d{1,3}(?:\s*[–—-]\s*\d{1,3})?)*)\s*\]")

# "Surname et al. [n]" / "Surname and Other [n]"  (yêu cầu et al./and -> tránh nhầm tên method)
AUTHOR_LINK_RE = re.compile(
    r"([A-ZÀ-Ý][A-Za-zÀ-ÿ'’\-]+)"
    r"\s+(?:et\s+al\.?|and\s+[A-ZÀ-Ý][A-Za-zÀ-ÿ'’\-]+)"
    r"\s*\[\s*(\d{1,3})\s*\]"
)

REF_HEADING_RE = re.compile(
    r"^\s*(?:\d+\s*\.?\s*)?(references|bibliography|works cited|"
    r"tài liệu tham khảo|tài liệu|trích dẫn|tham khảo)\s*$",
    re.I,
)

ENTRY_START_RE = re.compile(r"^\s*(?:\[(\d{1,3})\]|(\d{1,3})\.)\s+(.*)$")

DEFAULT_HEADERS = {"User-Agent": "citecheck/2.0 (mailto:unknown@example.com)"}


# --------------------------------------------------------------------------- #
# Mô hình dữ liệu
# --------------------------------------------------------------------------- #

@dataclass
class Reference:
    ident: str                      # "[24]" hoặc bib key
    number: Optional[int] = None    # số thứ tự nếu là dạng numeric
    raw: str = ""
    authors: str = ""
    surnames: list = field(default_factory=list)
    title: str = ""
    year: str = ""
    journal: str = ""
    doi: str = ""
    arxiv: str = ""
    url: str = ""
    cited: bool = False
    cite_count: int = 0
    duplicate_of: Optional[str] = None
    problems: list = field(default_factory=list)
    api: dict = field(default_factory=dict)


# --------------------------------------------------------------------------- #
# Đọc đầu vào
# --------------------------------------------------------------------------- #

def read_text_file(path: str) -> str:
    return Path(path).read_text(encoding="utf-8", errors="ignore")


def read_docx(path: str) -> str:
    """Trích toàn bộ text từ .docx, gồm cả đoạn đánh số tự động và bảng."""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError:
        sys.exit("Cần cài python-docx:  pip install python-docx")

    doc = Document(path)
    lines = []

    def para_text(p):
        txt = p.text or ""
        # Nếu đoạn nằm trong danh sách đánh số tự động mà text không có "[n]"/"n.",
        # cố gắng phục hồi số từ thuộc tính numbering để không mất mục tham khảo.
        try:
            numpr = p._p.pPr is not None and p._p.pPr.numPr is not None
        except Exception:
            numpr = False
        if numpr and not re.match(r"^\s*(\[\d+\]|\d+\.)", txt):
            return ("§LIST§ " + txt).strip()
        return txt

    for block in _iter_block_items(doc):
        if block[0] == "p":
            lines.append(para_text(block[1]))
        else:  # table
            for row in block[1].rows:
                cells = [c.text.strip() for c in row.cells]
                lines.append("\t".join(cells))
    return "\n".join(lines)


def _iter_block_items(doc):
    """Duyệt paragraph và table theo đúng thứ tự xuất hiện."""
    from docx.document import Document as _Doc
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    parent_elm = doc.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield ("p", Paragraph(child, doc))
        elif isinstance(child, CT_Tbl):
            yield ("tbl", Table(child, doc))


def read_pdf(path: str) -> str:
    """Đọc PDF nếu có pymupdf hoặc pdfplumber (tùy chọn)."""
    try:
        import fitz  # PyMuPDF
        return "\n".join(page.get_text() for page in fitz.open(path))
    except Exception:
        pass
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return "\n".join(pg.extract_text() or "" for pg in pdf.pages)
    except Exception:
        sys.exit("Để đọc PDF cần:  pip install pymupdf   (hoặc pdfplumber)")


def load_document_text(paths: list) -> str:
    chunks = []
    for p in paths:
        ext = Path(p).suffix.lower()
        if ext == ".docx":
            chunks.append(read_docx(p))
        elif ext == ".pdf":
            chunks.append(read_pdf(p))
        else:
            chunks.append(read_text_file(p))
    return "\n".join(chunks)


# --------------------------------------------------------------------------- #
# Phân tích danh mục tham khảo (docx / text / pdf)
# --------------------------------------------------------------------------- #

def split_body_and_refs(text: str):
    """Tách (thân bài, vùng tham khảo) theo tiêu đề cuối cùng. Tránh đếm
    các marker [n] ở đầu mỗi mục tham khảo như thể là trích dẫn trong văn bản."""
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if REF_HEADING_RE.match(ln.strip()):
            start = i
    if start is None:
        return text, text
    return "\n".join(lines[:start]), "\n".join(lines[start + 1:])


def split_reference_section(text: str) -> str:
    """Trả về phần văn bản từ tiêu đề 'References/Tài liệu...' tới cuối."""
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if REF_HEADING_RE.match(ln.strip()):
            start = i + 1  # lấy lần xuất hiện cuối -> tránh mục lục
    if start is None:
        return text  # fallback: quét toàn văn
    return "\n".join(lines[start:])


def parse_reference_entries(text: str) -> list:
    """Tách danh mục thành từng mục [n] ... hoặc n. ... (mục có thể nhiều dòng)."""
    region = split_reference_section(text)
    lines = region.splitlines()

    entries = []  # (number_or_None, buffer)
    cur_num = None
    cur_buf = []
    auto_idx = 0

    def flush():
        if cur_buf:
            entries.append((cur_num, " ".join(s.strip() for s in cur_buf).strip()))

    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        m = ENTRY_START_RE.match(s)
        if m:
            flush()
            cur_buf = []
            num = m.group(1) or m.group(2)
            cur_num = int(num)
            cur_buf.append(m.group(3))
        elif s.startswith("§LIST§"):
            # đoạn danh sách tự động: coi như mục mới, tự đánh số
            flush()
            cur_buf = []
            auto_idx += 1
            cur_num = None
            cur_buf.append(s.replace("§LIST§", "").strip())
        else:
            if cur_buf:
                cur_buf.append(s)
    flush()

    refs = []
    seq = 0
    for num, raw in entries:
        if len(raw) < 12:
            continue
        seq += 1
        n = num if num is not None else seq
        refs.append(build_reference(f"[{n}]", n, raw))
    return refs


def build_reference(ident: str, number: Optional[int], raw: str) -> Reference:
    doi = ""
    m = DOI_RE.search(raw)
    if m:
        doi = m.group(0).rstrip(".,;")
    arxiv = ""
    a = ARXIV_RE.search(raw)
    if a:
        arxiv = a.group(1)
    url = ""
    u = URL_RE.search(raw)
    if u:
        url = u.group(0).rstrip(".,);")

    title = ""
    qm = QUOTE_RE.search(raw)
    if qm:
        title = qm.group(1).strip().rstrip(",.")
    authors = ""
    if qm:
        authors = raw[:qm.start()].strip().rstrip(",")
    # nếu không có ngoặc kép, đoán: phần trước năm đầu tiên
    if not title:
        ym = YEAR_RE.search(raw)
        guess = raw[:ym.start()] if ym else raw
        # bỏ phần tác giả nếu có dấu phẩy nhiều
        title = guess.strip().rstrip(",.")

    year = ""
    years = YEAR_RE.findall(raw)
    if years:
        year = years[-1]  # năm xuất bản thường ở cuối

    journal = guess_journal(raw, title)

    ref = Reference(
        ident=ident, number=number, raw=raw,
        authors=authors, title=title, year=year,
        journal=journal, doi=doi, arxiv=arxiv, url=url,
    )
    ref.surnames = extract_surnames(authors)
    # cảnh báo thiếu trường
    if not ref.title:
        ref.problems.append("missing_title")
    if not ref.year:
        ref.problems.append("missing_year")
    if not ref.doi and not ref.url and not ref.arxiv:
        ref.problems.append("missing_doi_or_url")
    return ref


def guess_journal(raw: str, title: str) -> str:
    after = raw
    if title and title in raw:
        after = raw.split(title, 1)[-1]
    after = after.lstrip(' ,."”')
    # lấy cụm tới dấu phẩy/chấm hoặc "vol"
    m = re.split(r",|\bvol\b|\bpp\b|\(", after, maxsplit=1)
    j = m[0].strip(' ,."”') if m else ""
    return j[:120]


def extract_surnames(authors: str) -> list:
    if not authors:
        return []
    authors = re.sub(r"\bet al\.?", "", authors, flags=re.I)
    chunks = re.split(r",|\band\b|;|&", authors)
    surs = []
    for c in chunks:
        c = c.strip()
        if not c:
            continue
        words = [w for w in re.split(r"\s+", c) if w]
        # bỏ initial dạng "A." -> lấy từ alphabet dài nhất
        cand = [w for w in words if re.match(r"^[A-ZÀ-Ý][a-zà-ÿ'’\-]{1,}$", w)]
        if cand:
            surs.append(cand[-1].lower())
    return surs


# --------------------------------------------------------------------------- #
# Phân tích BibTeX (chế độ LaTeX)
# --------------------------------------------------------------------------- #

def parse_bib(bib_text: str) -> list:
    refs = []
    starts = list(BIB_ENTRY_RE.finditer(bib_text))
    seen = {}
    for i, m in enumerate(starts):
        key = m.group(1).strip()
        start = m.start()
        end = starts[i + 1].start() if i + 1 < len(starts) else len(bib_text)
        body = bib_text[start:end]
        ref = Reference(
            ident=key, number=None, raw=body.strip(),
            title=bib_field(body, "title"),
            doi=bib_field(body, "doi"),
            year=bib_field(body, "year"),
            journal=bib_field(body, "journal") or bib_field(body, "booktitle"),
            authors=bib_field(body, "author"),
            url=bib_field(body, "url"),
        )
        a = ARXIV_RE.search(body)
        if a:
            ref.arxiv = a.group(1)
        ref.surnames = extract_surnames(ref.authors)
        if not ref.title:
            ref.problems.append("missing_title")
        if not ref.year:
            ref.problems.append("missing_year")
        if not ref.doi and not ref.url and not ref.arxiv:
            ref.problems.append("missing_doi_or_url")
        seen[key] = seen.get(key, 0) + 1
        refs.append(ref)
    for r in refs:
        if seen.get(r.ident, 0) > 1:
            r.problems.append("duplicate_bib_key")
    return refs


def bib_field(entry: str, name: str) -> str:
    pat = re.compile(rf"{name}\s*=\s*[{{\"](.+?)[}}\"]\s*,?\s*(?:\n|$)", re.I | re.S)
    m = pat.search(entry)
    if not m:
        return ""
    val = re.sub(r"\s+", " ", m.group(1)).strip()
    return val.replace("{", "").replace("}", "")


# --------------------------------------------------------------------------- #
# Trích dẫn trong văn bản
# --------------------------------------------------------------------------- #

def expand_range(token: str) -> list:
    token = re.sub(r"[–—]", "-", token)
    if "-" in token:
        a, b = token.split("-", 1)
        try:
            a, b = int(a), int(b)
            if 0 < b - a < 200:
                return list(range(a, b + 1))
        except ValueError:
            return []
    try:
        return [int(token)]
    except ValueError:
        return []


def extract_numeric_intext(text: str) -> dict:
    counts = {}
    for m in INTEXT_NUM_RE.finditer(text):
        for tok in m.group(1).split(","):
            for n in expand_range(tok.strip()):
                counts[n] = counts.get(n, 0) + 1
    return counts


def extract_author_links(text: str) -> list:
    out = []
    for m in AUTHOR_LINK_RE.finditer(text):
        surname = m.group(1).lower()
        num = int(m.group(2))
        ctx = text[max(0, m.start() - 40): m.end() + 40].replace("\n", " ")
        out.append({"surname": surname, "number": num, "context": ctx.strip()})
    return out


def extract_broken_markers(text: str) -> list:
    out = []
    for m in BROKEN_RE.finditer(text):
        ctx = text[max(0, m.start() - 50): m.end() + 50].replace("\n", " ")
        out.append({"marker": m.group(0), "context": ctx.strip()})
    # \ref{} chưa biên dịch trong text thuần
    for m in REF_RE.finditer(text):
        out.append({"marker": f"\\ref{{{m.group(1)}}}", "context": ""})
    return out


# --------------------------------------------------------------------------- #
# Phát hiện trùng lặp
# --------------------------------------------------------------------------- #

def normalize_title(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def title_similarity(a: str, b: str) -> float:
    A, B = set(normalize_title(a).split()), set(normalize_title(b).split())
    if not A or not B:
        return 0.0
    jacc = len(A & B) / len(A | B)
    ratio = SequenceMatcher(None, normalize_title(a), normalize_title(b)).ratio()
    return round(0.5 * jacc + 0.5 * ratio, 3)


def detect_duplicates(refs: list, thr: float = 0.85) -> list:
    """Nhóm các mục cùng DOI hoặc tiêu đề rất giống nhau (khác số/khác key)."""
    groups = []
    used = set()
    for i in range(len(refs)):
        if i in used:
            continue
        gi = [i]
        for j in range(i + 1, len(refs)):
            if j in used:
                continue
            same_doi = (refs[i].doi and refs[j].doi
                        and refs[i].doi.lower() == refs[j].doi.lower())
            sim = title_similarity(refs[i].title, refs[j].title)
            if same_doi or sim >= thr:
                gi.append(j)
        if len(gi) > 1:
            for k in gi:
                used.add(k)
            members = [refs[k].ident for k in gi]
            for k in gi[1:]:
                refs[k].duplicate_of = refs[gi[0]].ident
                refs[k].problems.append("duplicate_reference")
            groups.append({
                "members": members,
                "title": refs[gi[0]].title,
                "doi": refs[gi[0]].doi,
            })
    return groups


# --------------------------------------------------------------------------- #
# Đối chiếu online — nhiều nguồn, có cache + retry
# --------------------------------------------------------------------------- #

class OnlineVerifier:
    def __init__(self, providers, mailto=None, cache_path=".citecache.json",
                 sleep=0.34, use_cache=True, timeout=20):
        if requests is None:
            sys.exit("Cần cài requests:  pip install requests")
        self.providers = providers
        self.mailto = mailto
        self.sleep = sleep
        self.timeout = timeout
        self.use_cache = use_cache
        self.cache_path = Path(cache_path)
        self.cache = {}
        if use_cache and self.cache_path.exists():
            try:
                self.cache = json.loads(self.cache_path.read_text("utf-8"))
            except Exception:
                self.cache = {}
        self.headers = dict(DEFAULT_HEADERS)
        if mailto:
            self.headers["User-Agent"] = f"citecheck/2.0 (mailto:{mailto})"

    def _save_cache(self):
        if self.use_cache:
            try:
                self.cache_path.write_text(
                    json.dumps(self.cache, ensure_ascii=False), "utf-8")
            except Exception:
                pass

    def _get(self, url, params=None):
        ck = url + "|" + urllib.parse.urlencode(params or {})
        if self.use_cache and ck in self.cache:
            return self.cache[ck]
        backoff = 1.0
        for attempt in range(4):
            try:
                r = requests.get(url, params=params, headers=self.headers,
                                 timeout=self.timeout)
                if r.status_code == 429:
                    wait = float(r.headers.get("Retry-After", backoff))
                    time.sleep(wait)
                    backoff *= 2
                    continue
                if r.status_code == 404:
                    data = {"_status": 404}
                    self.cache[ck] = data
                    return data
                r.raise_for_status()
                data = r.json()
                self.cache[ck] = data
                time.sleep(self.sleep)
                return data
            except Exception:
                time.sleep(backoff)
                backoff *= 2
        return None

    # ----- providers -----
    def crossref_doi(self, doi):
        d = self._get("https://api.crossref.org/works/" +
                      urllib.parse.quote(doi.strip()))
        if not d or d.get("_status") == 404 or "message" not in d:
            return None
        return self._cr_item(d["message"], "crossref-doi")

    def crossref_search(self, ref):
        q = " ".join(x for x in [ref.title, ref.authors] if x)[:300]
        d = self._get("https://api.crossref.org/works",
                      {"query.bibliographic": q, "rows": 3})
        if not d or "message" not in d:
            return None
        items = d["message"].get("items", [])
        return [self._cr_item(it, "crossref-search") for it in items]

    @staticmethod
    def _cr_item(it, source):
        title = (it.get("title") or [""])[0]
        year = ""
        for k in ("published-print", "published-online", "issued", "created"):
            if it.get(k, {}).get("date-parts"):
                year = str(it[k]["date-parts"][0][0])
                break
        authors = []
        for a in it.get("author", []) or []:
            fam = a.get("family")
            if fam:
                authors.append(fam.lower())
        return {"source": source, "title": title, "year": year,
                "doi": it.get("DOI", ""), "authors": authors,
                "url": ("https://doi.org/" + it["DOI"]) if it.get("DOI") else ""}

    def openalex_doi(self, doi):
        d = self._get(f"https://api.openalex.org/works/https://doi.org/{doi}",
                      {"mailto": self.mailto} if self.mailto else None)
        if not d or d.get("_status") == 404 or "id" not in d:
            return None
        return self._oa_item(d)

    def openalex_search(self, ref):
        params = {"search": ref.title[:300], "per-page": 3}
        if self.mailto:
            params["mailto"] = self.mailto
        d = self._get("https://api.openalex.org/works", params)
        if not d or "results" not in d:
            return None
        return [self._oa_item(it) for it in d["results"]]

    @staticmethod
    def _oa_item(it):
        authors = []
        for a in it.get("authorships", []) or []:
            nm = (a.get("author") or {}).get("display_name", "")
            if nm:
                authors.append(nm.split()[-1].lower())
        doi = (it.get("doi") or "").replace("https://doi.org/", "")
        return {"source": "openalex", "title": it.get("title") or "",
                "year": str(it.get("publication_year") or ""),
                "doi": doi, "authors": authors,
                "url": it.get("doi") or it.get("id", "")}

    def arxiv_search(self, ref):
        # arXiv trả Atom XML
        q = f'ti:"{ref.title[:200]}"' if ref.title else ""
        if ref.arxiv:
            q = f"id:{ref.arxiv}"
        if not q:
            return None
        ck = "arxiv|" + q
        if self.use_cache and ck in self.cache:
            xml = self.cache[ck]
        else:
            try:
                r = requests.get("http://export.arxiv.org/api/query",
                                 params={"search_query": q, "max_results": 3},
                                 headers=self.headers, timeout=self.timeout)
                xml = r.text
                self.cache[ck] = xml
                time.sleep(self.sleep)
            except Exception:
                return None
        import xml.etree.ElementTree as ET
        ns = {"a": "http://www.w3.org/2005/Atom"}
        out = []
        try:
            root = ET.fromstring(xml)
            for e in root.findall("a:entry", ns):
                title = (e.findtext("a:title", default="", namespaces=ns) or "").strip()
                pub = e.findtext("a:published", default="", namespaces=ns) or ""
                year = pub[:4]
                idtxt = e.findtext("a:id", default="", namespaces=ns) or ""
                authors = [au.findtext("a:name", default="", namespaces=ns).split()[-1].lower()
                           for au in e.findall("a:author", ns)
                           if au.findtext("a:name", default="", namespaces=ns)]
                out.append({"source": "arxiv", "title": title, "year": year,
                            "doi": "", "authors": authors, "url": idtxt})
        except Exception:
            return None
        return out

    # ----- chấm điểm & tổng hợp -----
    def score(self, ref, cand):
        sim = title_similarity(ref.title, cand.get("title", ""))
        s = sim
        ref_sur = set(ref.surnames)
        cand_sur = set(cand.get("authors", []))
        if ref_sur and cand_sur:
            overlap = len(ref_sur & cand_sur) / len(ref_sur | cand_sur)
            s += 0.15 * overlap
        if ref.year and cand.get("year"):
            try:
                if abs(int(ref.year) - int(cand["year"])) <= 1:
                    s += 0.05
            except ValueError:
                pass
        return round(min(s, 1.0), 3), sim

    def verify(self, ref: Reference):
        candidates = []
        # ưu tiên DOI (chính xác)
        if ref.doi:
            if "crossref" in self.providers:
                c = self.crossref_doi(ref.doi)
                if c:
                    candidates.append(c)
            if "openalex" in self.providers:
                c = self.openalex_doi(ref.doi)
                if c:
                    candidates.append(c)
        # tìm theo tiêu đề
        if ref.title:
            if "crossref" in self.providers:
                lst = self.crossref_search(ref) or []
                candidates.extend(lst)
            if "openalex" in self.providers:
                lst = self.openalex_search(ref) or []
                candidates.extend(lst)
        if "arxiv" in self.providers and (ref.arxiv or ref.title):
            lst = self.arxiv_search(ref) or []
            candidates.extend(lst)

        if not candidates:
            return {"status": "NOT_FOUND", "best": None, "score": 0,
                    "title_sim": 0, "candidates": []}

        scored = []
        for c in candidates:
            sc, sim = self.score(ref, c)
            scored.append({**c, "score": sc, "title_sim": sim})
        scored.sort(key=lambda x: x["score"], reverse=True)
        best = scored[0]

        status = "OK"
        if best["title_sim"] < 0.55:
            status = "TITLE_MISMATCH"
        elif ref.year and best.get("year") and ref.year != best["year"]:
            status = "YEAR_MISMATCH"
        elif best["score"] < 0.6:
            status = "LOW_CONFIDENCE"
        # DOI bài báo khác DOI tra được
        if (ref.doi and best.get("doi")
                and ref.doi.lower() != best["doi"].lower()
                and best["title_sim"] > 0.8):
            status = "DOI_MISMATCH"

        return {"status": status, "best": best, "score": best["score"],
                "title_sim": best["title_sim"], "candidates": scored[:3]}

    def run(self, refs, limit=None):
        targets = [r for r in refs if not r.duplicate_of]  # khỏi tra bản trùng
        if limit:
            targets = targets[:limit]
        for i, r in enumerate(targets, 1):
            r.api = self.verify(r)
            print(f"  [{i}/{len(targets)}] {r.ident} -> {r.api['status']} "
                  f"(score {r.api['score']})", file=sys.stderr)
            self._save_cache()
        self._save_cache()


# --------------------------------------------------------------------------- #
# Đối chiếu tên tác giả trong văn bản với danh mục
# --------------------------------------------------------------------------- #

def check_author_links(links, num2ref):
    issues = []
    for lk in links:
        ref = num2ref.get(lk["number"])
        if not ref:
            issues.append({**lk, "issue": "number_not_in_references",
                           "ref_authors": ""})
            continue
        if not ref.surnames:
            continue
        if lk["surname"] not in ref.surnames:
            issues.append({
                "surname": lk["surname"], "number": lk["number"],
                "context": lk["context"], "issue": "author_name_mismatch",
                "ref_authors": ref.authors,
            })
    return issues


# --------------------------------------------------------------------------- #
# Báo cáo
# --------------------------------------------------------------------------- #

def build_report(refs, intext_counts, author_links, broken, dup_groups,
                 author_issues, mode):
    num2ref = {r.number: r for r in refs if r.number is not None}

    cited_numbers = set(intext_counts.keys())
    ref_numbers = set(num2ref.keys())

    for r in refs:
        if r.number is not None:
            r.cite_count = intext_counts.get(r.number, 0)
            r.cited = r.cite_count > 0

    undefined = sorted(cited_numbers - ref_numbers)        # số trích nhưng không có mục
    unused = sorted(r.ident for r in refs if r.number is not None and not r.cited)

    nums = sorted(ref_numbers)
    gaps = []
    if nums:
        full = set(range(min(nums), max(nums) + 1))
        gaps = sorted(full - set(nums))

    report = {
        "summary": {
            "mode": mode,
            "num_references": len(refs),
            "num_unique_intext_numbers": len(cited_numbers),
            "num_intext_citation_tokens": sum(intext_counts.values()),
            "num_duplicate_groups": len(dup_groups),
            "num_undefined_citations": len(undefined),
            "num_unused_references": len(unused),
            "num_broken_markers": len(broken),
            "num_author_name_issues": len(author_issues),
            "num_refs_missing_fields": sum(1 for r in refs if r.problems),
        },
        "undefined_citations": undefined,
        "missing_numbers_in_list": gaps,
        "unused_references": unused,
        "duplicate_groups": dup_groups,
        "broken_markers": broken,
        "author_name_issues": author_issues,
        "references": [_ref_dict(r) for r in refs],
    }
    return report


def _ref_dict(r: Reference):
    d = asdict(r)
    # gọn cho JSON
    d["raw"] = (r.raw[:300] + "…") if len(r.raw) > 300 else r.raw
    return d


def write_markdown(report, refs, path):
    L = []
    s = report["summary"]
    L.append("# Báo cáo kiểm tra trích dẫn\n")
    L.append(f"- Chế độ: **{s['mode']}**")
    L.append(f"- Số mục tham khảo: **{s['num_references']}**")
    L.append(f"- Số trích dẫn trong văn bản (token): **{s['num_intext_citation_tokens']}** "
             f"(số hiệu duy nhất: {s['num_unique_intext_numbers']})")
    L.append(f"- Nhóm tài liệu TRÙNG LẶP: **{s['num_duplicate_groups']}**")
    L.append(f"- Trích dẫn không có mục (undefined): **{s['num_undefined_citations']}**")
    L.append(f"- Mục không được trích (unused): **{s['num_unused_references']}**")
    L.append(f"- Marker gãy ([?], ??): **{s['num_broken_markers']}**")
    L.append(f"- Sai khớp tên tác giả: **{s['num_author_name_issues']}**")
    L.append(f"- Mục thiếu trường (title/year/doi): **{s['num_refs_missing_fields']}**\n")

    if report["duplicate_groups"]:
        L.append("## ⚠️ Tài liệu trùng lặp (cùng DOI/tiêu đề, khác số)\n")
        for g in report["duplicate_groups"]:
            L.append(f"- **{' = '.join(g['members'])}** — {g['title'][:90]}"
                     + (f"  ·  DOI: {g['doi']}" if g['doi'] else ""))
        L.append("")

    if report["undefined_citations"]:
        L.append("## ⚠️ Trích dẫn không có trong danh mục\n")
        L.append(", ".join(f"[{n}]" for n in report["undefined_citations"]) + "\n")

    if report["missing_numbers_in_list"]:
        L.append("## Số bị thiếu trong dải đánh số\n")
        L.append(", ".join(str(n) for n in report["missing_numbers_in_list"]) + "\n")

    if report["unused_references"]:
        L.append("## Mục không được trích dẫn\n")
        L.append(", ".join(report["unused_references"]) + "\n")

    if report["broken_markers"]:
        L.append("## ⚠️ Marker gãy\n")
        for b in report["broken_markers"]:
            L.append(f"- `{b['marker']}` … {b['context']}")
        L.append("")

    if report["author_name_issues"]:
        L.append("## ⚠️ Sai khớp tên tác giả (Author et al. [n])\n")
        L.append("| Trong văn bản | Số | Tác giả trong danh mục | Ngữ cảnh |")
        L.append("|---|---|---|---|")
        for a in report["author_name_issues"]:
            L.append(f"| {a['surname']} | [{a['number']}] | "
                     f"{(a.get('ref_authors') or '')[:45]} | …{a['context'][:60]}… |")
        L.append("")

    # bảng đầy đủ để rà tay
    L.append("## Bảng chi tiết từng tham chiếu (để manual-check)\n")
    L.append("| # | Cited | Tiêu đề (parse) | Năm | DOI/arXiv | API | API title | Link |")
    L.append("|---|---|---|---|---|---|---|---|")
    for r in refs:
        api = r.api or {}
        best = api.get("best") or {}
        link = best.get("url") or (("https://doi.org/" + r.doi) if r.doi else r.url)
        flag = api.get("status", "—")
        dupmark = f" (=~{r.duplicate_of})" if r.duplicate_of else ""
        title = (r.title or "—")[:55].replace("|", "/")
        apititle = (best.get("title") or "")[:45].replace("|", "/")
        ident = r.ident + dupmark
        cited = "✓" if r.cited else ("·" if r.number is not None else "")
        doi = r.doi or (("arXiv:" + r.arxiv) if r.arxiv else "")
        L.append(f"| {ident} | {cited} | {title} | {r.year or '—'} | "
                 f"{doi or '—'} | {flag} | {apititle} | {link or ''} |")
    L.append("")

    Path(path).write_text("\n".join(L), encoding="utf-8")


def write_csv(refs, path):
    cols = ["ident", "number", "cited", "cite_count", "duplicate_of", "title",
            "year", "authors", "journal", "doi", "arxiv", "url",
            "api_status", "api_score", "api_title", "api_year", "api_doi",
            "api_link", "problems"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for r in refs:
            api = r.api or {}
            best = api.get("best") or {}
            w.writerow([
                r.ident, r.number, r.cited, r.cite_count, r.duplicate_of,
                r.title, r.year, r.authors, r.journal, r.doi, r.arxiv, r.url,
                api.get("status", ""), api.get("score", ""),
                best.get("title", ""), best.get("year", ""), best.get("doi", ""),
                best.get("url", ""), ";".join(r.problems),
            ])


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main():
    ap = argparse.ArgumentParser(
        description="Kiểm tra trích dẫn cho .docx / .tex+.bib / .pdf / .txt",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__)
    ap.add_argument("--input", "-i", nargs="+",
                    help="File đầu vào: .docx / .pdf / .txt / .md (chứa cả bài + danh mục)")
    ap.add_argument("--docx", nargs="+", help="(alias cũ của --input)")
    ap.add_argument("--tex", nargs="+", help="File .tex (chế độ LaTeX)")
    ap.add_argument("--bib", help="File .bib (đi kèm --tex)")
    ap.add_argument("--api", action="store_true", help="Bật đối chiếu online")
    ap.add_argument("--providers", nargs="+", default=["crossref", "openalex"],
                    choices=["crossref", "openalex", "arxiv"],
                    help="Nguồn tra cứu (mặc định: crossref openalex)")
    ap.add_argument("--mailto", default=None, help="Email cho polite pool (khuyến nghị)")
    ap.add_argument("--limit", type=int, default=None, help="Giới hạn số mục tra online")
    ap.add_argument("--no-cache", action="store_true")
    ap.add_argument("--dup-threshold", type=float, default=0.85)
    ap.add_argument("--out", default="citation_report", help="Tên file xuất (không đuôi)")
    args = ap.parse_args()

    inputs = args.input or args.docx
    if not inputs and not args.tex:
        ap.error("Cần --input (.docx/.pdf/.txt) HOẶC --tex (+--bib).")

    # ---- nạp dữ liệu & dựng danh mục + trích dẫn ----
    if args.tex:
        mode = "latex"
        tex_all = "\n".join(read_text_file(p) for p in args.tex)
        if not args.bib:
            ap.error("--tex cần kèm --bib")
        refs = parse_bib(read_text_file(args.bib))
        cite_keys = []
        for m in CITE_RE.finditer(tex_all):
            cite_keys += [k.strip() for k in m.group(1).split(",") if k.strip()]
        cite_set = set(cite_keys)
        bib_ids = {r.ident for r in refs}
        for r in refs:
            r.cite_count = cite_keys.count(r.ident)
            r.cited = r.cite_count > 0
        # cho khớp khung báo cáo numeric: dùng key
        intext_counts = {}        # numeric không áp dụng
        author_links = []
        broken = extract_broken_markers(tex_all)
        # undefined/unused theo key
        dup_groups = detect_duplicates(refs, args.dup_threshold)
        report = build_report(refs, intext_counts, author_links, broken,
                              dup_groups, [], mode)
        report["undefined_citations"] = sorted(cite_set - bib_ids)
        report["unused_references"] = sorted(bib_ids - cite_set)
        report["summary"]["num_undefined_citations"] = len(report["undefined_citations"])
        report["summary"]["num_unused_references"] = len(report["unused_references"])
    else:
        mode = "docx/text"
        text = load_document_text(inputs)
        body, _refs_region = split_body_and_refs(text)
        refs = parse_reference_entries(text)
        intext_counts = extract_numeric_intext(body)
        author_links = extract_author_links(body)
        broken = extract_broken_markers(body)
        dup_groups = detect_duplicates(refs, args.dup_threshold)
        num2ref = {r.number: r for r in refs if r.number is not None}
        author_issues = check_author_links(author_links, num2ref)
        report = build_report(refs, intext_counts, author_links, broken,
                              dup_groups, author_issues, mode)

    # ---- đối chiếu online ----
    if args.api:
        print("Đối chiếu online…", file=sys.stderr)
        ver = OnlineVerifier(args.providers, mailto=args.mailto,
                             use_cache=not args.no_cache)
        ver.run(refs, limit=args.limit)
        # cập nhật lại bảng references trong report
        report["references"] = [_ref_dict(r) for r in refs]
        bad = [r for r in refs if r.api and r.api.get("status") not in ("OK", None)]
        report["summary"]["num_api_warnings"] = len(bad)

    # ---- xuất ----
    Path(args.out + ".json").write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    write_markdown(report, refs, args.out + ".md")
    write_csv(refs, args.out + ".csv")

    s = report["summary"]
    print("\n=== TÓM TẮT ===")
    print(f"Mục tham khảo            : {s['num_references']}")
    print(f"Nhóm trùng lặp           : {s['num_duplicate_groups']}")
    print(f"Trích dẫn không có mục    : {s['num_undefined_citations']}")
    print(f"Mục không được trích      : {s['num_unused_references']}")
    print(f"Marker gãy               : {s['num_broken_markers']}")
    print(f"Sai khớp tên tác giả      : {s.get('num_author_name_issues', 0)}")
    print(f"Mục thiếu trường          : {s['num_refs_missing_fields']}")
    if args.api:
        print(f"Cảnh báo API             : {s.get('num_api_warnings', 0)}")
    print(f"\nĐã xuất: {args.out}.json / {args.out}.md / {args.out}.csv")


if __name__ == "__main__":
    main()
