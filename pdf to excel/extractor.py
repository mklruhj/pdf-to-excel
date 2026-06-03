"""
extractor.py — PDF extraction: Gemini Vision REST (primary) + robust OCR (fallback)

OCR pipeline (used when Gemini unavailable):
  1. Full-page OCR at 400 DPI — standard + whitelist variants
  2. Merged-slash recovery  (e.g. "00397397" → "0039/397")
     validated against known drawing numbers from clearer pages
  3. Cell-by-cell OCR for any remainder
"""
import base64
import io
import json
import os
import platform as _platform
import re
import shutil
from datetime import datetime
from typing import Optional

import pdfplumber
import requests
from dateutil import parser as dateparser


# ── Cross-platform path detection ─────────────────────────────────────────────
def _find_tesseract() -> Optional[str]:
    """Return Tesseract executable path for this system."""
    env = os.getenv('TESSERACT_CMD')
    if env and os.path.isfile(env):
        return env
    in_path = shutil.which('tesseract')
    if in_path:
        return in_path
    if _platform.system() == 'Windows':
        for p in [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]:
            if os.path.isfile(p):
                return p
    return None


def _find_poppler() -> Optional[str]:
    """Return Poppler bin directory, or None if it is already on PATH."""
    env = os.getenv('POPPLER_PATH')
    if env and os.path.isdir(env):
        return env
    if shutil.which('pdftoppm'):
        return None  # already on PATH (Linux/Mac or Windows with PATH configured)
    if _platform.system() == 'Windows':
        home = os.path.expanduser('~')
        for base in [os.path.join(home, 'Downloads'), r"C:\Program Files",
                     r"C:\Program Files (x86)", r"C:\tools"]:
            if not os.path.isdir(base):
                continue
            for name in os.listdir(base):
                for tail in [os.path.join('Library', 'bin'), 'bin']:
                    candidate = os.path.join(base, name, tail)
                    if os.path.isfile(os.path.join(candidate, 'pdftoppm.exe')):
                        return candidate
                    # one level deeper (e.g. release-folder/poppler-x.y/Library/bin)
                    inner = os.path.join(base, name)
                    if os.path.isdir(inner):
                        for sub in os.listdir(inner):
                            c2 = os.path.join(inner, sub, tail)
                            if os.path.isfile(os.path.join(c2, 'pdftoppm.exe')):
                                return c2
    return None


# ── Optional: pdf2image + PIL (needed for Gemini image conversion) ────────────
try:
    from pdf2image import convert_from_path
    from PIL import ImageEnhance, ImageFilter, Image
    POPPLER_PATH = _find_poppler()
    PDF2IMAGE_AVAILABLE = True
except Exception:
    PDF2IMAGE_AVAILABLE = False
    POPPLER_PATH = None

# ── Optional: Tesseract OCR (needed for OCR fallback) ────────────────────────
try:
    import pytesseract
    _tess = _find_tesseract()
    if _tess:
        pytesseract.pytesseract.tesseract_cmd = _tess
    OCR_AVAILABLE = PDF2IMAGE_AVAILABLE  # OCR needs both pytesseract and pdf2image
except Exception:
    OCR_AVAILABLE = False

# ── SSL verification (disable via DISABLE_SSL_VERIFY=true in .env for VPN/proxy)
_SSL_VERIFY = os.getenv('DISABLE_SSL_VERIFY', '').lower() not in ('1', 'true', 'yes')
if not _SSL_VERIFY:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

GEMINI_AVAILABLE = True
_GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-2.0-flash:generateContent"
)

# Character whitelist for item tags
_TAG_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789/- "


# ── Tag regex patterns ────────────────────────────────────────────────────────
#
# Format A: 066-PA-0039/397-002-S74   (after slash: NNN-NNN-LNN)
# Format B: 176-FW-0129/191-S14       (after slash: NNN-LNN)
#
GAL_TAG_STRICT = re.compile(
    r'\d{3}-[A-Z]{2,4}-\d{4}/\d{3}(?:-\d{3})?-[A-Z]\d{2,3}', re.IGNORECASE
)
GAL_TAG_LENIENT = re.compile(
    r'[0-9oOiIl]{3}[\s\-][A-Za-z]{2,4}[\s\-][0-9oOiIl]{4}'
    r'[/\\][0-9oOiIl]{3}(?:[\s\-][0-9oOiIl]{3})?[\s\-][A-Za-z][A-Za-z0-9]{2,3}'
)
PIP_TAG_RE = re.compile(
    r'MZ-\d{3}-[A-Z]{2,4}-[A-Z]{2,4}-ISO-[A-Z\d][\w\d]*-\d{2,4}', re.IGNORECASE
)
# Merged-slash: OCR dropped or misread '/' as '7', producing 7-8 digit number
# e.g. "066-PA-00397397-002-574" from "066-PA-0039/397-002-S74"
_MERGED_RE = re.compile(
    r'\b(\d{3})-([A-Za-z]{2,4})-(\d{7,8})-(\d{3})-([5689A-Za-z])(\d{2})\b'
)

NOI_RE = re.compile(
    r'MOZ[-\s]LNG[-\s][A-Z0-9][-A-Z0-9]*[-\s]QC[-\s]NOI[-\s]\d+', re.IGNORECASE
)
REPORT_RE = re.compile(
    r'REPORT\s+N[Oo°][°.\s]*:?\s*\n?\s*([A-Z]{2,5}\d[\dA-Z\-]+)', re.IGNORECASE
)
DATE_RE = re.compile(
    r'\b(\d{1,2}[-\s][A-Za-z]{3}[-\s]\d{2,4}'
    r'|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}'
    r'|\d{4}[-./]\d{1,2}[-./]\d{1,2}'
    r'|\d{1,2}[|/]\d{1,2}[|/]\d{4})\b',
    re.IGNORECASE,
)


# ── Date helpers ──────────────────────────────────────────────────────────────
def _parse_date(s):
    if not s:
        return None
    s = str(s).strip().replace('|', '/').replace('.', '/')
    s = re.sub(r'\s+', ' ', s)
    for fmt in ["%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %B %Y",
                "%d %b %y", "%Y/%m/%d", "%d/%m/%Y", "%d/%m/%y"]:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    try:
        return dateparser.parse(s, dayfirst=True)
    except Exception:
        return None


def _fmt(dt) -> str:
    return dt.strftime("%d-%b-%y").upper() if dt else ""


def normalise_date(value) -> str:
    dt = _parse_date(value)
    return _fmt(dt) if dt else (str(value).strip() if value else "")


# ── Tag normalisation ─────────────────────────────────────────────────────────
def _fix_ocr(s: str) -> str:
    return (s.replace('O', '0').replace('o', '0')
             .replace('I', '1').replace('i', '1').replace('l', '1').replace('|', '1')
             .replace(' ', ''))

# From: S s B b G g Z z q Q D d o O I i l L |   (19 chars)
# To:   5 5 8 6 6 9 2 2 9 9 0 0 0 0 1 1 1 1 1   (19 chars)
_DIGIT_FROM_LETTER = str.maketrans('SsBbGgZzqQDdoOIilL|', '5586692299000011111')

def _fix_spool_digits(s: str) -> str:
    """Fix letter→digit OCR errors in the numeric portion of a spool suffix.
    Keeps the leading spool-type letter (e.g. 'S') unchanged; converts
    ambiguous letters in the digit positions (S→5, G/b→6, g/q→9, Z→2, l→1, O→0).
    """
    if not s:
        return s
    leader = s[0].upper() if s[0].isalpha() else s[0]
    rest = s[1:].translate(_DIGIT_FROM_LETTER)
    return leader + rest


def _normalise_gal(raw: str):
    raw = re.sub(r'\s*([/\-])\s*', r'\1', str(raw).strip())
    parts = raw.split('/')
    if len(parts) != 2:
        return None
    left, right = parts
    lp = left.split('-')
    if len(lp) < 3:
        lp = left.split()   # OCR may use spaces instead of dashes
    if len(lp) < 3:
        return None
    lp[0] = _fix_ocr(lp[0])
    lp[2] = _fix_ocr(lp[2])
    left = '-'.join(lp).upper()
    rp = right.split('-')
    if len(rp) == 1:
        rp = right.split()  # OCR may use spaces instead of dashes
    rp = [_fix_ocr(p) for p in rp]
    if rp:
        rp[-1] = _fix_spool_digits(rp[-1])  # correct digit OCR errors in spool suffix
    right = '-'.join(rp).upper()
    result = f"{left}/{right}"
    if re.match(r'^\d{3}-[A-Z]{2,4}-\d{4}/\d{3}', result):
        return result
    return None


# ── Gemini REST ───────────────────────────────────────────────────────────────
def _img_part(img) -> dict:
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return {"inline_data": {"mime_type": "image/jpeg",
                            "data": base64.b64encode(buf.getvalue()).decode()}}


def _gemini_post(api_key: str, parts: list) -> str:
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"maxOutputTokens": 8192, "temperature": 0.0},
    }
    resp = requests.post(
        _GEMINI_URL,
        params={"key": api_key},
        json=payload,
        verify=_SSL_VERIFY,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()["candidates"][0]["content"]["parts"][0]["text"]


def _parse_json(text: str):
    text = re.sub(r'^```(?:json)?\s*\n?', '', text.strip())
    text = re.sub(r'\n?```\s*$', '', text)
    m = re.search(r'\{[\s\S]*\}', text)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return None


# ── Gemini prompts ────────────────────────────────────────────────────────────
_ITEM_HINTS = {
    'GAL': (
        'Items are in the DESCRIPTION column. Two formats:\n'
        '  "066-PA-0039/397-002-S74"  "176-FW-0129/191-S14"\n'
        '  "163-SW-0010/191-S02"  "263-SW-0003/291-S01"  "276-FW-0064/291-S19"'
    ),
    'PIP': (
        'Items are the SPOOL NUMBERS from the "SPOOL NUMBER" column, '
        'e.g. "141-HO-0091/191-S35", "141-HO-0091/191-S36". '
        'Each row has a unique spool number — do NOT use the Drawing Number (MZ-... format). '
        'There can be 5-30 rows per page.'
    ),
    'PAI': 'Items are pipe/spool tags like "066-PA-0039/397-002-S74" or "176-FW-0129/191-S15"',
    'UNK': 'Extract every item/tag identifier from the main data table.',
}
_ITEMS_PROMPT_TMPL = (
    "This is ONE page of a QC certificate (Mozambique LNG).\n"
    "Extract ALL item tags from every row of the main data table — there can be 30-60+ rows.\n\n"
    "{hint}\n\n"
    "Return ONLY valid JSON:\n"
    '{{"items": ["tag1", "tag2", ...]}}\n\n'
    'If no rows: {{"items": []}}'
)
_META_PROMPT = (
    "This is a QC certificate page (Mozambique LNG / TOTAL E&P / Yanda-CCSJV).\n"
    "Extract certificate metadata. Return ONLY valid JSON:\n"
    "{{\n"
    '  "report_no": "e.g. GAL010101M-YAN-00003",\n'
    '  "noi_no": "MOZ-LNG-... format",\n'
    '  "issue_date": "DD-MMM-YY e.g. 10-JAN-26",\n'
    '  "ccsjv_name": "CCSJV inspector full name",\n'
    '  "ccsjv_date": "DD-MMM-YY",\n'
    '  "client_name": "client inspector full name",\n'
    '  "client_date": "DD-MMM-YY"\n'
    "}}\n"
    'Use "" for missing fields.'
)


# ── Form type detection ───────────────────────────────────────────────────────
def _detect_form_type(pdf_path: str) -> str:
    name = os.path.basename(pdf_path).upper()
    for prefix in ('GAL', 'PIP', 'PAI'):
        if name.startswith(prefix):
            return prefix
    return 'UNK'


# ── Gemini extraction ─────────────────────────────────────────────────────────
def _gemini_extract(pdf_path: str, api_key: str):
    if not (api_key and PDF2IMAGE_AVAILABLE):
        return None
    try:
        imgs = convert_from_path(pdf_path, dpi=250, poppler_path=POPPLER_PATH)
    except Exception:
        return None

    form_type = _detect_form_type(pdf_path)
    items_prompt = _ITEMS_PROMPT_TMPL.format(
        hint=_ITEM_HINTS.get(form_type, _ITEM_HINTS['UNK'])
    )

    all_items, seen = [], set()
    for pg_img in imgs:
        try:
            raw = _gemini_post(api_key, [_img_part(pg_img), {"text": items_prompt}])
            data = _parse_json(raw)
            if data:
                for item in (data.get("items") or []):
                    item = str(item).strip()
                    if item and item not in seen:
                        seen.add(item); all_items.append(item)
        except Exception:
            pass

    if not all_items:
        return None

    meta = {k: "" for k in (
        "report_no", "noi_no", "issue_date",
        "ccsjv_name", "ccsjv_date", "client_name", "client_date",
    )}
    for pg_img in imgs:
        try:
            raw = _gemini_post(api_key, [_img_part(pg_img), {"text": _META_PROMPT}])
            page_meta = _parse_json(raw)
            if not page_meta:
                continue
            for k in meta:
                if not meta[k] and page_meta.get(k):
                    meta[k] = str(page_meta[k]).strip()
        except Exception:
            pass
        if all(meta[k] for k in ("report_no", "noi_no", "issue_date",
                                  "ccsjv_name", "client_name")):
            break

    for field in ("issue_date", "ccsjv_date", "client_date"):
        if meta[field]:
            meta[field] = normalise_date(meta[field])

    return {"items": all_items, **meta}


# ── OCR helpers ───────────────────────────────────────────────────────────────
def _preprocess(img):
    gray = img.convert("L")
    return ImageEnhance.Contrast(gray).enhance(3.0).filter(ImageFilter.SHARPEN)


def _preprocess_binarize(img, threshold=128):
    gray = img.convert("L")
    enhanced = ImageEnhance.Contrast(gray).enhance(4.0)
    return enhanced.point(lambda x: 255 if x > threshold else 0)


def _ocr_image(img) -> str:
    """Standard OCR for metadata — 2 sources × 2 PSM = 4 calls."""
    results = []
    for src in [img, _preprocess(img)]:
        for psm in ["3", "6"]:
            try:
                results.append(pytesseract.image_to_string(src, config=f"--oem 3 --psm {psm}"))
            except Exception:
                pass
    return "\n".join(results)


def _ocr_tags(img) -> str:
    """
    Whitelist OCR for items — 3 sources × 4 PSM = 12 calls.
    PSM 3: best on clear pages (page 2).
    PSM 11: best on garbled/poor-scan pages (page 1) — finds sparse text.
    PSM 12: additional items missed by others.
    PSM 6: adds unique items on clear pages.
    """
    results = []
    for src in [img, _preprocess(img), _preprocess_binarize(img, 128)]:
        for psm in ["3", "6", "11", "12"]:
            cfg = f"--oem 3 --psm {psm} -c tessedit_char_whitelist={_TAG_CHARS}"
            try:
                t = pytesseract.image_to_string(src, config=cfg).strip()
                if t:
                    results.append(t)
            except Exception:
                pass
    return "\n".join(results)


def _ocr_word_lines(img, config: str) -> list:
    """Return OCR text grouped into lines by word y-position.
    Adjacent words on the same row are concatenated without a separator,
    which reconstructs tags that OCR splits into multiple word tokens.
    """
    try:
        data = pytesseract.image_to_data(
            img, config=config, output_type=pytesseract.Output.DICT)
    except Exception:
        return []

    words = [
        (data['top'][i], data['left'][i], data['text'][i].strip())
        for i in range(len(data['text']))
        if data['text'][i].strip() and int(data['conf'][i]) > 5
    ]
    if not words:
        return []

    words.sort(key=lambda w: w[0])
    lines: list = []
    cur: list = []
    cur_y = -100
    for top, left, text in words:
        if abs(top - cur_y) > 18:
            if cur:
                lines.append(cur)
            cur = [(left, text)]
            cur_y = top
        else:
            cur.append((left, text))
    if cur:
        lines.append(cur)

    result = []
    for line in lines:
        line.sort(key=lambda w: w[0])
        result.append(''.join(w[1] for w in line))
    return result


# ── Item finding ──────────────────────────────────────────────────────────────
def _find_items_in_text(text: str, form_type: str, seen: set,
                         known_drawings: Optional[set] = None) -> list:
    # Rejoin tags split across lines by OCR (e.g. "141-HO-0091/191-\nS35")
    text = re.sub(r'-[ \t]*\n[ \t]*', '-', text)
    found = []

    def _add(tag):
        t = tag.strip().upper()
        if t and t not in seen:
            seen.add(t); found.append(t)

    if form_type == "PIP":
        for m in PIP_TAG_RE.finditer(text):
            _add(m.group(0))
        # Spool numbers (e.g. 141-HO-0091/191-S35) share the GAL tag format
        for m in GAL_TAG_STRICT.finditer(text):
            _add(m.group(0).upper())
        for m in GAL_TAG_LENIENT.finditer(text):
            n = _normalise_gal(m.group(0))
            if n:
                _add(n)
        # Recovery: OCR sometimes reads '/' as '-', giving 141-HO-0091-191-S35
        _noslash = re.compile(r'(\d{3}-[A-Z]{2,4}-\d{4})-(\d{3}-[A-Z]\d{2,3})', re.I)
        for m in _noslash.finditer(text):
            _add(f"{m.group(1)}/{m.group(2)}".upper())
        # Reconstruct split "MZ-...-ISO-" prefix + spool on next line
        prefix_re = re.compile(r'MZ[-\s]\d{3}[-\s]\w{2,4}[-\s]\w{2,4}[-\s]ISO[-\s]*', re.I)
        spool_re = re.compile(r'[A-Z][0O]?\d{4,5}[-~\s]+\d{3}', re.I)
        for pm in prefix_re.finditer(text):
            window = text[pm.start(): pm.start() + 120]
            sm = spool_re.search(window, len(pm.group(0)))
            if sm:
                raw = pm.group(0) + sm.group(0)
                norm = re.sub(r'[-~\s]+', '-', raw).upper().rstrip('-')
                if re.match(r'^MZ-\d{3}-[A-Z]{2,4}-[A-Z]{2,4}-ISO-[A-Z]\d', norm):
                    _add(norm)
    else:
        for m in GAL_TAG_STRICT.finditer(text):
            _add(m.group(0).upper())

        for m in GAL_TAG_LENIENT.finditer(text):
            n = _normalise_gal(m.group(0))
            if n:
                _add(n)

        # Merged-slash recovery: OCR reads '/' as '7' (8-digit) or drops it (7-digit).
        # Guard with known_drawings when available to prevent false positives.
        for m in _MERGED_RE.finditer(text):
            pfx, code, num, sec, sp_lead, sp_digits = m.groups()
            pfx_u, code_u = pfx.upper(), code.upper()
            candidates = []
            if len(num) == 8 and num[4] == '7':
                candidates.append((num[:4], num[5:]))
            if len(num) == 7:
                candidates.append((num[:4], num[4:]))
            for part1, part2 in candidates:
                drawing_key = f"{pfx_u}-{code_u}-{part1}/{part2}"
                if known_drawings is not None and drawing_key not in known_drawings:
                    continue
                sp = sp_lead.upper()
                sp_letter = 'S' if sp in '5689BD' else sp
                tag = f"{drawing_key}-{sec}-{sp_letter}{sp_digits}"
                if re.match(r'^\d{3}-[A-Z]{2,4}-\d{4}/\d{3}', tag):
                    _add(tag)
                break

    return found


# ── Cell-by-cell OCR ──────────────────────────────────────────────────────────
def _cell_by_cell_ocr(pdf_path: str, form_type: str, seen: set) -> list:
    if not OCR_AVAILABLE:
        return []
    found = []
    try:
        page_imgs = convert_from_path(pdf_path, dpi=350, poppler_path=POPPLER_PATH)
    except Exception:
        return []

    strategies = [
        {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
        {"vertical_strategy": "lines_strict", "horizontal_strategy": "lines_strict"},
    ]
    with pdfplumber.open(pdf_path) as pdf:
        for pg_idx, page in enumerate(pdf.pages):
            if pg_idx >= len(page_imgs):
                break
            pimg = page_imgs[pg_idx]
            pw, ph = page.width, page.height
            iw, ih = pimg.size
            sx, sy = iw / pw, ih / ph

            for strategy in strategies:
                try:
                    tables = page.find_tables(strategy)
                except Exception:
                    continue
                for tbl in tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            if cell is None:
                                continue
                            x0, top, x1, bot = cell
                            px0 = max(0, int(x0 * sx))
                            py0 = max(0, int(top * sy))
                            px1 = min(iw, int(x1 * sx))
                            py1 = min(ih, int(bot * sy))
                            if px1 - px0 < 10 or py1 - py0 < 4:
                                continue
                            cell_img = pimg.crop((px0, py0, px1, py1))
                            w, h = cell_img.size
                            if h < 40:
                                scale = max(2, 40 // max(h, 1))
                                cell_img = cell_img.resize(
                                    (w * scale, h * scale), Image.LANCZOS)
                            cell_text = _ocr_tags(cell_img)
                            found.extend(_find_items_in_text(
                                cell_text, form_type, seen))
    return found


# ── Full OCR fallback pipeline ────────────────────────────────────────────────
def _ocr_fallback(pdf_path: str):
    """
    4 OCR calls per page (2 standard + 2 whitelist) — fast single render.
    After OCR, two cheap text passes:
      Pass A: build known_drawings from strict matches across ALL pages (free)
      Pass B: extract items using known_drawings guard for merged-slash recovery
    Target: ~50-60s for a 2-page PDF.
    """
    form_type = _detect_form_type(pdf_path)
    seen: set = set()
    items: list = []
    full_text = ""

    _empty_meta = {k: "" for k in (
        "report_no", "noi_no", "issue_date",
        "ccsjv_name", "ccsjv_date", "client_name", "client_date",
    )}

    if not OCR_AVAILABLE:
        return [], _empty_meta

    try:
        imgs = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
    except Exception:
        return [], _empty_meta

    # ── Step 1: OCR all pages ────────────────────────────────────────────────
    page_std_texts: list = []    # standard OCR (reliable, for metadata + known_drawings)
    page_tag_texts: list = []    # whitelist OCR (more items, but noisier)
    for img in imgs:
        std_text = _ocr_image(img)   # 4 standard calls
        tag_text = _ocr_tags(img)    # 12 whitelist calls
        full_text += "\n" + std_text
        page_std_texts.append(std_text)
        page_tag_texts.append(tag_text)

    # ── Step 2: Build known_drawings from STANDARD OCR only ─────────────────
    # Using standard (no-whitelist) OCR avoids forced-character artefacts like
    # "0030" instead of "0039".  Page-2 strict matches establish all valid drawings.
    known_drawings: set = set()
    for std_text in page_std_texts:
        for m in GAL_TAG_STRICT.finditer(std_text):
            tag = m.group(0).upper()
            slash = tag.split('/')
            if len(slash) == 2:
                known_drawings.add(f"{slash[0]}/{slash[1].split('-')[0]}")

    # ── Step 3: Full item extraction with known_drawings guard ────────────────
    for std_text, tag_text in zip(page_std_texts, page_tag_texts):
        combined = std_text + "\n" + tag_text
        items.extend(_find_items_in_text(combined, form_type, seen, known_drawings))

    # ── Step 4: pdfplumber direct extraction ────────────────────────────────
    # For digital PDFs (e.g. PIP) embedded text is cleaner than OCR.
    # extract_tables() gives individual cell strings so tags split across
    # lines in extract_text() are still found intact.
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                pt = page.extract_text() or ""
                if pt.strip():
                    full_text += "\n" + pt
                    items.extend(_find_items_in_text(pt, form_type, seen, known_drawings))
                for tbl in (page.extract_tables() or []):
                    for row in (tbl or []):
                        for cell in (row or []):
                            if cell:
                                items.extend(_find_items_in_text(
                                    str(cell), form_type, seen, known_drawings))
    except Exception:
        pass

    # ── Step 5: High-DPI supplemental OCR for PIP spool recovery ────────────
    imgs_hd: list = []
    if form_type == 'PIP' and OCR_AVAILABLE:
        try:
            imgs_hd = convert_from_path(pdf_path, dpi=400, poppler_path=POPPLER_PATH)
            for img in imgs_hd:
                t = _ocr_tags(img)
                items.extend(_find_items_in_text(t, form_type, seen, known_drawings))
        except Exception:
            pass

    # ── Step 6: Word-level OCR reconstruction for PIP ────────────────────────
    # Tesseract sometimes splits one spool number across multiple word tokens.
    # image_to_data returns per-word bounding boxes; grouping by y-position and
    # concatenating adjacent words on the same line rebuilds those split tags.
    if form_type == 'PIP' and OCR_AVAILABLE:
        cfg = f"--oem 3 --psm 11 -c tessedit_char_whitelist={_TAG_CHARS}"
        for img in (imgs + imgs_hd):
            for src in [img, _preprocess(img), _preprocess_binarize(img, 128)]:
                for line in _ocr_word_lines(src, cfg):
                    items.extend(_find_items_in_text(line, form_type, seen, known_drawings))

    items = _clean_items(items, form_type, known_drawings)

    # ── Step 7: PIP prefix + suffix fragment recovery ────────────────────────
    # Scan OCR text for spool suffixes (S35, S4l→S41, S3S→S35) and combine
    # with the known drawing prefix.  Range-filter against found spool numbers
    # to suppress false positives from the Actual-Length column (S100, S775…).
    if form_type == 'PIP' and items:
        _pip_ok = re.compile(r'^\d{3}-[A-Z]{2,4}-\d{4}/\d{3}-[A-Z]\d{2,3}$')
        pfx_counts: dict = {}
        for item in items:
            m = re.match(r'^(\d{3}-[A-Z]{2,4}-\d{4}/\d{3})', item)
            if m:
                pfx_counts[m.group(1)] = pfx_counts.get(m.group(1), 0) + 1
        if pfx_counts:
            prefix = max(pfx_counts, key=pfx_counts.get)

            # Numeric range of already-found spools (±10 tolerance)
            found_nums = []
            for item in items:
                sm = re.search(r'-[A-Z](\d{2,3})$', item)
                if sm:
                    found_nums.append(int(sm.group(1)))
            lo = (min(found_nums) - 10) if found_nums else 0
            hi = (max(found_nums) + 10) if found_nums else 999

            all_ocr = ' '.join(page_std_texts + page_tag_texts)
            # Match S + 2 alphanumeric chars (catches S3S, S4l, S35, etc.)
            _sfx_re = re.compile(
                r'[Ss]([0-9SsBbGgZzqQlLiIoODd]{2})', re.IGNORECASE)
            for m in _sfx_re.finditer(all_ocr):
                raw = ('S' + m.group(1)).upper()
                fixed = _fix_spool_digits(raw)          # S3S→S35, S4l→S41
                nm = re.match(r'^[A-Z](\d{2,3})$', fixed)
                if not nm:
                    continue
                n = int(nm.group(1))
                if not (lo <= n <= hi):
                    continue                             # reject S100, S775…
                cand = f"{prefix}-{fixed}"
                if cand not in seen and _pip_ok.match(cand):
                    seen.add(cand)
                    items.append(cand)
            items = sorted(set(items))

    return items, _extract_meta(full_text)


# ── Item cleanup — remove false positives ────────────────────────────────────
_SPOOL_2DIG = re.compile(r'^[A-Z]\d{2}$')
_SPOOL_3DIG = re.compile(r'^[A-Z]\d{3}$')
_SPOOL_VALID = re.compile(r'^[A-Z]\d{2,3}$')

def _clean_items(items: list, form_type: str, known_drawings: set) -> list:
    """
    Remove OCR false positives:
    1. Dominant-suffix voting: for each prefix (e.g. 066-PA-0039), keep only the
       drawing suffix that appears most often.  Garbled OCR produces minority variants
       (0039/307 vs correct 0039/397 which has 4× more hits) — they get removed.
    2. Spool validation: reject SE5, SL1, S1E1 (non-digit chars in spool).
    3. 3-digit spool dedup: S431 is almost always S43 + stray row-number '1'.
       Keep 3-digit version only when no 2-digit equivalent exists.
    """
    from collections import defaultdict, Counter

    if form_type == 'PIP':
        pip_spool_ok = re.compile(r'^\d{3}-[A-Z]{2,4}-\d{4}/\d{3}(?:-\d{3})?-[A-Z]\d{2,3}$')
        return [i for i in items if pip_spool_ok.match(i)]

    if not items:
        return items

    # Step 1 — count how often each (prefix, drawing-suffix) pair appears
    prefix_suffix_counts: dict = defaultdict(Counter)
    for item in items:
        slash = item.split('/')
        if len(slash) == 2:
            pfx = slash[0]
            sfx = slash[1].split('-')[0]
            prefix_suffix_counts[pfx][sfx] += 1

    # Trusted drawings:
    #  - Always trust if confirmed by known_drawings (strict match on a clear page).
    #  - Without confirmation, require ≥3 count to filter rare OCR prefix errors
    #    like "0030" instead of "0039" or "776-FW" instead of "176-FW".
    trusted: set = set()
    for pfx, counts in prefix_suffix_counts.items():
        best_sfx, best_n = counts.most_common(1)[0]
        drawing = f"{pfx}/{best_sfx}"
        if drawing in known_drawings or best_n >= 3:
            trusted.add(drawing)

    if not trusted:
        return items  # can't decide; return unchanged

    # Step 2 — first pass: collect only 2-digit spool items
    two_digit_set: set = set()
    pass1: list = []
    for item in sorted(items):
        slash = item.split('/')
        if len(slash) != 2:
            continue
        pfx = slash[0]
        parts = slash[1].split('-')
        sfx, spool = parts[0], parts[-1]
        if f"{pfx}/{sfx}" not in trusted:
            continue
        if not _SPOOL_VALID.match(spool) or not spool[1:].isdigit():
            continue
        if _SPOOL_2DIG.match(spool):
            pass1.append(item)
            two_digit_set.add(item)

    # Step 3 — second pass: handle 3-digit spools
    pass2: list = []
    for item in sorted(items):
        slash = item.split('/')
        if len(slash) != 2:
            continue
        pfx = slash[0]
        parts = slash[1].split('-')
        sfx, spool = parts[0], parts[-1]
        if f"{pfx}/{sfx}" not in trusted:
            continue
        if not _SPOOL_3DIG.match(spool) or not spool[1:].isdigit():
            continue

        # If 3-digit spool ends in '1' it is almost always a 2-digit spool
        # with a stray row-number digit appended by OCR (e.g. S471 → S47).
        # Strip the trailing '1' and use the 2-digit form instead.
        if spool.endswith('1'):
            stripped = spool[:-1]
            stripped_item = item.replace(f"-{spool}", f"-{stripped}")
            if stripped_item not in two_digit_set:
                pass2.append(stripped_item)
                two_digit_set.add(stripped_item)
            continue  # never keep the 3-digit form in this case

        # Genuine 3-digit spool (e.g. S100) — keep only if no 2-digit version
        base_spool = spool[:3]
        base_item = item.replace(f"-{spool}", f"-{base_spool}")
        if base_item not in two_digit_set:
            pass2.append(item)

    combined = pass1 + pass2

    # Final dedup: if the same (prefix, drawing-suffix, spool) exists with different
    # middle-sections (e.g. -000- vs -002-), keep only the most-common section variant.
    from collections import defaultdict as _dd, Counter as _C
    section_groups: dict = _dd(list)
    for item in combined:
        slash = item.split('/')
        if len(slash) != 2:
            continue
        pfx = slash[0]
        parts = slash[1].split('-')
        sfx, spool = parts[0], parts[-1]
        section = '-'.join(parts[1:-1])   # e.g. "002", "000", or ""
        section_groups[(pfx, sfx, spool)].append((section, item))

    # Count how often each section appears globally (to find the "correct" one)
    global_section_counts: dict = _C()
    for item in combined:
        slash = item.split('/')
        if len(slash) == 2:
            parts = slash[1].split('-')
            section = '-'.join(parts[1:-1])
            global_section_counts[section] += 1

    final: list = []
    for (pfx, sfx, spool), variants in section_groups.items():
        if len(variants) == 1:
            final.append(variants[0][1])
        else:
            # Keep the variant whose section is most globally common
            best = max(variants, key=lambda x: global_section_counts.get(x[0], 0))
            final.append(best[1])

    return sorted(final)


# ── Metadata extraction ───────────────────────────────────────────────────────
def _extract_meta(text: str) -> dict:
    meta = {k: "" for k in (
        "report_no", "noi_no", "issue_date",
        "ccsjv_name", "ccsjv_date", "client_name", "client_date",
    )}

    m = REPORT_RE.search(text)
    if m:
        meta["report_no"] = m.group(1).strip()
    else:
        for pfx in ("PIP", "PAI", "GAL", "RFI", "GWT", "NDE"):
            m2 = re.search(rf"({pfx}\d[\dA-Z\-]+)", text)
            if m2:
                meta["report_no"] = m2.group(1)
                break

    m = NOI_RE.search(text)
    if m:
        meta["noi_no"] = re.sub(r"\s+", "", m.group(0)).upper()

    dated = []
    for m in DATE_RE.finditer(text):
        dt = _parse_date(m.group(1))
        if dt:
            dated.append((dt, m.start()))
    dated.sort(key=lambda x: x[1])

    seen_d, ordered = set(), []
    for dt, pos in dated:
        k = dt.date()
        if k not in seen_d:
            seen_d.add(k); ordered.append((dt, pos))

    if ordered:
        meta["issue_date"] = _fmt(min(d for d, _ in ordered))
    if len(ordered) >= 2:
        meta["ccsjv_date"] = _fmt(ordered[1][0])
    if len(ordered) >= 3:
        meta["client_date"] = _fmt(ordered[-1][0])

    # Names — look for "Firstname Lastname" patterns, reject form vocabulary
    name_re = re.compile(r'\b([A-Z][a-z]{1,15}(?:[\s\-][A-Z][a-zA-Z]{1,15})+)\b')
    _label_words = re.compile(
        r'\b(Mozambique|Mocambique|Limitada|Total|Project|Code|Form|Report|Page|'
        r'Rev|Surface|Steel|Shot|Paint|Insul|Galvan|Prefab|Piping|Spool|Check|'
        r'Release|Inspection|Notification|Reference|Certificate|Drawing|Number|'
        r'Area|Fabrication|Erection|Procedure|Internal|External|Accepted|'
        r'Company|Subcontractor|Client|Quality|Manager|Engineer|Inspec|'
        r'Lng|Ccx|Iso|Ndt|Qc|Dft|Rfi|Noi|Rev)\b',
        re.I
    )

    def _is_person_name(n: str) -> bool:
        if len(n) < 4 or len(n) > 40:
            return False
        if _label_words.search(n):
            return False
        return 2 <= len(n.split()) <= 4

    for kw in (r'IQC\b', r'CCSJV', r'CCS\b'):
        m = re.search(rf'{kw}.{{0,150}}', text, re.I | re.S)
        if m:
            for n in name_re.findall(m.group(0)):
                if _is_person_name(n):
                    meta["ccsjv_name"] = n
                    break
        if meta["ccsjv_name"]:
            break

    for kw in (r'\bTPI\b', r'CLIENT\b', r'COMPANY\s+NAME'):
        m = re.search(rf'{kw}.{{0,150}}', text, re.I | re.S)
        if m:
            for n in name_re.findall(m.group(0)):
                if _is_person_name(n):
                    meta["client_name"] = n
                    break
        if meta["client_name"]:
            break

    return meta


# ── Main pipeline ─────────────────────────────────────────────────────────────
def process_pdf(pdf_path: str, api_key: str = "") -> tuple:
    """Returns (rows, status, error_message)."""

    data = _gemini_extract(pdf_path, api_key)

    if data:
        items = data.get("items") or []
        meta = {k: data.get(k, "") for k in (
            "report_no", "noi_no", "issue_date",
            "ccsjv_name", "ccsjv_date", "client_name", "client_date",
        )}
        method = "gemini-vision"
    else:
        items, meta = _ocr_fallback(pdf_path)
        method = "ocr-fallback"

    if not items:
        items = ["NO_ITEM_FOUND"]

    required_ok = all(v not in ("", "NOT_FOUND")
                      for v in [meta["report_no"], meta["noi_no"], meta["issue_date"]])
    optional_ok = all(v not in ("", "NOT_FOUND")
                      for v in [meta["ccsjv_name"], meta["ccsjv_date"], meta["client_name"]])
    has_items = items != ["NO_ITEM_FOUND"]

    status = (
        "complete" if (required_ok and optional_ok and has_items)
        else "failed" if not has_items
        else "partial"
    )

    rows = []
    for item in items:
        rows.append({
            "Item":                   item,
            "RFI":                    meta["noi_no"],
            "Certificate_Code":       meta["report_no"],
            "Issue_date":             meta["issue_date"],
            "Notes":                  "",
            "Quality Signature Name": meta["ccsjv_name"],
            "Quality Signature Date": meta["ccsjv_date"],
            "Client Signature Name":  meta["client_name"],
            "Client Signature Date":  meta["client_date"],
            "_status":                status,
            "_method":                method,
        })

    return rows, status, None

