import csv
import json
import queue
import re
import sys
import threading
import tkinter as tk
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import customtkinter as ctk
import fitz  # PyMuPDF
from PIL import Image
from tkinter import filedialog, messagebox

try:
    import winreg
except ImportError:
    winreg = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None


DATE_PATTERN = re.compile(r"(\d{4}[-./]\d{1,2}[-./]\d{1,2}|\d{1,2}[-./]\d{1,2}[-./]\d{4})")
DATE_LABEL_PATTERNS = [
    re.compile(r"(?:발주일|주문일|수주일|po\s*date|release\s*date)\s*[:：]?\s*(\d{4}[-./]\d{1,2}[-./]\d{1,2})", re.IGNORECASE),
    re.compile(r"(?:발주일|주문일|수주일|po\s*date|release\s*date)\s*[:：]?\s*(\d{1,2}[-./]\d{1,2}[-./]\d{4})", re.IGNORECASE),
    re.compile(r"(?:발주일|주문일|수주일|po\s*date)\s*[:：]?\s*(\d{4}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일)", re.IGNORECASE),
]
ORDER_LABEL_PATTERNS = [
    re.compile(
        r"(?:발주번호|주문번호|수주번호|등록번호|po\s*number|po\s*no\.?|p/o\s*no\.?)\s*[:：#]?\s*([A-Z0-9][A-Z0-9\-_/\s]{2,})",
        re.IGNORECASE,
    ),
]
GENERIC_ORDER_FALLBACK = re.compile(r"\b[A-Z]{1,6}(?:[-/]\d+|\d+)(?:[-/]\d+)?\b")
PO_CODE_PATTERN = re.compile(r"\b[A-Z0-9][A-Z0-9\-/]{4,19}\b")
PO_ALLOWED_PATTERN = re.compile(r"^[A-Z0-9][A-Z0-9\-/]{4,19}$")
DISALLOWED_PO_TOKENS = {"shall", "upon", "terms", "conditions", "delivery", "acceptance"}
COMPANY_LABEL_EXCLUDE = {"company", "supplier", "vendor", "contact person", "telephone"}
STRICT_COMPANY_BANNED_TOKENS = {
    "tel", "fax", "telephone", "phone", "mobile", "contact person",
    "vendor code", "supplier code", "code:", "email", "@", "http", "www",
}
HEADER_COMPANY_EXCLUDE_TOKENS = {"buyer", "contact", "attn", "tel", "fax", "email", "@", "phone", "mobile"}
TERMS_KEYWORDS = {
    "terms", "conditions", "delivery", "acceptance", "warranty", "liability",
    "agreement", "payment", "shall", "upon",
}
CORE_FIELD_LABELS = [
    re.compile(r"\bcompany\b", re.IGNORECASE),
    re.compile(r"\bpo\s*number\b", re.IGNORECASE),
    re.compile(r"\bpo\s*no\.?\b", re.IGNORECASE),
    re.compile(r"\bp\s*/\s*o\s*no\.?\b", re.IGNORECASE),
    re.compile(r"\bpo\s*date\b", re.IGNORECASE),
    re.compile(r"\brelease\s*date\b", re.IGNORECASE),
]
SUPPORTED_EXTENSIONS = {".pdf"}
LANDSCAPE_SIZE = (1200, 800)
PORTRAIT_SIZE = (800, 1200)
RENDER_ZOOM = 2.0
MISSING_VALUE = "확인필요"
TITLE_PREFIX = "[주문서]"
TESSERACT_CANDIDATES = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
]
REGISTRY_BASE_KEY = r"Software\OrderConverterStudio"
REGISTRY_MEMORY_VALUE = "CompanyAliasMemory"
REGISTRY_EXPORT_VALUE = "LastMemoryExport"
AUTO_COMPANY_EXCLUDE_NAMES = [
    "케이엑스하이텍",
    "케이엑스 하이텍",
    "KX HITECH",
    "KXHITECH",
    "(주)케이엑스하이텍",
    "주식회사 케이엑스하이텍",
]
COMPANY_STOPWORDS = {
    "purchase order", "order sheet", "invoice", "quotation", "견적서", "발주서", "주문서", "거래명세서",
    "packing list", "commercial invoice", "proforma invoice", "ship to", "bill to", "buyer", "seller",
    "vendor", "supplier", "customer", "consignee", "notify", "attn", "tel", "fax", "email", "address",
}
AUTO_COMPANY_LABEL_PATTERNS = [
    re.compile(r"(?:supplier|vendor|seller|maker|manufacturer|from)\s*[:：]\s*([^\n]{2,80})", re.IGNORECASE),
    re.compile(r"(?:공급자|판매자|납품처|제조사|업체명|상호)\s*[:：]\s*([^\n]{2,80})", re.IGNORECASE),
    re.compile(r"(?:bill\s*to|ship\s*to|buyer|customer|consignee|수신|공급받는자|발주처)\s*[:：]\s*([^\n]{2,80})", re.IGNORECASE),
]
CORPORATE_NAME_LINE_PATTERNS = [
    re.compile(r"([가-힣A-Za-z0-9&().,/\-\s]{2,80}(?:주식회사|㈜|Co\.?\s*,?\s*Ltd\.?|CO\.?\s*,?\s*LTD\.?|Inc\.?|LLC))", re.IGNORECASE),
    re.compile(r"((?:주식회사|㈜)\s*[가-힣A-Za-z0-9&().,/\-\s]{2,60})", re.IGNORECASE),
]



@dataclass
class CompanyRule:
    display_name: str
    aliases: List[str] = field(default_factory=list)
    order_patterns: List[re.Pattern] = field(default_factory=list)
    source: str = "companies.txt"

    @property
    def all_names(self) -> List[str]:
        values = [self.display_name, *self.aliases]
        return unique_preserve_order([value for value in values if value])


@dataclass
class DocumentInfo:
    pdf_path: Path
    company_name: str
    document_date: str
    order_numbers: List[str]
    representative_order_number: str
    page_count: int
    status: str
    text_excerpt: str = ""
    used_ocr: bool = False
    company_match_status: str = ""
    raw_order_candidates: List[str] = field(default_factory=list)
    matched_alias: str = ""
    company_rule_source: str = ""


@dataclass
class ProgressEvent:
    event_type: str
    message: str = ""
    current_file: int = 0
    total_files: int = 0
    current_page: int = 0
    total_pages: int = 0
    success_count: int = 0
    fail_count: int = 0
    documents: List[DocumentInfo] = field(default_factory=list)


def sanitize_filename_part(text: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", text.strip())
    return cleaned or MISSING_VALUE



def compile_order_patterns(patterns: List[str]) -> List[re.Pattern]:
    compiled: List[re.Pattern] = []
    for pattern in patterns:
        pattern_text = pattern.strip()
        if not pattern_text:
            continue
        try:
            compiled.append(re.compile(pattern_text, re.IGNORECASE))
        except re.error:
            continue
    return compiled


def load_company_rules(companies_path: Path) -> List[CompanyRule]:
    rules: List[CompanyRule] = []
    csv_path = companies_path.with_name("companies_rules.csv")

    if csv_path.exists():
        with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)
            for row in reader:
                display_name = (row.get("display_name") or row.get("company_name") or "").strip()
                if not display_name:
                    continue
                aliases = [
                    alias.strip()
                    for alias in re.split(r"[;,]", row.get("aliases", ""))
                    if alias.strip()
                ]
                order_patterns = compile_order_patterns(
                    [item.strip() for item in re.split(r"[;|]", row.get("order_regexes", "")) if item.strip()]
                )
                rules.append(
                    CompanyRule(
                        display_name=display_name,
                        aliases=aliases,
                        order_patterns=order_patterns,
                        source=csv_path.name,
                    )
                )

    if not rules and companies_path.exists():
        with companies_path.open("r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()
                if not line:
                    continue

                if "|" in line:
                    parts = [part.strip() for part in line.split("|")]
                    display_name = parts[0] if parts else ""
                    aliases = [alias.strip() for alias in re.split(r"[;,]", parts[1])] if len(parts) >= 2 and parts[1].strip() else []
                    order_patterns = compile_order_patterns(
                        [item.strip() for item in re.split(r"[;|]", parts[2])] if len(parts) >= 3 and parts[2].strip() else []
                    )
                    if display_name:
                        rules.append(
                            CompanyRule(
                                display_name=display_name,
                                aliases=aliases,
                                order_patterns=order_patterns,
                                source=companies_path.name,
                            )
                        )
                else:
                    rules.append(CompanyRule(display_name=line, aliases=[], order_patterns=[], source=companies_path.name))

    rules.sort(key=lambda rule: max((len(normalize_for_match(name)) for name in rule.all_names), default=0), reverse=True)
    return rules



def configure_tesseract() -> bool:
    """Windows 기본 설치 경로를 우선 확인하고, 있으면 pytesseract에 연결한다."""
    if pytesseract is None:
        return False

    if pytesseract.pytesseract.tesseract_cmd:
        current = Path(pytesseract.pytesseract.tesseract_cmd)
        if current.exists():
            return True

    for candidate in TESSERACT_CANDIDATES:
        if candidate.exists():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            return True

    return False


def normalize_date(raw_date: str) -> str:
    compact = re.sub(r"\s+", "", raw_date)
    compact = compact.replace(".", "-").replace("/", "-")
    compact = compact.replace("년", "-").replace("월", "-").replace("일", "")
    parts = compact.split("-")
    if len(parts) != 3:
        return MISSING_VALUE

    try:
        first, second, third = (int(part) for part in parts)
        if first >= 1900:
            year, month, day = first, second, third
        elif third >= 1900:
            year, month, day = third, first, second
        else:
            return MISSING_VALUE
        parsed = datetime(year, month, day)
        return parsed.strftime("%Y-%m-%d")
    except ValueError:
        return MISSING_VALUE


def is_date_like_number(compact: str) -> bool:
    if not re.fullmatch(r"\d{8}", compact):
        return False
    year_first = int(compact[:4])
    month_first = int(compact[4:6])
    day_first = int(compact[6:8])
    if 1900 <= year_first <= 2100 and 1 <= month_first <= 12 and 1 <= day_first <= 31:
        return True

    month_second = int(compact[:2])
    day_second = int(compact[2:4])
    year_second = int(compact[4:8])
    return 1 <= month_second <= 12 and 1 <= day_second <= 31 and 1900 <= year_second <= 2100


def is_full_date_token(value: str) -> bool:
    token = value.strip()
    if not token:
        return False
    if re.fullmatch(r"\d{4}[-./]\d{1,2}[-./]\d{1,2}", token):
        return normalize_date(token) != MISSING_VALUE
    if re.fullmatch(r"\d{1,2}[-./]\d{1,2}[-./]\d{4}", token):
        return normalize_date(token) != MISSING_VALUE
    if re.fullmatch(r"\d{8}", token):
        return is_date_like_number(token)
    return False


def normalize_for_match(text: str) -> str:
    """회사명 비교를 위해 공백, 줄바꿈, 구분기호를 제거한다."""
    lowered = text.lower()
    return re.sub(r"[\s\-_()/\\.,:]+", "", lowered)


def normalize_document_text(text: str) -> str:
    """OCR과 본문 텍스트에서 끊긴 줄/하이픈을 복원해 추출 성공률을 높인다."""
    normalized = text.replace("\r", "\n")
    normalized = re.sub(r"([A-Za-z]{2,}\d{2,})\s*-\s*\n\s*(\d{2,})", r"\1-\2", normalized)
    normalized = re.sub(r"([A-Za-z]{2,}\d{2,})\s*-\s*(\d{2,})", r"\1-\2", normalized)
    normalized = re.sub(
        r"((?:발주번호|주문번호|수주번호|등록번호|po\s*number|po\s*no\.?|p/o\s*no\.?)\s*[:：#]?)\s*\n+\s*",
        r"\1 ",
        normalized,
        flags=re.IGNORECASE,
    )
    normalized = re.sub(r"[ \t]+", " ", normalized)
    return normalized


def clean_order_candidate(candidate: str) -> str:
    cleaned = candidate.strip().strip(".,;:)]}")
    cleaned = re.sub(r"\s+", "", cleaned)
    cleaned = re.sub(r"-{2,}", "-", cleaned)
    cleaned = re.sub(r"/{2,}", "/", cleaned)
    return cleaned


def extract_date_from_filename(filename: str) -> str:
    match = DATE_PATTERN.search(filename)
    if not match:
        return MISSING_VALUE
    return normalize_date(match.group(1))


def extract_text(document: fitz.Document) -> str:
    parts: List[str] = []
    for page in document:
        try:
            parts.append(page.get_text())
        except Exception:
            continue
    return normalize_document_text("\n".join(parts))


def perform_ocr_on_document(document: fitz.Document) -> str:
    """스캔본 PDF를 위해 각 페이지 이미지를 OCR로 읽는다."""
    if pytesseract is None:
        return ""

    text_parts: List[str] = []
    for page in document:
        try:
            image = render_page_to_image(page)
            text_parts.append(pytesseract.image_to_string(image, lang="kor+eng"))
        except Exception:
            continue
    return normalize_document_text("\n".join(text_parts))


def perform_ocr_on_top_region(page: fitz.Page) -> str:
    if pytesseract is None:
        return ""
    try:
        rect = page.rect
        clip = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.4)
        matrix = fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)
        pixmap = page.get_pixmap(matrix=matrix, alpha=False, clip=clip)
        image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        return normalize_document_text(pytesseract.image_to_string(image, lang="kor+eng"))
    except Exception:
        return ""



def is_excluded_company_name(name: str) -> bool:
    normalized = normalize_for_match(name)
    if not normalized:
        return True
    for banned in AUTO_COMPANY_EXCLUDE_NAMES:
        banned_normalized = normalize_for_match(banned)
        if banned_normalized and banned_normalized in normalized:
            return True
    return False



def clean_company_candidate(candidate: str) -> str:
    candidate = candidate.strip().strip("|:-/\\")
    candidate = re.sub(r"\s+", " ", candidate)
    candidate = re.sub(r"\b(?:tel|fax|email|attn|address|phone)\b.*$", "", candidate, flags=re.IGNORECASE).strip()
    candidate = candidate.strip(" ,;:|/-")
    return candidate



def is_plausible_company_candidate(candidate: str) -> bool:
    value = clean_company_candidate(candidate)
    if len(value) < 2 or len(value) > 80:
        return False
    normalized = normalize_for_match(value)
    if not normalized or normalized.isdigit():
        return False
    lowered = value.lower().strip()
    if lowered in COMPANY_STOPWORDS:
        return False
    if is_excluded_company_name(value):
        return False
    digit_ratio = sum(ch.isdigit() for ch in value) / max(len(value), 1)
    if digit_ratio > 0.45:
        return False
    return True


def is_valid_company_candidate_strict(candidate: str) -> bool:
    value = clean_company_candidate(candidate)
    if len(value) < 2 or len(value) > 70:
        return False
    if "\n" in candidate:
        return False
    if is_excluded_company_name(value):
        return False
    lowered = value.lower()
    if any(token in lowered for token in STRICT_COMPANY_BANNED_TOKENS):
        return False
    if ":" in value:
        return False
    if re.search(r"[\w\.-]+@[\w\.-]+\.\w+", value):
        return False
    if re.search(r"(?:\+82|82-|031-|02-|010-)", value):
        return False
    if re.search(r"\b(?:zip|postal\s*code|postcode)\b", lowered):
        return False
    if re.search(r"\b\d{3,5}-\d{3,5}\b", value):
        return False
    if looks_like_person_name(value):
        return False
    if re.fullmatch(r"[\d\W]+", value):
        return False

    digit_ratio = sum(ch.isdigit() for ch in value) / max(len(value), 1)
    if digit_ratio >= 0.4:
        return False

    address_markers = ["road", "street", "st.", "avenue", "building", "floor", "dong", "gu", "si", "city", "address"]
    if len(value) >= 45 and any(marker in lowered for marker in address_markers):
        return False

    has_corp_suffix = any(token in lowered for token in [" co", " ltd", " inc", " llc", " corp", "주식회사", "㈜"])
    has_letter_dominance = (
        (sum(ch.isalpha() or ("가" <= ch <= "힣") for ch in value) / max(len(value), 1)) >= 0.6
        and digit_ratio <= 0.2
    )
    return has_corp_suffix or has_letter_dominance


def looks_like_person_name(candidate: str) -> bool:
    value = clean_company_candidate(candidate)
    if not value:
        return False
    words = [word for word in re.split(r"\s+", value) if word]
    if len(words) < 2 or len(words) > 3:
        return False
    lowered = value.lower()
    if any(token in lowered for token in ["co", "ltd", "inc", "corp", "llc", "주식회사", "㈜"]):
        return False
    alpha_only_words = [word for word in words if re.fullmatch(r"[A-Za-z][A-Za-z'.-]{1,20}", word)]
    if len(alpha_only_words) != len(words):
        return False
    title_case_like = sum(1 for word in words if word[0].isupper() and word[1:].islower()) >= max(2, len(words) - 1)
    return title_case_like



def score_company_candidate(candidate: str, source: str = "") -> int:
    value = clean_company_candidate(candidate)
    lowered = value.lower()
    score = 0
    if source == "label_primary":
        score += 70
    elif source == "label_secondary":
        score += 45
    elif source == "top_lines":
        score += 25
    if any(token in lowered for token in ["주식회사", "㈜", "co", "ltd", "inc", "llc"]):
        score += 20
    if re.search(r"[가-힣]", value):
        score += 8
    if re.search(r"[A-Za-z]", value):
        score += 5
    if len(value) <= 24:
        score += 6
    if len(value) <= 40:
        score += 4
    if any(word in lowered for word in ["buyer", "bill to", "ship to", "consignee", "customer", "수신", "공급받는자", "발주처"]):
        score -= 15
    return score



def collect_auto_company_candidates(full_text: str) -> List[str]:
    candidates: List[Tuple[int, str]] = []
    lines = [line.strip() for line in full_text.splitlines() if line.strip()]

    for pattern_index, pattern in enumerate(AUTO_COMPANY_LABEL_PATTERNS):
        source = "label_primary" if pattern_index < 2 else "label_secondary"
        for match in pattern.finditer(full_text):
            candidate = clean_company_candidate(match.group(1))
            if is_plausible_company_candidate(candidate):
                candidates.append((score_company_candidate(candidate, source), candidate))

    for line in lines[:40]:
        candidate = clean_company_candidate(line)
        if is_plausible_company_candidate(candidate):
            if any(token in candidate.lower() for token in ["주식회사", "㈜", "co", "ltd", "inc", "llc"]):
                candidates.append((score_company_candidate(candidate, "top_lines"), candidate))
            for pattern in CORPORATE_NAME_LINE_PATTERNS:
                for match in pattern.finditer(line):
                    inner = clean_company_candidate(match.group(1))
                    if is_plausible_company_candidate(inner):
                        candidates.append((score_company_candidate(inner, "top_lines"), inner))

    unique: Dict[str, Tuple[int, str]] = {}
    for score, candidate in candidates:
        key = normalize_for_match(candidate)
        if not key:
            continue
        previous = unique.get(key)
        if previous is None or score > previous[0]:
            unique[key] = (score, candidate)

    ordered = sorted(unique.values(), key=lambda item: (-item[0], len(item[1])))
    return [candidate for _score, candidate in ordered[:5]]


def is_valid_po_number(candidate: str) -> bool:
    raw = candidate.strip()
    if not raw:
        return False
    if len(raw) < 5 or len(raw) > 20:
        return False
    if re.search(r"\s{3,}", raw):
        return False
    if ":" in raw:
        return False
    tokens = [token for token in raw.split() if token]
    if len(tokens) >= 4:
        return False
    lowered = raw.lower()
    if any(token in lowered for token in DISALLOWED_PO_TOKENS):
        return False
    punctuation_count = sum(raw.count(symbol) for symbol in [".", ",", ";", ":"])
    if punctuation_count >= 2:
        return False
    lower_alpha = sum(ch.isalpha() and ch.islower() for ch in raw)
    if lower_alpha >= 3:
        return False
    if len(re.findall(r"[A-Za-z]{3,}", raw)) >= 3 and len(tokens) >= 3:
        return False
    if re.search(r"[\w\.-]+@[\w\.-]+\.\w+", raw):
        return False
    if re.search(r"(?:\+82|82-|031-|02-|010-)", raw):
        return False
    compact = clean_order_candidate(raw)
    # 날짜 제거는 "문자열 전체가 날짜일 때"만 적용한다.
    # 예: PO20260124-123 / AB-20260124-99 는 허용.
    if is_full_date_token(raw) or is_full_date_token(compact):
        return False
    if not PO_ALLOWED_PATTERN.fullmatch(compact):
        return False
    if re.search(r"\d", compact) is None:
        return False
    if re.search(r"[A-Z]", compact) is None and re.search(r"[0-9]", compact) is None:
        return False
    return True


def extract_po_from_filename(filename_stem: str) -> List[str]:
    normalized = filename_stem.replace("_", " ").replace(".", " ").strip()
    candidates: List[str] = []
    for match in re.finditer(r"[A-Z0-9\-/]{5,20}", normalized.upper()):
        token = match.group(0).strip("-/")
        if not token:
            continue
        lowered = token.lower()
        if any(word in lowered for word in ["shall", "terms", "conditions", "order", "delivery", "acceptance"]):
            continue
        if is_valid_po_number(token):
            candidates.append(clean_order_candidate(token))
    return unique_preserve_order(candidates)


def has_core_label(text: str) -> bool:
    return any(pattern.search(text) for pattern in CORE_FIELD_LABELS)


def is_terms_block(text: str) -> bool:
    content = text.strip()
    if not content:
        return False
    lowered = content.lower()
    keyword_hits = sum(1 for keyword in TERMS_KEYWORDS if keyword in lowered)
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    english_sentence_lines = [
        line for line in lines
        if len(line) >= 40 and re.search(r"[A-Za-z]{4,}", line) and sum(line.count(s) for s in [".", ",", ";"]) >= 2
    ]
    if has_core_label(content):
        return False
    return len(english_sentence_lines) >= 7 and keyword_hits >= 2


def is_terms_page(page_text: str) -> bool:
    text = page_text.strip()
    if not text:
        return False
    lowered = text.lower()
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    english_sentence_lines = [
        line for line in lines
        if len(line) >= 40 and re.search(r"[A-Za-z]{4,}", line) and sum(line.count(s) for s in [".", ",", ";"]) >= 2
    ]
    keyword_hits = sum(1 for keyword in TERMS_KEYWORDS if keyword in lowered)
    return len(english_sentence_lines) >= 7 and keyword_hits >= 2 and not has_core_label(text)


def _to_block_dict(block: Tuple) -> Optional[Dict[str, object]]:
    if len(block) < 5:
        return None
    text = str(block[4]).strip()
    if not text:
        return None
    return {
        "x0": float(block[0]),
        "y0": float(block[1]),
        "x1": float(block[2]),
        "y1": float(block[3]),
        "text": text,
    }


def find_label_block(blocks: List[Dict[str, object]], label_patterns: List[re.Pattern]) -> Optional[Dict[str, object]]:
    for block in blocks:
        text = str(block["text"])
        for pattern in label_patterns:
            if pattern.search(text):
                return block
    return None


def get_nearby_value(label_block: Dict[str, object], blocks: List[Dict[str, object]]) -> str:
    lx0 = float(label_block["x0"])
    ly0 = float(label_block["y0"])
    lx1 = float(label_block["x1"])
    ly1 = float(label_block["y1"])
    line_h = max(12.0, ly1 - ly0)

    def eligible_right(block: Dict[str, object]) -> bool:
        return (
            float(block["x0"]) >= lx1 - 3
            and abs(float(block["y0"]) - ly0) <= line_h * 0.6
            and float(block["x0"]) <= lx1 + 420
        )

    def eligible_down(block: Dict[str, object]) -> bool:
        return (
            abs(float(block["x0"]) - lx0) <= 80
            and float(block["y0"]) > ly1 - 2
            and float(block["y0"]) <= ly1 + line_h * 3.5
        )

    def eligible_diag(block: Dict[str, object]) -> bool:
        return (
            float(block["x0"]) > lx1 - 3
            and float(block["x0"]) <= lx1 + 420
            and float(block["y0"]) > ly1 - 2
            and float(block["y0"]) <= ly1 + line_h * 3.5
        )

    right = [block for block in blocks if block is not label_block and eligible_right(block)]
    if right:
        right.sort(key=lambda block: (abs(float(block["y0"]) - ly0), float(block["x0"])))
        return str(right[0]["text"]).strip()

    down = [block for block in blocks if block is not label_block and eligible_down(block)]
    if down:
        down.sort(key=lambda block: (float(block["y0"]), abs(float(block["x0"]) - lx0)))
        return str(down[0]["text"]).strip()

    diag = [block for block in blocks if block is not label_block and eligible_diag(block)]
    if diag:
        diag.sort(key=lambda block: (float(block["y0"]), float(block["x0"])))
        return str(diag[0]["text"]).strip()

    return ""


def extract_from_blocks(page: fitz.Page) -> Dict[str, object]:
    rect = page.rect
    blocks_raw = page.get_text("blocks")
    blocks: List[Dict[str, object]] = []
    for block in blocks_raw:
        converted = _to_block_dict(block)
        if converted is None:
            continue
        if is_terms_block(str(converted["text"])):
            continue
        blocks.append(converted)
    blocks.sort(key=lambda item: (float(item["y0"]), float(item["x0"])))

    company_label_patterns = [re.compile(r"\bcompany\b", re.IGNORECASE)]
    vendor_label_patterns = [re.compile(r"\b(?:vendor|supplier|seller|from)\b", re.IGNORECASE)]
    po_label_patterns = [
        re.compile(r"\bpo\s*number\b", re.IGNORECASE),
        re.compile(r"\bpo\s*no\.?\b", re.IGNORECASE),
        re.compile(r"\bp\s*/\s*o\s*no\.?\b", re.IGNORECASE),
    ]
    date_label_patterns = [
        re.compile(r"\bpo\s*date\b", re.IGNORECASE),
        re.compile(r"\brelease\s*date\b", re.IGNORECASE),
    ]

    company_value = ""
    po_candidates: List[str] = []
    date_value = ""

    def corp_token_bonus(text: str) -> int:
        lowered = text.lower()
        if any(token in lowered for token in [" co", " ltd", " inc", " corp", " llc", "주식회사", "㈜"]):
            return 24
        if re.fullmatch(r"[A-Z0-9&().,\-/\s]{4,40}", text) and sum(ch.isalpha() for ch in text) >= 3:
            return 12
        return 0

    def evaluate_company_text(text: str, require_corp: bool = False) -> Optional[str]:
        if text.count("\n") >= 1:
            return None
        cleaned = clean_company_candidate(text)
        if not cleaned:
            return None
        lowered = cleaned.lower()
        if lowered in COMPANY_LABEL_EXCLUDE:
            return None
        if looks_like_person_name(cleaned):
            return None
        if not is_valid_company_candidate_strict(cleaned):
            return None
        if require_corp and corp_token_bonus(cleaned) <= 0:
            return None
        return cleaned

    def collect_label_candidates(label_patterns: List[re.Pattern], base_score: int, require_corp: bool = False) -> List[Tuple[int, str]]:
        label = find_label_block(blocks, label_patterns)
        if label is None:
            return []
        lx0 = float(label["x0"])
        ly0 = float(label["y0"])
        lx1 = float(label["x1"])
        ly1 = float(label["y1"])
        line_h = max(12.0, ly1 - ly0)
        nearby_pool: List[Tuple[Dict[str, object], int]] = []

        right = [
            block for block in blocks
            if block is not label
            and float(block["x0"]) >= lx1 - 3
            and abs(float(block["y0"]) - ly0) <= line_h * 0.6
            and float(block["x0"]) <= lx1 + 420
        ]
        right.sort(key=lambda block: (abs(float(block["y0"]) - ly0), float(block["x0"])))
        nearby_pool.extend((block, 40) for block in right[:2])

        diag = [
            block for block in blocks
            if block is not label
            and float(block["x0"]) >= lx1 - 3
            and float(block["x0"]) <= lx1 + 420
            and float(block["y0"]) > ly1 - 2
            and float(block["y0"]) <= ly1 + line_h * 2.6
        ]
        diag.sort(key=lambda block: (float(block["y0"]), float(block["x0"])))
        nearby_pool.extend((block, 20) for block in diag[:1])

        down = [
            block for block in blocks
            if block is not label
            and abs(float(block["x0"]) - lx0) <= 120
            and float(block["y0"]) > ly1 - 2
            and float(block["y0"]) <= ly1 + line_h * 3.0
        ]
        down.sort(key=lambda block: (float(block["y0"]), abs(float(block["x0"]) - lx0)))
        nearby_pool.extend((block, 10) for block in down[:2])

        scored: List[Tuple[int, str]] = []
        for block, pos_bonus in nearby_pool:
            candidate = evaluate_company_text(str(block["text"]), require_corp=require_corp)
            if not candidate:
                continue
            distance_penalty = int(abs(float(block["y0"]) - ly0) + abs(float(block["x0"]) - lx1) * 0.1)
            score = base_score + pos_bonus + corp_token_bonus(candidate) - distance_penalty
            scored.append((score, candidate))
        return scored

    header_y_limit = rect.y0 + rect.height * 0.30
    left_header_x_limit = rect.x0 + rect.width * 0.40
    right_header_x_limit = rect.x0 + rect.width * 0.60
    left_header_blocks = [
        block for block in blocks
        if float(block["y0"]) <= header_y_limit
        and float(block["x0"]) <= left_header_x_limit
    ]
    right_header_blocks = [
        block for block in blocks
        if float(block["y0"]) <= header_y_limit
        and float(block["x0"]) >= right_header_x_limit
    ]
    header_candidates: List[Tuple[int, str]] = []
    for block in left_header_blocks:
        text = clean_company_candidate(str(block["text"]))
        lowered = text.lower()
        if not text or any(token in lowered for token in HEADER_COMPANY_EXCLUDE_TOKENS):
            continue
        candidate = evaluate_company_text(text)
        if not candidate:
            continue
        short_logo_bonus = 12 if len(candidate) <= 30 and "\n" not in str(block["text"]) else 0
        score = 240 + corp_token_bonus(candidate) + short_logo_bonus - int(float(block["y0"]) * 0.02 + float(block["x0"]) * 0.02)
        header_candidates.append((score, candidate))

    for block in right_header_blocks:
        text = clean_company_candidate(str(block["text"]))
        lowered = text.lower()
        if not text or any(token in lowered for token in HEADER_COMPANY_EXCLUDE_TOKENS):
            continue
        candidate = evaluate_company_text(text, require_corp=True)
        if not candidate:
            continue
        short_logo_bonus = 6 if len(candidate) <= 30 and "\n" not in str(block["text"]) else 0
        score = 210 + corp_token_bonus(candidate) + short_logo_bonus - int(float(block["y0"]) * 0.02 + float(block["x0"]) * 0.02)
        header_candidates.append((score, candidate))

    company_candidates = collect_label_candidates(company_label_patterns, base_score=170)
    vendor_candidates = collect_label_candidates(vendor_label_patterns, base_score=130)
    prioritized = [header_candidates, company_candidates, vendor_candidates]
    for group in prioritized:
        if not group:
            continue
        group.sort(key=lambda item: item[0], reverse=True)
        top_score = group[0][0]
        tied = [item for item in group if abs(item[0] - top_score) <= 8]
        if len(tied) >= 2:
            company_value = ""
        else:
            company_value = group[0][1]
        break

    po_label = find_label_block(blocks, po_label_patterns)
    if po_label is not None:
        around = [get_nearby_value(po_label, blocks)]
        px0 = float(po_label["x0"])
        py0 = float(po_label["y0"])
        py1 = float(po_label["y1"])
        around.extend(
            str(block["text"]).strip()
            for block in blocks
            if block is not po_label
            and float(block["x0"]) >= px0 - 20
            and float(block["x0"]) <= px0 + 520
            and float(block["y0"]) >= py0 - 20
            and float(block["y0"]) <= py1 + 90
        )
        for item in around:
            for match in PO_CODE_PATTERN.findall(item or ""):
                if is_valid_po_number(match):
                    po_candidates.append(clean_order_candidate(match))

    date_label = find_label_block(blocks, date_label_patterns)
    if date_label is not None:
        nearby = get_nearby_value(date_label, blocks)
        if nearby:
            for pattern in DATE_LABEL_PATTERNS:
                match = pattern.search(f"PO Date: {nearby}")
                if match:
                    normalized = normalize_date(match.group(1))
                    if normalized != MISSING_VALUE:
                        date_value = normalized
                        break
            if not date_value:
                direct = DATE_PATTERN.search(nearby)
                if direct:
                    normalized = normalize_date(direct.group(1))
                    if normalized != MISSING_VALUE:
                        date_value = normalized

    return {
        "company_name": clean_company_candidate(company_value) if company_value else "",
        "order_numbers": unique_preserve_order([po for po in po_candidates if is_valid_po_number(po)]),
        "document_date": date_value,
    }



def detect_company_name(
    full_text: str,
    company_rules: List[CompanyRule],
    session_company_memory: Optional[Dict[str, str]] = None,
) -> Tuple[str, str, str, List[str]]:
    normalized_text = normalize_for_match(full_text)
    if company_rules and normalized_text:
        for rule in company_rules:
            for alias in rule.all_names:
                normalized_alias = normalize_for_match(alias)
                if not normalized_alias or is_excluded_company_name(alias):
                    continue
                if normalized_alias in normalized_text:
                    return rule.display_name, alias, rule.source, []

    candidates = collect_auto_company_candidates(full_text)
    if session_company_memory:
        for candidate in candidates:
            learned_name = lookup_company_mapping(session_company_memory, candidate)
            if learned_name and not is_excluded_company_name(learned_name):
                return learned_name, candidate, "session-memory", candidates

    if candidates:
        return candidates[0], candidates[0], "auto-detected", candidates
    return MISSING_VALUE, "", "", []


def lookup_company_mapping(mapping: Dict[str, str], source_name: str) -> str:
    normalized_source = normalize_for_match(source_name)
    if not normalized_source:
        return ""
    for raw_key, mapped in mapping.items():
        if normalize_for_match(str(raw_key)) == normalized_source:
            value = str(mapped).strip()
            if value:
                return value
    return ""


def extract_document_date(full_text: str, filename: str) -> str:
    for pattern in DATE_LABEL_PATTERNS:
        match = pattern.search(full_text)
        if match:
            normalized = normalize_date(match.group(1))
            if normalized != MISSING_VALUE:
                return normalized
    return extract_date_from_filename(filename)


def unique_preserve_order(values: List[str]) -> List[str]:
    seen = set()
    result: List[str] = []
    for value in values:
        key = value.upper()
        if key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result



def extract_order_numbers(full_text: str, company_rule: Optional[CompanyRule] = None) -> List[str]:
    matches: List[str] = []

    if company_rule and company_rule.order_patterns:
        for pattern in company_rule.order_patterns:
            found = pattern.findall(full_text)
            if isinstance(found, list):
                for item in found:
                    if isinstance(item, tuple):
                        matches.extend([part for part in item if part])
                    else:
                        matches.append(item)
        matches = [item for item in matches if item]

    if not matches:
        for pattern in ORDER_LABEL_PATTERNS:
            matches.extend(pattern.findall(full_text))

    cleaned: List[str] = []
    for value in matches:
        candidate = clean_order_candidate(value)
        if is_valid_po_number(candidate):
            cleaned.append(candidate)

    return unique_preserve_order(cleaned)


def collect_raw_order_candidates(full_text: str, company_rule: Optional[CompanyRule] = None) -> List[str]:
    raw_matches: List[str] = []
    if company_rule and company_rule.order_patterns:
        for pattern in company_rule.order_patterns:
            found = pattern.findall(full_text)
            if isinstance(found, list):
                for item in found:
                    if isinstance(item, tuple):
                        raw_matches.extend([part for part in item if part])
                    else:
                        raw_matches.append(item)

    if not raw_matches:
        for pattern in ORDER_LABEL_PATTERNS:
            raw_matches.extend(pattern.findall(full_text))
    cleaned = [clean_order_candidate(value) for value in raw_matches if clean_order_candidate(value)]
    return unique_preserve_order([value for value in cleaned if is_valid_po_number(value)])



def analyze_pdf(pdf_path: Path, company_rules: List[CompanyRule], session_company_memory: Optional[Dict[str, str]] = None) -> DocumentInfo:
    with fitz.open(pdf_path) as document:
        full_text = extract_text(document)
        page_count = len(document)
        used_ocr = False

        first_page = document.load_page(0) if page_count else None
        top_text = ""
        block_result: Dict[str, object] = {"company_name": "", "order_numbers": [], "document_date": ""}
        if first_page is not None:
            rect = first_page.rect
            clip = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.4)
            try:
                top_text = normalize_document_text(first_page.get_text(clip=clip))
            except Exception:
                top_text = ""
            try:
                block_result = extract_from_blocks(first_page)
            except Exception:
                block_result = {"company_name": "", "order_numbers": [], "document_date": ""}

        text_mode_enough = len(re.sub(r"\s+", "", top_text)) >= 120
        working_text = top_text if text_mode_enough else ""
        merged_for_alias = normalize_document_text("\n".join([part for part in [top_text, full_text] if part.strip()]))
        detected_company, matched_alias, company_source, _auto_candidates = detect_company_name(merged_for_alias, company_rules, session_company_memory)
        block_company = str(block_result.get("company_name", "")).strip()
        prefer_detected = detected_company != MISSING_VALUE and (company_source not in {"", "auto-detected"} or not block_company)
        company_name = detected_company if prefer_detected else block_company
        if not company_name:
            company_name = MISSING_VALUE
            matched_alias = ""
            company_source = ""

        block_orders = [value for value in block_result.get("order_numbers", []) if isinstance(value, str)]
        scan_text = normalize_document_text("\n".join(part for part in [working_text, full_text] if part.strip()))
        regex_orders = [clean_order_candidate(match.group(0)) for match in PO_CODE_PATTERN.finditer(scan_text) if is_valid_po_number(match.group(0))]
        order_numbers = unique_preserve_order([*block_orders, *regex_orders])
        filename_orders = extract_po_from_filename(pdf_path.stem)
        if not order_numbers and filename_orders:
            order_numbers = filename_orders[:]

        date_from_blocks = str(block_result.get("document_date", "")).strip()
        if date_from_blocks and normalize_date(date_from_blocks) != MISSING_VALUE:
            document_date = normalize_date(date_from_blocks)
        else:
            document_date = extract_document_date(working_text, pdf_path.stem)

        needs_ocr = (not text_mode_enough) or company_name == MISSING_VALUE or document_date == MISSING_VALUE or not order_numbers

        if needs_ocr and first_page is not None and configure_tesseract():
            ocr_text = perform_ocr_on_top_region(first_page)
            if ocr_text.strip():
                used_ocr = True
                merged_top_text = normalize_document_text("\n".join(part for part in [working_text, ocr_text] if part.strip()))
                detected_company, matched_alias, company_source, _auto_candidates = detect_company_name(merged_top_text, company_rules, session_company_memory)
                if company_name == MISSING_VALUE and detected_company != MISSING_VALUE:
                    company_name = detected_company

                merged_scan_text = normalize_document_text("\n".join(part for part in [merged_top_text, full_text] if part.strip()))
                regex_orders = [
                    clean_order_candidate(match.group(0))
                    for match in PO_CODE_PATTERN.finditer(merged_scan_text)
                    if is_valid_po_number(match.group(0))
                ]
                order_numbers = unique_preserve_order([*order_numbers, *regex_orders])
                if not order_numbers and filename_orders:
                    order_numbers = filename_orders[:]
                if document_date == MISSING_VALUE:
                    document_date = extract_document_date(merged_top_text, pdf_path.stem)

        if company_name == MISSING_VALUE:
            company_source = ""
            matched_alias = ""
        if not order_numbers:
            order_numbers = [MISSING_VALUE]

        raw_order_candidates = order_numbers[:]

    representative_order_number = order_numbers[0] if order_numbers else MISSING_VALUE
    missing_order_only = (not order_numbers) or (len(order_numbers) == 1 and order_numbers[0] == MISSING_VALUE)
    status = "OCR사용" if used_ocr else ("번호없음" if missing_order_only else "분석완료")
    excerpt = " ".join(full_text.split())[:160]
    company_match_status = "회사명매칭성공" if company_name != MISSING_VALUE else "회사명매칭실패"
    if company_name != MISSING_VALUE and company_source == "auto-detected":
        company_match_status = "회사명자동추출"
    elif company_name != MISSING_VALUE and company_source == "session-memory":
        company_match_status = "회사명메모리적용"

    return DocumentInfo(
        pdf_path=pdf_path,
        company_name=company_name,
        document_date=document_date,
        order_numbers=order_numbers,
        representative_order_number=representative_order_number,
        page_count=page_count,
        status=status,
        text_excerpt=excerpt,
        used_ocr=used_ocr,
        company_match_status=company_match_status,
        raw_order_candidates=raw_order_candidates,
        matched_alias=matched_alias,
        company_rule_source=company_source if company_name != MISSING_VALUE else "",
    )


def render_page_to_image(page: fitz.Page) -> Image.Image:
    matrix = fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM)
    pixmap = page.get_pixmap(matrix=matrix, alpha=False)
    return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def fit_image_to_canvas(image: Image.Image) -> Image.Image:
    target_size = LANDSCAPE_SIZE if image.width >= image.height else PORTRAIT_SIZE
    resized = image.copy()
    resized.thumbnail(target_size, Image.LANCZOS)

    canvas = Image.new("RGB", target_size, "white")
    offset_x = (target_size[0] - resized.width) // 2
    offset_y = (target_size[1] - resized.height) // 2
    canvas.paste(resized, (offset_x, offset_y))
    return canvas


def convert_pdf(document_info: DocumentInfo, file_index: int, total_files: int, progress_callback) -> None:
    pdf_path = document_info.pdf_path
    output_dir = pdf_path.parent / f"{pdf_path.stem}_jpg"
    output_dir.mkdir(exist_ok=True)

    company_name = sanitize_filename_part(document_info.company_name)
    document_date = sanitize_filename_part(document_info.document_date)
    order_number = sanitize_filename_part(document_info.representative_order_number)

    with fitz.open(pdf_path) as document:
        total_pages = len(document)
        for page_index in range(total_pages):
            page_number = page_index + 1
            page = document.load_page(page_index)
            page_text = ""
            try:
                page_text = normalize_document_text(page.get_text())
            except Exception:
                page_text = ""
            should_skip_terms = (
                page_index >= 1
                and bool(page_text.strip())
                and is_terms_page(page_text)
            )

            if should_skip_terms:
                progress_callback(
                    ProgressEvent(
                        event_type="page",
                        message=f"{pdf_path.name} {page_number}/{total_pages} 페이지 약관으로 판단되어 JPG 저장 생략",
                        current_file=file_index,
                        total_files=total_files,
                        current_page=page_number,
                        total_pages=total_pages,
                    )
                )
                continue

            progress_callback(
                ProgressEvent(
                    event_type="page",
                    message=(
                        f"{pdf_path.name} 변환 중  |  "
                        f"{page_number}/{total_pages} 페이지  |  "
                        f"{company_name} / {order_number}"
                    ),
                    current_file=file_index,
                    total_files=total_files,
                    current_page=page_number,
                    total_pages=total_pages,
                )
            )

            image = render_page_to_image(page)
            final_image = fit_image_to_canvas(image)
            output_name = f"{company_name}-{document_date}-{order_number}-{page_number}.jpg"
            final_image.save(output_dir / output_name, "JPEG", quality=95)


class PdfToJpgApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        if TkinterDnD is not None:
            try:
                TkinterDnD._require(self)
            except Exception:
                pass
        self.title("PDF to JPG Studio")
        self.geometry("1480x920")
        self.minsize(1260, 820)
        try:
            self.state("zoomed")
        except Exception:
            pass

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.script_dir = Path(__file__).resolve().parent
        self.companies_path = self.script_dir / "companies.txt"
        self.company_rules_csv_path = self.script_dir / "companies_rules.csv"
        self.company_mapping_path = self.script_dir / "company_mapping.json"
        self.selected_folder: Optional[Path] = None
        self.selected_inputs: List[Path] = []
        self.worker_thread: Optional[threading.Thread] = None
        self.event_queue: "queue.Queue[ProgressEvent]" = queue.Queue()
        self.is_running = False
        self.documents: List[DocumentInfo] = []
        self.selected_company: Optional[str] = None
        self.selection_vars: Dict[str, tk.BooleanVar] = {}
        self.selection_checkboxes: Dict[str, ctk.CTkCheckBox] = {}
        self.selection_order: List[str] = []
        self.selection_meta: Dict[str, Tuple[str, str]] = {}
        self.detected_text = "감지된 회사/발주번호가 여기에 표시됩니다."
        self.preview_text = tk.StringVar(value="선택한 발주번호가 여기에 표시됩니다.")
        self.filter_mode_var = tk.StringVar(value="전체")
        self.filter_value_var = tk.StringVar(value="전체")
        self.is_left_panel_collapsed = False
        self.session_company_memory: Dict[str, str] = self.load_persistent_company_memory()

        self._build_ui()
        self.after(100, self.process_event_queue)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        shell = ctk.CTkFrame(self, fg_color="#f6efe7", corner_radius=0)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.grid_columnconfigure(0, weight=4)
        shell.grid_columnconfigure(1, weight=5)
        shell.grid_rowconfigure(2, weight=1)

        hero = ctk.CTkFrame(shell, fg_color="#f7cfb9", corner_radius=28)
        hero.grid(row=0, column=0, columnspan=2, padx=24, pady=(24, 14), sticky="ew")
        hero.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            hero,
            text="PDF to JPG Studio",
            font=ctk.CTkFont(family="Malgun Gothic", size=30, weight="bold"),
            text_color="#3f2f2b",
        ).grid(row=0, column=0, padx=24, pady=(22, 6), sticky="w")

        ctk.CTkLabel(
            hero,
            text=(
                "PDF를 분석해서 회사명, 발주일, 발주번호를 정리하고\n"
                "같은 회사 번호만 골라 기안 제목을 바로 복사할 수 있게 준비해둡니다."
            ),
            font=ctk.CTkFont(family="Malgun Gothic", size=14),
            text_color="#5d4a45",
            justify="left",
        ).grid(row=1, column=0, padx=24, pady=(0, 22), sticky="w")

        action_card = ctk.CTkFrame(shell, fg_color="#fffaf6", corner_radius=24)
        action_card.grid(row=1, column=0, columnspan=2, padx=24, pady=(0, 14), sticky="ew")
        action_card.grid_columnconfigure(0, weight=1)
        for column in range(1, 6):
            action_card.grid_columnconfigure(column, weight=0)

        self.folder_label = ctk.CTkLabel(
            action_card,
            text="아직 작업 폴더가 선택되지 않았어요.",
            font=ctk.CTkFont(family="Malgun Gothic", size=15, weight="bold"),
            text_color="#4a3f35",
            anchor="w",
        )
        self.folder_label.grid(row=0, column=0, padx=(20, 16), pady=(18, 6), sticky="ew")

        self.info_label = ctk.CTkLabel(
            action_card,
            text="폴더를 고르거나 PDF 파일을 끌어다 놓으면, 오른쪽에 회사별 발주번호 목록이 채워집니다.",
            font=ctk.CTkFont(family="Malgun Gothic", size=13),
            text_color="#7b6c61",
            justify="left",
            anchor="w",
        )
        self.info_label.grid(row=1, column=0, padx=(20, 16), pady=(0, 18), sticky="ew")

        self.select_button = self._make_action_button(action_card, "폴더 선택", self.select_folder, "#e07a5f", "#d26449")
        self.select_button.grid(row=0, column=1, rowspan=2, padx=6, pady=18)

        self.file_button = self._make_action_button(action_card, "파일 선택", self.select_files, "#d98c5f", "#c6764a")
        self.file_button.grid(row=0, column=2, rowspan=2, padx=6, pady=18)

        self.analyze_button = self._make_action_button(action_card, "문서 분석", self.start_analysis, "#6e9f87", "#5b8b75")
        self.analyze_button.grid(row=0, column=3, rowspan=2, padx=6, pady=18)

        self.convert_button = self._make_action_button(action_card, "JPG 변환", self.start_conversion, "#4b6cb7", "#3e5c9d")
        self.convert_button.grid(row=0, column=4, rowspan=2, padx=(6, 20), pady=18)

        self.export_button = self._make_action_button(action_card, "요약 내보내기", self.export_summary, "#7d6bb3", "#6b5b9c")
        self.export_button.grid(row=0, column=5, rowspan=2, padx=(0, 12), pady=18)

        drop_message = "여기로 PDF 파일이나 폴더를 드래그해도 됩니다."
        if DND_FILES is None:
            drop_message += "  드래그 기능을 쓰려면 `py -m pip install tkinterdnd2`"

        self.drop_label = ctk.CTkLabel(
            action_card,
            text=drop_message,
            font=ctk.CTkFont(family="Malgun Gothic", size=12),
            text_color="#8a7767",
            corner_radius=14,
            fg_color="#fff2e7",
            height=36,
        )
        self.drop_label.grid(row=2, column=0, columnspan=5, padx=20, pady=(0, 18), sticky="ew")

        self.global_toggle_button = ctk.CTkButton(
            action_card,
            text="작업현황 접기",
            command=self.toggle_left_panel,
            width=120,
            height=34,
            corner_radius=14,
            fg_color="#d9b08c",
            hover_color="#c69b76",
            text_color="#fffaf6",
            font=ctk.CTkFont(family="Malgun Gothic", size=12, weight="bold"),
        )
        self.global_toggle_button.grid(row=0, column=6, rowspan=2, padx=(0, 20), pady=18)
        self.setup_drag_and_drop()

        self.left_panel = ctk.CTkFrame(shell, fg_color="#fffaf6", corner_radius=24)
        self.left_panel.grid(row=2, column=0, padx=(24, 10), pady=(0, 24), sticky="nsew")
        self.left_panel.grid_columnconfigure(0, weight=1)
        self.left_panel.grid_columnconfigure(1, weight=0)
        self.left_panel.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            self.left_panel,
            text="작업 현황",
            font=ctk.CTkFont(family="Malgun Gothic", size=22, weight="bold"),
            text_color="#40342c",
        ).grid(row=0, column=0, padx=20, pady=(20, 8), sticky="w")

        self.toggle_left_panel_button = ctk.CTkButton(
            self.left_panel,
            text="접기",
            command=self.toggle_left_panel,
            width=80,
            height=34,
            corner_radius=14,
            fg_color="#d9b08c",
            hover_color="#c69b76",
            text_color="#fffaf6",
            font=ctk.CTkFont(family="Malgun Gothic", size=13, weight="bold"),
        )
        self.toggle_left_panel_button.grid(row=0, column=1, padx=(0, 20), pady=(16, 8), sticky="e")

        self.status_label = ctk.CTkLabel(
            self.left_panel,
            text="대기 중입니다. 폴더를 고르고 문서 분석을 눌러주세요.",
            font=ctk.CTkFont(family="Malgun Gothic", size=14),
            text_color="#6d6156",
            justify="left",
            wraplength=620,
            anchor="w",
        )
        self.status_label.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 12), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(
            self.left_panel,
            height=18,
            corner_radius=20,
            progress_color="#81b29a",
            fg_color="#eadfd2",
        )
        self.progress_bar.grid(row=2, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="ew")
        self.progress_bar.set(0)

        self.progress_detail_label = ctk.CTkLabel(
            self.left_panel,
            text="준비 완료",
            font=ctk.CTkFont(family="Malgun Gothic", size=13),
            text_color="#7b6c61",
        )
        self.progress_detail_label.grid(row=3, column=0, columnspan=2, padx=20, pady=(0, 14), sticky="w")

        self.log_textbox = ctk.CTkTextbox(
            self.left_panel,
            corner_radius=18,
            fg_color="#fff4eb",
            text_color="#4a4038",
            font=ctk.CTkFont(family="Consolas", size=12),
            border_width=0,
        )
        self.log_textbox.grid(row=4, column=0, columnspan=2, padx=20, pady=(0, 20), sticky="nsew")
        self.log_textbox.insert("end", "준비 완료. `companies.txt` 또는 `companies_rules.csv`로 거래처 규칙을 넣을 수 있습니다.\n")
        self.log_textbox.configure(state="disabled")

        self.right_panel = ctk.CTkFrame(shell, fg_color="#fffaf6", corner_radius=24)
        self.right_panel.grid(row=2, column=1, padx=(10, 24), pady=(0, 24), sticky="nsew")
        self.right_panel.grid_columnconfigure(0, weight=1)
        self.right_panel.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            self.right_panel,
            text="기안 제목 만들기",
            font=ctk.CTkFont(family="Malgun Gothic", size=22, weight="bold"),
            text_color="#40342c",
        ).grid(row=0, column=0, padx=20, pady=(20, 8), sticky="w")

        summary_row = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        summary_row.grid(row=1, column=0, padx=20, pady=(0, 6), sticky="ew")
        summary_row.grid_columnconfigure(0, weight=1)
        summary_row.grid_columnconfigure(1, weight=0)

        self.summary_label = ctk.CTkLabel(
            summary_row,
            text="분석 전",
            font=ctk.CTkFont(family="Malgun Gothic", size=14),
            text_color="#7b6c61",
            justify="left",
            anchor="w",
        )
        self.summary_label.grid(row=0, column=0, sticky="ew")

        self.selection_hint_label = ctk.CTkLabel(
            summary_row,
            text="번호를 체크하면 아래 제목이 바로 만들어집니다.",
            font=ctk.CTkFont(family="Malgun Gothic", size=12),
            text_color="#9a8b7f",
            anchor="e",
        )
        self.selection_hint_label.grid(row=0, column=1, padx=(8, 0), sticky="e")

        filter_row = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        filter_row.grid(row=2, column=0, padx=20, pady=(0, 8), sticky="ew")
        filter_row.grid_columnconfigure(1, weight=0)
        filter_row.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(
            filter_row,
            text="날짜 필터",
            font=ctk.CTkFont(family="Malgun Gothic", size=12, weight="bold"),
            text_color="#7b6c61",
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")

        self.filter_mode_menu = ctk.CTkOptionMenu(
            filter_row,
            values=["전체", "일간", "주간", "월간"],
            variable=self.filter_mode_var,
            command=self.on_filter_mode_changed,
            width=110,
            height=32,
        )
        self.filter_mode_menu.grid(row=0, column=1, padx=(0, 8), sticky="w")

        self.filter_value_menu = ctk.CTkOptionMenu(
            filter_row,
            values=["전체"],
            variable=self.filter_value_var,
            command=self.on_filter_value_changed,
            width=180,
            height=32,
        )
        self.filter_value_menu.grid(row=0, column=2, sticky="w")

        self.selection_frame = ctk.CTkScrollableFrame(
            self.right_panel,
            fg_color="#fff6ef",
            corner_radius=18,
            label_text="회사별 발주번호 목록",
            label_font=ctk.CTkFont(family="Malgun Gothic", size=14, weight="bold"),
            label_fg_color="#fff6ef",
        )
        self.selection_frame.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="nsew")
        self.selection_frame.grid_columnconfigure(0, weight=1)

        preview_frame = ctk.CTkFrame(self.right_panel, fg_color="#fff2e7", corner_radius=18)
        preview_frame.grid(row=4, column=0, padx=20, pady=(0, 16), sticky="ew")
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(1, weight=0)
        preview_frame.grid_columnconfigure(2, weight=0)
        preview_frame.grid_columnconfigure(3, weight=0)

        ctk.CTkLabel(
            preview_frame,
            text="복붙용 제목",
            font=ctk.CTkFont(family="Malgun Gothic", size=12, weight="bold"),
            text_color="#5b4a44",
        ).grid(row=0, column=0, columnspan=3, padx=14, pady=(10, 3), sticky="w")

        self.preview_entry = ctk.CTkEntry(
            preview_frame,
            height=40,
            corner_radius=14,
            fg_color="#fffaf6",
            text_color="#352d29",
            font=ctk.CTkFont(family="Malgun Gothic", size=13),
        )
        self.preview_entry.grid(row=1, column=0, padx=(14, 10), pady=(0, 10), sticky="ew")
        self.preview_entry.insert(0, self.preview_text.get())
        self.preview_entry.configure(state="readonly")

        self.copy_button = ctk.CTkButton(
            preview_frame,
            text="제목 복사",
            command=self.copy_preview,
            width=120,
            height=40,
            corner_radius=14,
            fg_color="#e9a03b",
            hover_color="#d18e2e",
            text_color="#fffaf6",
            font=ctk.CTkFont(family="Malgun Gothic", size=13, weight="bold"),
        )
        self.copy_button.grid(row=1, column=1, padx=(0, 8), pady=(0, 10), sticky="ew")

        self.clear_button = ctk.CTkButton(
            preview_frame,
            text="초기화",
            command=self.clear_selection,
            width=100,
            height=40,
            corner_radius=14,
            fg_color="#c97b63",
            hover_color="#b56750",
            text_color="#fffaf6",
            font=ctk.CTkFont(family="Malgun Gothic", size=13, weight="bold"),
        )
        self.clear_button.grid(row=1, column=2, padx=(0, 14), pady=(0, 10), sticky="ew")

        self.memory_export_button = ctk.CTkButton(
            preview_frame,
            text="회사명 기억 복사",
            command=self.export_session_memory,
            width=120,
            height=34,
            corner_radius=14,
            fg_color="#7c8db5",
            hover_color="#6c7da2",
            text_color="#fffaf6",
            font=ctk.CTkFont(family="Malgun Gothic", size=12, weight="bold"),
        )
        self.memory_export_button.grid(row=2, column=1, padx=(0, 8), pady=(0, 12), sticky="ew")

        self.memory_import_button = ctk.CTkButton(
            preview_frame,
            text="회사명 기억 붙여넣기",
            command=self.open_memory_import_dialog,
            width=140,
            height=34,
            corner_radius=14,
            fg_color="#6e9f87",
            hover_color="#5b8b75",
            text_color="#fffaf6",
            font=ctk.CTkFont(family="Malgun Gothic", size=12, weight="bold"),
        )
        self.memory_import_button.grid(row=2, column=2, padx=(0, 14), pady=(0, 12), sticky="ew")

        self.memory_status_label = ctk.CTkLabel(
            preview_frame,
            text="회사명 기억 0개 | 파일 저장 없이 현재 실행 중에만 유지",
            font=ctk.CTkFont(family="Malgun Gothic", size=11),
            text_color="#8a7767",
            justify="left",
            anchor="w",
        )
        self.memory_status_label.grid(row=2, column=0, padx=(14, 10), pady=(0, 12), sticky="ew")

        self._populate_empty_selection_state()
        self.update_memory_status()

    def _make_action_button(self, parent, text: str, command, fg_color: str, hover_color: str) -> ctk.CTkButton:
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=130,
            height=42,
            corner_radius=18,
            fg_color=fg_color,
            hover_color=hover_color,
            text_color="#fffaf6",
            font=ctk.CTkFont(family="Malgun Gothic", size=14, weight="bold"),
        )

    def toggle_left_panel(self) -> None:
        self.is_left_panel_collapsed = not self.is_left_panel_collapsed

        if self.is_left_panel_collapsed:
            self.left_panel.grid_remove()
            self.right_panel.grid_configure(row=2, column=0, columnspan=2, padx=24, pady=(0, 24), sticky="nsew")
            self.toggle_left_panel_button.configure(text="펼치기")
            self.global_toggle_button.configure(text="작업현황 펼치기")
        else:
            self.left_panel.grid()
            self.right_panel.grid_configure(row=2, column=1, columnspan=1, padx=(10, 24), pady=(0, 24), sticky="nsew")
            self.toggle_left_panel_button.configure(text="접기")
            self.global_toggle_button.configure(text="작업현황 접기")

    def on_filter_mode_changed(self, _choice: str) -> None:
        self.refresh_filter_values()
        self.refresh_selection_panel()

    def on_filter_value_changed(self, _choice: str) -> None:
        self.refresh_selection_panel()

    def get_filter_values(self) -> List[str]:
        dates = sorted({doc.document_date for doc in self.documents if doc.document_date != MISSING_VALUE})
        mode = self.filter_mode_var.get()
        if mode == "일간":
            return ["전체"] + dates
        if mode == "주간":
            weeks = sorted({self.get_week_label(date_text) for date_text in dates})
            return ["전체"] + weeks
        if mode == "월간":
            months = sorted({date_text[:7] for date_text in dates})
            return ["전체"] + months
        return ["전체"]

    def refresh_filter_values(self) -> None:
        values = self.get_filter_values()
        current = self.filter_value_var.get()
        self.filter_value_menu.configure(values=values)
        if current not in values:
            self.filter_value_var.set("전체")

    def get_week_label(self, date_text: str) -> str:
        parsed = datetime.strptime(date_text, "%Y-%m-%d")
        iso_year, iso_week, _weekday = parsed.isocalendar()
        return f"{iso_year}-W{iso_week:02d}"

    def get_filtered_documents(self) -> List[DocumentInfo]:
        mode = self.filter_mode_var.get()
        filter_value = self.filter_value_var.get()
        if mode == "전체" or filter_value == "전체":
            return self.documents

        filtered: List[DocumentInfo] = []
        for doc in self.documents:
            if doc.document_date == MISSING_VALUE:
                continue
            if mode == "일간" and doc.document_date == filter_value:
                filtered.append(doc)
            elif mode == "주간" and self.get_week_label(doc.document_date) == filter_value:
                filtered.append(doc)
            elif mode == "월간" and doc.document_date.startswith(filter_value):
                filtered.append(doc)
        return filtered

    def open_edit_dialog(self, document: DocumentInfo) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title("행 편집")
        dialog.geometry("420x260")
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(1, weight=1)

        company_var = tk.StringVar(value=document.company_name)
        date_var = tk.StringVar(value=document.document_date)
        order_var = tk.StringVar(value=document.representative_order_number)

        fields = [
            ("회사명", company_var),
            ("날짜", date_var),
            ("발주번호", order_var),
        ]
        for index, (label_text, variable) in enumerate(fields):
            ctk.CTkLabel(dialog, text=label_text, font=ctk.CTkFont(family="Malgun Gothic", size=13, weight="bold")).grid(
                row=index, column=0, padx=(18, 10), pady=(18 if index == 0 else 8, 0), sticky="w"
            )
            ctk.CTkEntry(dialog, textvariable=variable, height=34).grid(
                row=index, column=1, padx=(0, 18), pady=(18 if index == 0 else 8, 0), sticky="ew"
            )

        def save_edit() -> None:
            previous_candidate = (document.matched_alias or document.company_name).strip()
            previous_company_name = document.company_name
            document.company_name = company_var.get().strip() or MISSING_VALUE
            document.document_date = date_var.get().strip() or MISSING_VALUE
            document.representative_order_number = order_var.get().strip() or MISSING_VALUE
            document.order_numbers = [document.representative_order_number]
            if previous_candidate and document.company_name != MISSING_VALUE:
                self.remember_company_mapping(previous_candidate, document.company_name)
                document.matched_alias = previous_candidate
                document.company_rule_source = "session-memory"
                if previous_company_name != document.company_name:
                    document.company_match_status = "회사명수동확정"
            self.append_log(
                f"[수정] {document.pdf_path.name} | 회사: {document.company_name} | 날짜: {document.document_date} | 번호: {document.representative_order_number}"
            )
            dialog.destroy()
            self.refresh_filter_values()
            self.refresh_selection_panel()

        button_row = ctk.CTkFrame(dialog, fg_color="transparent")
        button_row.grid(row=4, column=0, columnspan=2, padx=18, pady=20, sticky="ew")
        button_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(button_row, text="저장", command=save_edit).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkButton(button_row, text="취소", command=dialog.destroy, fg_color="#b0a89f", hover_color="#9b938a").grid(
            row=0, column=1, padx=(8, 0), sticky="ew"
        )

    def update_memory_status(self) -> None:
        count = len(self.session_company_memory)
        storage_text = f"JSON 저장 ({self.company_mapping_path.name})"
        self.memory_status_label.configure(text=f"회사명 기억 {count}개 | {storage_text}")

    def remember_company_mapping(self, detected_name: str, confirmed_name: str) -> None:
        detected = detected_name.strip()
        confirmed = confirmed_name.strip()
        if not detected or not confirmed or detected == MISSING_VALUE or confirmed == MISSING_VALUE:
            return
        if is_excluded_company_name(detected) or is_excluded_company_name(confirmed):
            return
        self.session_company_memory[detected] = confirmed
        self.save_persistent_company_memory()
        self.update_memory_status()
        self.append_log(f"[기억] {detected} -> {confirmed}")

    def build_memory_export_text(self) -> str:
        lines = [
            f"{key}={value}"
            for key, value in sorted(self.session_company_memory.items(), key=lambda item: item[1].lower())
            if value.strip()
        ]
        return "\n".join(lines)

    def export_session_memory(self) -> None:
        text = self.build_memory_export_text()
        if not text:
            messagebox.showwarning("기억 없음", "복사할 회사명 기억이 아직 없습니다. 문서 수정 후 다시 시도해주세요.")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.save_last_memory_export(text)
        self.append_log(f"[복사] 회사명 기억 {len(self.session_company_memory)}개")
        messagebox.showinfo("복사 완료", "회사명 기억 목록을 클립보드에 복사했습니다. 메모장이나 카톡에 붙여 넣어 보관하세요.")

    def import_session_memory_text(self, raw_text: str) -> Tuple[int, int]:
        added = 0
        skipped = 0
        for line in raw_text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                skipped += 1
                continue
            key, value = line.split("=", 1)
            raw_key = key.strip()
            company_name = value.strip()
            if not raw_key:
                skipped += 1
                continue
            if company_name in {"-", "__DELETE__", "삭제"}:
                if raw_key in self.session_company_memory:
                    del self.session_company_memory[raw_key]
                    added += 1
                else:
                    skipped += 1
                continue
            if not company_name or is_excluded_company_name(company_name):
                skipped += 1
                continue
            self.session_company_memory[raw_key] = company_name
            added += 1
        self.save_persistent_company_memory()
        self.update_memory_status()
        return added, skipped

    def open_memory_import_dialog(self) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title("회사명 기억 붙여넣기")
        dialog.geometry("560x380")
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            dialog,
            text="복사해둔 회사명 기억을 붙여넣으세요. 형식: 감지값=확정회사명 (삭제: 감지값=-)",
            font=ctk.CTkFont(family="Malgun Gothic", size=13, weight="bold"),
            text_color="#4a3f35",
            justify="left",
            wraplength=500,
        ).grid(row=0, column=0, padx=18, pady=(18, 8), sticky="w")

        textbox = ctk.CTkTextbox(dialog, corner_radius=14, fg_color="#fffaf6", text_color="#352d29")
        textbox.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="nsew")
        sample = self.build_memory_export_text()
        if sample:
            textbox.insert("end", sample)

        def apply_import() -> None:
            raw_text = textbox.get("1.0", "end").strip()
            if not raw_text:
                messagebox.showwarning("입력 필요", "붙여넣을 내용을 입력해주세요.")
                return
            added, skipped = self.import_session_memory_text(raw_text)
            self.append_log(f"[불러오기] 회사명 기억 추가 {added}개 | 건너뜀 {skipped}개")
            dialog.destroy()
            messagebox.showinfo("불러오기 완료", f"회사명 기억 {added}개를 반영했습니다. 건너뜀: {skipped}개")

        button_row = ctk.CTkFrame(dialog, fg_color="transparent")
        button_row.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="ew")
        button_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(button_row, text="반영", command=apply_import).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkButton(button_row, text="닫기", command=dialog.destroy, fg_color="#b0a89f", hover_color="#9b938a").grid(
            row=0, column=1, padx=(8, 0), sticky="ew"
        )

    def load_persistent_company_memory(self) -> Dict[str, str]:
        memory: Dict[str, str] = {}
        try:
            if not self.company_mapping_path.exists():
                self.company_mapping_path.write_text("{}", encoding="utf-8")
            loaded = json.loads(self.company_mapping_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                for key, value in loaded.items():
                    raw_key = str(key).strip()
                    mapped_name = str(value).strip()
                    if raw_key and mapped_name:
                        memory[raw_key] = mapped_name
        except Exception:
            pass
        return memory

    def save_persistent_company_memory(self) -> None:
        raw_json = json.dumps(self.session_company_memory, ensure_ascii=False, indent=2)
        self.company_mapping_path.write_text(raw_json, encoding="utf-8")

    def save_last_memory_export(self, export_text: str) -> None:
        if not export_text.strip() or winreg is None:
            return
        try:
            registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_BASE_KEY)
            winreg.SetValueEx(registry_key, REGISTRY_EXPORT_VALUE, 0, winreg.REG_SZ, export_text)
            winreg.CloseKey(registry_key)
        except Exception:
            pass

    def open_alias_register_dialog(self) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title("거래처명 등록")
        dialog.geometry("520x260")
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(1, weight=1)

        po_company_var = tk.StringVar()
        alias_var = tk.StringVar()

        selected_keys = [key for key in self.selection_order if self.selection_vars.get(key) and self.selection_vars[key].get()]
        if selected_keys:
            selected_company, _selected_order = self.selection_meta[selected_keys[0]]
            if selected_company and selected_company != MISSING_VALUE:
                po_company_var.set(selected_company)

        ctk.CTkLabel(dialog, text="PO상 사명", font=ctk.CTkFont(family="Malgun Gothic", size=13, weight="bold")).grid(
            row=0, column=0, padx=(18, 10), pady=(18, 8), sticky="w"
        )
        ctk.CTkEntry(dialog, textvariable=po_company_var, height=36).grid(
            row=0, column=1, padx=(0, 18), pady=(18, 8), sticky="ew"
        )

        ctk.CTkLabel(dialog, text="별칭", font=ctk.CTkFont(family="Malgun Gothic", size=13, weight="bold")).grid(
            row=1, column=0, padx=(18, 10), pady=8, sticky="w"
        )
        ctk.CTkEntry(dialog, textvariable=alias_var, height=36).grid(
            row=1, column=1, padx=(0, 18), pady=8, sticky="ew"
        )

        ctk.CTkLabel(
            dialog,
            text="등록하면 다음부터 PO상 사명으로 읽혀도 별칭으로 자동 치환됩니다.",
            font=ctk.CTkFont(family="Malgun Gothic", size=12),
            text_color="#7b6c61",
            justify="left",
            wraplength=470,
        ).grid(row=2, column=0, columnspan=2, padx=18, pady=(4, 12), sticky="w")

        def apply_register() -> None:
            detected = po_company_var.get().strip()
            alias = alias_var.get().strip()
            if not detected or not alias:
                messagebox.showwarning("입력 필요", "PO상 사명과 별칭을 모두 입력해주세요.")
                return
            self.remember_company_mapping(detected, alias)
            for document in self.documents:
                candidate_names = [document.company_name, document.matched_alias]
                if detected in candidate_names or normalize_for_match(document.company_name) == normalize_for_match(detected):
                    document.company_name = alias
                    document.company_rule_source = "registry-memory" if winreg is not None else "local-json-memory"
            self.refresh_selection_panel()
            dialog.destroy()
            messagebox.showinfo("등록 완료", "거래처명을 내부에 기록했습니다. 다음 분석부터 자동 적용됩니다.")

        button_row = ctk.CTkFrame(dialog, fg_color="transparent")
        button_row.grid(row=3, column=0, columnspan=2, padx=18, pady=(0, 18), sticky="ew")
        button_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(button_row, text="등록", command=apply_register).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkButton(button_row, text="닫기", command=dialog.destroy, fg_color="#b0a89f", hover_color="#9b938a").grid(
            row=0, column=1, padx=(8, 0), sticky="ew"
        )

    def _populate_empty_selection_state(self) -> None:
        for child in self.selection_frame.winfo_children():
            child.destroy()

        self.set_detected_text("감지된 회사/발주번호가 여기에 표시됩니다.")

        ctk.CTkLabel(
            self.selection_frame,
            text="문서 분석을 실행하면 회사별 발주번호 후보가 아래 체크박스로 나타납니다.",
            font=ctk.CTkFont(family="Malgun Gothic", size=12),
            text_color="#7b6c61",
            justify="left",
            wraplength=520,
        ).grid(row=0, column=0, padx=10, pady=(8, 8), sticky="w")

    def append_log(self, message: str) -> None:
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def set_preview_text(self, text: str) -> None:
        self.preview_entry.configure(state="normal")
        self.preview_entry.delete(0, "end")
        self.preview_entry.insert(0, text)
        self.preview_entry.configure(state="readonly")
        self.preview_text.set(text)

    def set_detected_text(self, text: str) -> None:
        self.detected_text = text

    def setup_drag_and_drop(self) -> None:
        if DND_FILES is None:
            return

        for widget in (self, self.drop_label):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self.handle_drop)
            except Exception:
                continue

    def parse_drop_paths(self, raw_data: str) -> List[Path]:
        try:
            items = self.tk.splitlist(raw_data)
        except tk.TclError:
            items = [raw_data]
        return [Path(item.strip("{}")) for item in items if item]

    def set_selected_inputs(self, paths: List[Path], source_label: str) -> None:
        deduped: List[Path] = []
        seen = set()
        for path in paths:
            resolved = Path(path)
            key = str(resolved).lower()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(resolved)

        self.selected_inputs = deduped
        self.selected_folder = deduped[0] if len(deduped) == 1 and deduped[0].is_dir() else None
        label = ", ".join(path.name for path in deduped[:3])
        if len(deduped) > 3:
            label += f" 외 {len(deduped) - 3}개"
        self.folder_label.configure(text=f"선택한 대상: {label}")
        self.info_label.configure(text=f"{source_label} 완료. 이제 문서 분석을 눌러서 회사별 발주번호 후보를 확인해보세요.")
        self.append_log(f"[선택] {source_label}: {', '.join(str(path) for path in deduped)}")

    def collect_pdf_files(self) -> List[Path]:
        pdf_files: List[Path] = []
        seen = set()
        for item in self.selected_inputs:
            if item.is_dir():
                candidates = sorted(
                    child for child in item.iterdir()
                    if child.is_file() and child.suffix.lower() in SUPPORTED_EXTENSIONS
                )
            elif item.is_file() and item.suffix.lower() in SUPPORTED_EXTENSIONS:
                candidates = [item]
            else:
                candidates = []

            for candidate in candidates:
                key = str(candidate.resolve()).lower()
                if key in seen:
                    continue
                seen.add(key)
                pdf_files.append(candidate)
        return sorted(pdf_files)

    def handle_drop(self, event) -> None:
        dropped_paths = [path for path in self.parse_drop_paths(event.data) if path.exists()]
        if not dropped_paths:
            messagebox.showwarning("드롭 실패", "유효한 PDF 파일이나 폴더를 찾지 못했습니다.")
            return
        self.set_selected_inputs(dropped_paths, "드래그 앤 드롭")

    def select_folder(self) -> None:
        folder = filedialog.askdirectory(title="PDF 파일이 있는 폴더를 선택하세요")
        if not folder:
            return

        self.set_selected_inputs([Path(folder)], "폴더 선택")

    def select_files(self) -> None:
        file_paths = filedialog.askopenfilenames(
            title="PDF 파일을 선택하세요",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not file_paths:
            return
        self.set_selected_inputs([Path(path) for path in file_paths], "파일 선택")

    def set_running_state(self, is_running: bool) -> None:
        self.is_running = is_running
        state = "disabled" if is_running else "normal"
        self.select_button.configure(state=state)
        self.file_button.configure(state=state)
        self.analyze_button.configure(state=state)
        self.convert_button.configure(state=state)
        self.export_button.configure(state=state)

    def start_analysis(self) -> None:
        if self.is_running:
            return
        if not self.selected_inputs:
            messagebox.showwarning("선택 필요", "먼저 PDF 파일이나 폴더를 선택해주세요.")
            return

        self.progress_bar.set(0)
        self.status_label.configure(text="문서 분석 준비 중...")
        self.progress_detail_label.configure(text="PDF 목록 확인 중")
        self.summary_label.configure(text="분석 중")
        self.documents = []
        self.filter_mode_var.set("전체")
        self.filter_value_var.set("전체")
        self.clear_selection(reset_documents=False)
        self.append_log("[시작] PDF 분석을 시작합니다.")
        self.set_running_state(True)
        self.worker_thread = threading.Thread(target=self.run_analysis, daemon=True)
        self.worker_thread.start()

    def run_analysis(self) -> None:
        try:
            company_rules = load_company_rules(self.companies_path)
            if company_rules:
                source_names = sorted({rule.source for rule in company_rules})
                self.event_queue.put(
                    ProgressEvent(
                        event_type="status",
                        message=f"거래처 규칙 {len(company_rules)}개를 불러왔습니다. (출처: {', '.join(source_names)})",
                    )
                )
            else:
                self.event_queue.put(ProgressEvent(event_type="status", message="companies.txt/companies_rules.csv가 없어도 자동 회사명 추출을 시도합니다. 규칙 파일이 있으면 더 정확합니다."))

            if configure_tesseract():
                self.event_queue.put(ProgressEvent(event_type="status", message="Tesseract OCR이 연결되어 스캔 PDF도 보조 분석합니다."))
            else:
                self.event_queue.put(ProgressEvent(event_type="status", message="Tesseract OCR을 찾지 못해 이미지형 PDF는 추출이 제한될 수 있습니다."))

            pdf_files = self.collect_pdf_files()
            total_files = len(pdf_files)

            if not pdf_files:
                self.event_queue.put(ProgressEvent(event_type="done", message="선택한 항목에서 PDF 파일을 찾지 못했습니다."))
                return

            documents: List[DocumentInfo] = []
            success_count = 0
            fail_count = 0

            for file_index, pdf_path in enumerate(pdf_files, start=1):
                self.event_queue.put(
                    ProgressEvent(
                        event_type="status",
                        message=f"{pdf_path.name} 분석 중...",
                        current_file=file_index,
                        total_files=total_files,
                    )
                )
                try:
                    document_info = analyze_pdf(pdf_path, company_rules, self.session_company_memory.copy())
                    mapped_company = lookup_company_mapping(self.session_company_memory, document_info.company_name)
                    if mapped_company:
                        document_info.company_name = mapped_company
                        document_info.company_rule_source = "company-mapping-json"
                    documents.append(document_info)
                    success_count += 1
                    source_label = "OCR 보강" if document_info.used_ocr else "일반 추출"
                    order_debug = ", ".join(document_info.raw_order_candidates) if document_info.raw_order_candidates else "없음"
                    matched_alias_text = document_info.matched_alias or "없음"
                    rule_source_text = document_info.company_rule_source or "기본패턴"
                    log_message = (
                        f"[분석] {pdf_path.name} | {source_label} | {document_info.company_match_status} | "
                        f"회사: {document_info.company_name} | 매칭명: {matched_alias_text} | 규칙: {rule_source_text} | "
                        f"날짜: {document_info.document_date} | "
                        f"번호: {', '.join(document_info.order_numbers) if document_info.order_numbers else '없음'} | "
                        f"번호후보: {order_debug}"
                    )
                    self.event_queue.put(
                        ProgressEvent(
                            event_type="status",
                            message=log_message,
                            current_file=file_index,
                            total_files=total_files,
                        )
                    )
                except Exception as error:
                    fail_count += 1
                    self.event_queue.put(
                        ProgressEvent(
                            event_type="status",
                            message=f"[오류] {pdf_path.name} 분석 실패: {error}",
                            current_file=file_index,
                            total_files=total_files,
                        )
                    )

                self.event_queue.put(
                    ProgressEvent(
                        event_type="summary",
                        total_files=total_files,
                        success_count=success_count,
                        fail_count=fail_count,
                        current_file=file_index,
                    )
                )

            self.event_queue.put(
                ProgressEvent(
                    event_type="analysis_complete",
                    message=f"분석 완료: 성공 {success_count}개 / 실패 {fail_count}개",
                    total_files=total_files,
                    success_count=success_count,
                    fail_count=fail_count,
                    documents=documents,
                )
            )
        except Exception as error:
            self.event_queue.put(ProgressEvent(event_type="done", message=f"분석 중 오류가 발생했습니다: {error}"))

    def start_conversion(self) -> None:
        if self.is_running:
            return
        if not self.documents:
            messagebox.showwarning("분석 필요", "먼저 문서 분석을 완료해주세요.")
            return

        self.progress_bar.set(0)
        self.status_label.configure(text="JPG 변환 준비 중...")
        self.progress_detail_label.configure(text="출력 폴더 준비 중")
        self.append_log("[시작] JPG 변환을 시작합니다.")
        self.set_running_state(True)
        self.worker_thread = threading.Thread(target=self.run_conversion, daemon=True)
        self.worker_thread.start()

    def run_conversion(self) -> None:
        success_count = 0
        fail_count = 0
        total_files = len(self.documents)

        try:
            for file_index, document_info in enumerate(self.documents, start=1):
                try:
                    convert_pdf(document_info, file_index, total_files, self.event_queue.put)
                    success_count += 1
                    self.event_queue.put(
                        ProgressEvent(
                            event_type="status",
                            message=f"[완료] {document_info.pdf_path.name} 변환 완료",
                            current_file=file_index,
                            total_files=total_files,
                        )
                    )
                except Exception as error:
                    fail_count += 1
                    self.event_queue.put(
                        ProgressEvent(
                            event_type="status",
                            message=f"[오류] {document_info.pdf_path.name} 변환 실패: {error}",
                            current_file=file_index,
                            total_files=total_files,
                        )
                    )

                self.event_queue.put(
                    ProgressEvent(
                        event_type="summary",
                        total_files=total_files,
                        success_count=success_count,
                        fail_count=fail_count,
                        current_file=file_index,
                    )
                )

            self.event_queue.put(
                ProgressEvent(
                    event_type="done",
                    message=(
                        f"JPG 변환 완료\n"
                        f"성공: {success_count}개\n"
                        f"실패: {fail_count}개\n"
                        f"전체: {total_files}개"
                    ),
                    total_files=total_files,
                    success_count=success_count,
                    fail_count=fail_count,
                )
            )
        except Exception as error:
            self.event_queue.put(ProgressEvent(event_type="done", message=f"변환 중 오류가 발생했습니다: {error}"))

    def process_event_queue(self) -> None:
        while not self.event_queue.empty():
            event = self.event_queue.get()
            self.handle_progress_event(event)
        self.after(100, self.process_event_queue)

    def handle_progress_event(self, event: ProgressEvent) -> None:
        if event.event_type == "status":
            self.status_label.configure(text=event.message)
            self.append_log(event.message)
            if event.total_files:
                self.progress_detail_label.configure(text=f"파일 진행: {event.current_file}/{event.total_files}")
                if event.current_file:
                    self.progress_bar.set(event.current_file / max(event.total_files, 1))

        elif event.event_type == "page":
            self.status_label.configure(text=event.message)
            if event.total_files and event.total_pages:
                completed_files = event.current_file - 1
                completed_pages_ratio = event.current_page / event.total_pages
                overall_progress = (completed_files + completed_pages_ratio) / event.total_files
                self.progress_bar.set(overall_progress)
            self.progress_detail_label.configure(
                text=f"파일 {event.current_file}/{event.total_files} | 페이지 {event.current_page}/{event.total_pages}"
            )

        elif event.event_type == "summary":
            self.summary_label.configure(
                text=(
                    f"총 {event.total_files}개 문서\n"
                    f"성공 {event.success_count}개\n"
                    f"실패 {event.fail_count}개"
                )
            )

        elif event.event_type == "analysis_complete":
            self.documents = event.documents
            self.refresh_filter_values()
            self.status_label.configure(text=event.message)
            self.append_log(event.message)
            self.progress_bar.set(1 if event.total_files else 0)
            self.progress_detail_label.configure(text="분석 종료")
            self.refresh_selection_panel()
            self.update_idletasks()
            self.set_running_state(False)
            self.append_log("[안내] 오른쪽 '기안 제목 만들기' 영역에서 회사별 발주번호를 선택할 수 있습니다.")

        elif event.event_type == "done":
            self.progress_bar.set(1 if event.total_files else 0)
            self.status_label.configure(text=event.message)
            self.progress_detail_label.configure(text="작업 종료")
            self.append_log(event.message.replace("\n", " | "))
            self.set_running_state(False)
            messagebox.showinfo("작업 결과", event.message)

    def refresh_selection_panel(self) -> None:
        for child in self.selection_frame.winfo_children():
            child.destroy()

        self.selection_vars.clear()
        self.selection_checkboxes.clear()
        self.selection_order.clear()
        self.selection_meta.clear()
        self.selected_company = None
        self.set_preview_text("선택한 발주번호가 여기에 표시됩니다.")

        grouped = self.group_documents_by_company()
        if not grouped:
            self._populate_empty_selection_state()
            return

        total_companies = len(grouped)
        total_order_candidates = sum(len(docs) for docs in grouped.values())

        row_index = 0
        for company_name, docs in grouped.items():
            header_text = f"{company_name}  |  문서 {len(docs)}개"
            ctk.CTkLabel(
                self.selection_frame,
                text=header_text,
                font=ctk.CTkFont(family="Malgun Gothic", size=15, weight="bold"),
                text_color="#4b3d37",
            ).grid(row=row_index, column=0, padx=12, pady=(8, 2), sticky="w")
            row_index += 1

            doc_dates = sorted({doc.document_date for doc in docs if doc.document_date != MISSING_VALUE})
            header_subtext = (
                f"날짜: {', '.join(doc_dates[:2])}{' 외' if len(doc_dates) > 2 else ''}"
                if doc_dates else "날짜: 확인필요"
            )
            ctk.CTkLabel(
                self.selection_frame,
                text=header_subtext,
                font=ctk.CTkFont(family="Malgun Gothic", size=11),
                text_color="#8d7f73",
            ).grid(row=row_index, column=0, padx=20, pady=(0, 2), sticky="w")
            row_index += 1

            for doc in docs:
                row_frame = ctk.CTkFrame(self.selection_frame, fg_color="transparent")
                row_frame.grid(row=row_index, column=0, padx=16, pady=2, sticky="ew")
                row_frame.grid_columnconfigure(0, weight=1)
                row_frame.grid_columnconfigure(1, weight=0)

                row_key = str(doc.pdf_path.resolve())
                variable = tk.BooleanVar(value=False)
                display_number = doc.representative_order_number or MISSING_VALUE
                checkbox = ctk.CTkCheckBox(
                    row_frame,
                    text=f"{display_number}   ({doc.pdf_path.name})",
                    variable=variable,
                    command=lambda key=row_key: self.on_selection_changed(key),
                    font=ctk.CTkFont(family="Malgun Gothic", size=14, weight="bold"),
                    text_color="#5a4b44",
                    fg_color="#81b29a",
                    hover_color="#6b9b84",
                    border_color="#b9a899",
                )
                checkbox.grid(row=0, column=0, sticky="w")
                edit_button = ctk.CTkButton(
                    row_frame,
                    text="편집",
                    width=70,
                    height=30,
                    corner_radius=12,
                    fg_color="#d9b08c",
                    hover_color="#c69b76",
                    text_color="#fffaf6",
                    font=ctk.CTkFont(family="Malgun Gothic", size=12, weight="bold"),
                    command=lambda document=doc: self.open_edit_dialog(document),
                )
                edit_button.grid(row=0, column=1, padx=(8, 0), sticky="e")

                self.selection_vars[row_key] = variable
                self.selection_checkboxes[row_key] = checkbox
                self.selection_order.append(row_key)
                self.selection_meta[row_key] = (company_name, display_number)
                row_index += 1

        self.summary_label.configure(
            text=(
                f"회사 {total_companies}개\n"
                f"번호 후보 {total_order_candidates}개\n"
                f"위 목록에서 바로 체크"
            )
        )

    def group_documents_by_company(self) -> Dict[str, List[DocumentInfo]]:
        grouped: Dict[str, List[DocumentInfo]] = {}
        for document in self.get_filtered_documents():
            grouped.setdefault(document.company_name, []).append(document)
        return dict(sorted(grouped.items(), key=lambda item: item[0]))

    def on_selection_changed(self, current_key: str) -> None:
        company_name, _order_number = self.selection_meta[current_key]
        is_checked = self.selection_vars[current_key].get()

        if is_checked and self.selected_company is None:
            self.selected_company = company_name
        elif not is_checked:
            selected_keys = [key for key, var in self.selection_vars.items() if var.get()]
            self.selected_company = self.selection_meta[selected_keys[0]][0] if selected_keys else None

        self.update_checkbox_states()
        self.update_title_preview()

    def update_checkbox_states(self) -> None:
        for key, checkbox in self.selection_checkboxes.items():
            company_name, _order_number = self.selection_meta[key]
            checkbox.configure(state="disabled" if self.selected_company and company_name != self.selected_company else "normal")

    def update_title_preview(self) -> None:
        selected_keys = [key for key in self.selection_order if self.selection_vars.get(key) and self.selection_vars[key].get()]
        if not selected_keys:
            self.set_preview_text("선택한 발주번호가 여기에 표시됩니다.")
            return

        company_name, _first_order = self.selection_meta[selected_keys[0]]
        if company_name == MISSING_VALUE:
            self.set_preview_text("회사명을 자동으로 찾지 못했습니다. companies.txt 또는 PDF 원문을 확인해주세요.")
            return

        seen_numbers = set()
        order_numbers: List[str] = []
        for key in selected_keys:
            _company, number = self.selection_meta[key]
            if number == MISSING_VALUE:
                continue
            normalized = number.upper()
            if normalized in seen_numbers:
                continue
            seen_numbers.add(normalized)
            order_numbers.append(number)

        if not order_numbers:
            self.set_preview_text("선택한 항목에서 사용할 발주번호를 찾지 못했습니다.")
            return

        title = f"{TITLE_PREFIX} {company_name} {', '.join(order_numbers)}"
        self.set_preview_text(title)

    def clear_selection(self, reset_documents: bool = False) -> None:
        self.selected_company = None
        for variable in self.selection_vars.values():
            variable.set(False)
        self.update_checkbox_states()
        self.set_preview_text("선택한 발주번호가 여기에 표시됩니다.")
        if reset_documents:
            self.documents = []
            self.refresh_filter_values()


    def export_summary(self) -> None:
        if not self.documents:
            messagebox.showwarning("분석 필요", "먼저 문서 분석을 완료해주세요.")
            return

        base_dir = self.selected_folder if self.selected_folder and self.selected_folder.exists() else self.script_dir
        output_path = base_dir / f"분석요약_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        fieldnames = [
            "pdf파일명",
            "회사명",
            "매칭된문구",
            "규칙출처",
            "문서날짜",
            "대표발주번호",
            "전체발주번호",
            "상태",
            "페이지수",
            "OCR사용",
            "텍스트미리보기",
        ]

        with output_path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            for document in self.documents:
                writer.writerow(
                    {
                        "pdf파일명": document.pdf_path.name,
                        "회사명": document.company_name,
                        "매칭된문구": document.matched_alias,
                        "규칙출처": document.company_rule_source,
                        "문서날짜": document.document_date,
                        "대표발주번호": document.representative_order_number,
                        "전체발주번호": ", ".join(document.order_numbers),
                        "상태": document.status,
                        "페이지수": document.page_count,
                        "OCR사용": "Y" if document.used_ocr else "N",
                        "텍스트미리보기": document.text_excerpt,
                    }
                )

        self.append_log(f"[내보내기] 분석 요약 CSV 저장: {output_path}")
        messagebox.showinfo("내보내기 완료", f"분석 요약을 저장했습니다.\n{output_path}")

    def copy_preview(self) -> None:
        text = self.preview_text.get().strip()
        invalid_messages = {
            "",
            "선택한 발주번호가 여기에 표시됩니다.",
            "회사명을 자동으로 찾지 못했습니다. companies.txt 또는 PDF 원문을 확인해주세요.",
            "선택한 항목에서 사용할 발주번호를 찾지 못했습니다.",
        }
        if text in invalid_messages:
            messagebox.showwarning("복사할 제목 없음", "먼저 발주번호를 선택해서 제목을 만들어주세요.")
            return

        self.clipboard_clear()
        self.clipboard_append(text)
        self.append_log(f"[복사] {text}")
        messagebox.showinfo("복사 완료", "기안 제목을 클립보드에 복사했습니다.")


def main() -> None:
    app = PdfToJpgApp()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
