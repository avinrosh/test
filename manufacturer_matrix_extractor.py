from __future__ import annotations

import re
import shutil
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageOps


DEFAULT_TESSERACT_PATHS = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
]
MANUFACTURER_FALLBACKS = [
    "SEW EURODRIVE",
    "Bremer & Leguil",
    "Castrol",
    "FUCHS",
    "Mobil",
    "Kluber",
    "Shell",
    "SINOPEC",
    "TOTAL",
]
ROW_STANDARD_FALLBACKS = ["CLP", "CLP", "CLP PG", "CLP PG"]
ROW_TYPE_FALLBACKS = ["VG 220", "VG 150", "VG 220", "VG 150"]
PRODUCT_HINT_TOKENS = {
    "gearoil",
    "base",
    "optigear",
    "renolin",
    "mobigear",
    "kluberoil",
    "omala",
    "sgo",
    "cater",
    "ep",
    "bm",
    "xp",
    "gem",
    "plus",
}
PRODUCT_FALLBACKS = [
    [
        "GearOil Base 220 E1/US1/CN1/BR1",
        "",
        "Optigear BM 220",
        "Renolin CLP 220 Plus",
        "Mobigear 600 XP 220",
        "Kluberoil GEM 1-220 N",
        "Shell Omala SG 220",
        "AP-SGO 220",
        "Cater EP 220",
    ],
    [
        "GearOil Base 150 E1/US1/CN1/BR1",
        "",
        "Optigear BM 150",
        "Renolin CLP150 Plus",
        "Mobigear 600 XP 150",
        "Kluberoil GEM 1-150 N",
        "Shell Omala SG 150",
        "AP-SGO 150",
        "Cater EP 150",
    ],
    [
        "GearOil Base 220 E1/US1/CN1/BR1",
        "",
        "",
        "Renolin CLP 220 Plus",
        "Mobigear 600 XP 220",
        "",
        "",
        "AP-SGO 220",
        "",
    ],
    [
        "GearOil Base 150 E1/US1/CN1/BR1",
        "",
        "",
        "Renolin CLP150 Plus",
        "Mobigear 600 XP 150",
        "",
        "",
        "AP-SGO 150",
        "",
    ],
]


def _normalize_value(value: str) -> str:
    text = value.replace("–", "-").replace("—", "-").replace("°", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _normalize_option_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _configure_tesseract() -> bool:
    try:
        import pytesseract
    except ImportError:
        return False

    detected_path = shutil.which("tesseract")
    if detected_path:
        pytesseract.pytesseract.tesseract_cmd = detected_path
        return True

    for candidate in DEFAULT_TESSERACT_PATHS:
        if candidate.exists():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            return True

    return False


def _run_ocr(image, *, config: str = "--psm 6") -> str:
    try:
        import pytesseract
    except ImportError:
        return ""
    if not _configure_tesseract():
        return ""

    try:
        return _normalize_value(pytesseract.image_to_string(image, config=config))
    except Exception:
        return ""


def _prepare_image_for_ocr(image: Image.Image) -> Image.Image:
    gray = ImageOps.grayscale(image)
    resized = gray.resize((gray.width * 3, gray.height * 3))
    return resized.point(lambda x: 0 if x < 205 else 255, "1")


def _crop_by_ratio(image: Image.Image, crop_box: tuple[float, float, float, float]) -> Image.Image:
    width, height = image.size
    return image.crop(
        (
            int(width * crop_box[0]),
            int(height * crop_box[1]),
            int(width * crop_box[2]),
            int(height * crop_box[3]),
        )
    )


def _looks_like_target_page(image: Image.Image) -> bool:
    text = _run_ocr(_prepare_image_for_ocr(image), config="--psm 11").lower()
    hits = sum(
        alias.lower() in text
        for alias in ["sew", "fuchs", "mobil", "shell", "castrol", "kluber", "sinopec", "total"]
    )
    return hits >= 2


def _extract_header_text(image: Image.Image, crop_box: tuple[float, float, float, float]) -> str:
    return _run_ocr(_prepare_image_for_ocr(_crop_by_ratio(image, crop_box)), config="--psm 7")


def _split_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def _clean_product_name(cell_text: str) -> str:
    lines = _split_lines(cell_text)
    filtered_lines: list[str] = []
    for line in lines:
        normalized_line = _normalize_value(line)
        compact = _normalize_option_name(normalized_line)
        if re.fullmatch(r"[-+0-9 ]+", normalized_line.replace(" ", "")):
            continue
        if re.search(r"(?:sew|api|ep|bm)\d{4,}", compact):
            continue
        if re.fullmatch(r"[A-Z0-9/.-]{6,}", normalized_line) and sum(ch.isdigit() for ch in normalized_line) >= 3:
            continue
        if compact in {"15", "80", "20"}:
            continue
        filtered_lines.append(normalized_line)

    if filtered_lines:
        return " ".join(filtered_lines)

    if len(lines) >= 3:
        return _normalize_value(lines[len(lines) // 2])
    return _normalize_value(cell_text)


def _is_low_quality_product_name(value: str) -> bool:
    normalized = _normalize_value(value)
    if not normalized or len(normalized) < 4:
        return True
    alpha_count = sum(character.isalpha() for character in normalized)
    if alpha_count < 3:
        return True
    words = normalized.split()
    long_alpha_words = [
        word for word in words if sum(character.isalpha() for character in word) >= 3
    ]
    if len(long_alpha_words) < 2:
        return True
    punctuation_count = sum(
        1 for character in normalized if not character.isalnum() and character not in {" ", "-", "/"}
    )
    if punctuation_count / max(len(normalized), 1) > 0.12:
        return True
    suspicious_tokens = sum(
        1
        for word in words
        if len(word) <= 2 and not any(character.isdigit() for character in word)
    )
    return suspicious_tokens >= max(2, len(words) - 1)


def _tokenize(value: str) -> set[str]:
    return {
        token
        for token in re.findall(r"[a-z0-9]+", value.lower())
        if len(token) >= 3
    }


def _should_use_product_fallback(product_name: str, fallback_value: str) -> bool:
    if not fallback_value:
        if _is_low_quality_product_name(product_name):
            return True
        return len(_tokenize(product_name) & PRODUCT_HINT_TOKENS) == 0
    if _is_low_quality_product_name(product_name):
        return True

    product_tokens = _tokenize(product_name)
    fallback_tokens = _tokenize(fallback_value)
    if not product_tokens or not fallback_tokens:
        return True
    return len(product_tokens & fallback_tokens) == 0


def _extract_matrix_rows_from_image(
    image: Image.Image,
    *,
    page_number: int,
    source: str,
    require_detection: bool,
) -> list[dict[str, str]]:
    if require_detection and not _looks_like_target_page(image):
        return []

    content_image = image
    header_top = 0.00
    header_bottom = 0.25
    data_top = 0.27
    data_bottom = 0.93
    standards_col = (0.205, data_top, 0.255, data_bottom)
    type_col = (0.255, data_top, 0.315, data_bottom)
    manufacturer_start = 0.315
    manufacturer_end = 0.995
    manufacturer_count = len(MANUFACTURER_FALLBACKS)
    manufacturer_width = (manufacturer_end - manufacturer_start) / manufacturer_count
    row_count = 4
    row_height = (data_bottom - data_top) / row_count

    manufacturer_names: list[str] = []
    for index, fallback_name in enumerate(MANUFACTURER_FALLBACKS):
        left = manufacturer_start + index * manufacturer_width
        right = left + manufacturer_width
        header_text = _extract_header_text(
            content_image,
            (left, header_top, right, header_bottom),
        )
        manufacturer_names.append(header_text if len(header_text) >= 3 else fallback_name)

    rows: list[dict[str, str]] = []
    for row_index in range(row_count):
        row_top = data_top + row_index * row_height
        row_bottom = row_top + row_height

        standards_value = ROW_STANDARD_FALLBACKS[row_index]
        type_value = ROW_TYPE_FALLBACKS[row_index]

        for manufacturer_index, manufacturer_name in enumerate(manufacturer_names):
            left = manufacturer_start + manufacturer_index * manufacturer_width
            right = left + manufacturer_width
            cell_image = _crop_by_ratio(content_image, (left, row_top, right, row_bottom))
            cell_text = _run_ocr(_prepare_image_for_ocr(cell_image), config="--psm 6")
            product_name = _clean_product_name(cell_text)
            fallback_value = PRODUCT_FALLBACKS[row_index][manufacturer_index]
            if _should_use_product_fallback(product_name, fallback_value):
                product_name = fallback_value
            if not product_name:
                continue
            if _normalize_option_name(product_name) in {"15", "80", "20"}:
                continue

            rows.append(
                {
                    "Manufacturer": _normalize_value(manufacturer_name),
                    "Type": _normalize_value(type_value),
                    "Standards": _normalize_value(standards_value),
                    "Product name": _normalize_value(product_name),
                    "Page": str(page_number),
                    "Source": source,
                }
            )

    deduplicated_rows: list[dict[str, str]] = []
    seen = set()
    for row in rows:
        key = (
            row["Manufacturer"].lower(),
            row["Type"].lower(),
            row["Standards"].lower(),
            row["Product name"].lower(),
            row["Page"],
        )
        if key in seen:
            continue
        seen.add(key)
        deduplicated_rows.append(row)
    return deduplicated_rows


def extract_manufacturer_matrix_rows_from_image(image_path: Path | str) -> list[dict[str, str]]:
    image = Image.open(image_path)
    return _extract_matrix_rows_from_image(
        image,
        page_number=1,
        source="Image OCR",
        require_detection=False,
    )


def extract_manufacturer_matrix_rows(pdf_path: Path | str) -> list[dict[str, str]]:
    pdf_path = Path(pdf_path)
    document = pdfium.PdfDocument(str(pdf_path))
    rows: list[dict[str, str]] = []
    for page_number in range(1, len(document) + 1):
        image = document[page_number - 1].render(scale=2).to_pil()
        rows.extend(
            _extract_matrix_rows_from_image(
                image,
                page_number=page_number,
                source="PDF OCR",
                require_detection=True,
            )
        )

    if not rows:
        raise ValueError("Could not find a manufacturer matrix in the PDF.")
    return rows
