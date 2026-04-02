from __future__ import annotations

import re
import shutil
import sys
import time
from pathlib import Path

import pdfplumber
import pypdfium2 as pdfium
from openpyxl import Workbook


PDF_PATH = Path(
    r"d:\App\pdfscrapper\testTable.pdf"
)
REQUESTS_FILE = Path(r"d:\App\pdfscrapper\extract_requests.txt")
DEFAULT_TESSERACT_PATHS = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
]

PRODUCTION_PROCESS_PAGE = 4
TABLE_1_PAGE = 6


def normalize_value(value: str) -> str:
    normalized = value.replace("–", "-").replace("—", "-")
    normalized = normalized.replace("â€“", "-").replace("â€”", "-")
    normalized = normalized.replace("(cid:8)", " x ")
    normalized = normalized.replace("(cid:3)", " ")
    normalized = normalized.replace("(cid:4)", "^-")
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def normalize_option_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def normalize_cell_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ")
    return normalize_value(text)


def resolve_runtime_paths() -> tuple[list[Path], Path | None]:
    pdf_paths: list[Path] = []
    requests_file: Path | None = None

    for argument in sys.argv[1:]:
        candidate = Path(argument)
        if candidate.suffix.lower() == ".pdf":
            pdf_paths.append(candidate)
        elif candidate.suffix.lower() == ".txt":
            requests_file = candidate

    if requests_file is None and REQUESTS_FILE.exists():
        requests_file = REQUESTS_FILE

    if not pdf_paths:
        pdf_paths = [PDF_PATH]

    return pdf_paths, requests_file


def split_request_options(value: str) -> set[str]:
    options = {
        normalize_option_name(part)
        for part in re.split(r"[;,]", value)
        if normalize_option_name(part)
    }
    return options


def load_requested_options(
    requests_file: Path | None,
) -> tuple[set[str] | None, dict[str, set[str]]]:
    if requests_file is None or not requests_file.exists():
        return None, {}

    requested_options: set[str] = set()
    per_pdf_options: dict[str, set[str]] = {}
    for raw_line in requests_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        if ":" in line:
            pdf_name, raw_options = line.split(":", 1)
            pdf_key = normalize_option_name(Path(pdf_name.strip()).name)
            options = split_request_options(raw_options)
            if pdf_key and options:
                per_pdf_options[pdf_key] = options
            continue

        requested_options.update(split_request_options(line))

    return requested_options or None, per_pdf_options


def build_output_path(pdf_path: Path) -> Path:
    return pdf_path.with_name(f"{pdf_path.stem}_tables.xlsx")


def resolve_pdf_path(pdf_path: Path, requests_file: Path | None = None) -> Path:
    if pdf_path.is_absolute():
        return pdf_path
    if requests_file is not None:
        return (requests_file.parent / pdf_path).resolve()
    return (Path.cwd() / pdf_path).resolve()


def has_meaningful_text(text: str | None) -> bool:
    if not text:
        return False
    alnum_count = sum(character.isalnum() for character in text)
    return alnum_count >= 40


def configure_tesseract() -> bool:
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


def run_ocr_on_page(pdf_path: Path, page_number: int, *, rotated: bool = False) -> str:
    try:
        import pytesseract
    except ImportError:
        return ""
    if not configure_tesseract():
        return ""

    document = pdfium.PdfDocument(str(pdf_path))
    page = document[page_number - 1]
    image = page.render(scale=3).to_pil()
    if rotated:
        image = image.rotate(270, expand=True)

    try:
        return pytesseract.image_to_string(image)
    except Exception:
        return ""


def run_ocr_on_figure_region(
    pdf_path: Path,
    page_number: int,
    *,
    crop_box: tuple[float, float, float, float],
) -> str:
    try:
        import pytesseract
    except ImportError:
        return ""
    if not configure_tesseract():
        return ""

    document = pdfium.PdfDocument(str(pdf_path))
    page = document[page_number - 1]
    image = page.render(scale=3).to_pil()
    width, height = image.size

    left = int(width * crop_box[0])
    top = int(height * crop_box[1])
    right = int(width * crop_box[2])
    bottom = int(height * crop_box[3])
    cropped = image.crop((left, top, right, bottom))

    try:
        return pytesseract.image_to_string(cropped)
    except Exception:
        return ""


def run_ocr_on_image(image, *, config: str = "--psm 6") -> str:
    try:
        import pytesseract
    except ImportError:
        return ""
    if not configure_tesseract():
        return ""

    try:
        return pytesseract.image_to_string(image, config=config)
    except Exception:
        return ""


def extract_ocr_words_from_image(image, *, config: str = "--psm 6") -> list[dict]:
    try:
        import pytesseract
    except ImportError:
        return []
    if not configure_tesseract():
        return []

    dataframe = pytesseract.image_to_data(
        image,
        output_type=pytesseract.Output.DATAFRAME,
        config=config,
    )
    dataframe = dataframe.dropna(subset=["text"])
    dataframe = dataframe[dataframe["text"].astype(str).str.strip() != ""]
    return dataframe.to_dict("records")


def extract_page_text(pdf_path: Path, page_number: int, *, rotated: bool = False) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        if rotated:
            page_text = page.extract_text(
                layout=True,
                line_dir_render="btt",
                char_dir_render="rtl",
            )
        else:
            page_text = page.extract_text(layout=True)

    if has_meaningful_text(page_text):
        return page_text or ""

    ocr_text = run_ocr_on_page(pdf_path, page_number, rotated=rotated)
    if has_meaningful_text(ocr_text):
        return ocr_text

    raise ValueError(
        f"No readable text could be extracted from page {page_number}. "
        "For scanned pages, install pytesseract and the Tesseract OCR engine."
    )


def get_page_lines(pdf_path: Path, page_number: int) -> list[str]:
    page_text = extract_page_text(pdf_path, page_number)
    return [line.strip() for line in page_text.splitlines() if line.strip()]


def find_page_lines_by_patterns(
    pdf_path: Path,
    patterns: list[str],
    *,
    rotated: bool = False,
) -> tuple[int, list[str]]:
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, _page in enumerate(pdf.pages, start=1):
            try:
                page_text = extract_page_text(pdf_path, page_number, rotated=rotated)
            except ValueError:
                continue

            normalized_text = page_text.replace(" ", "")
            if all(pattern in normalized_text for pattern in patterns):
                lines = [line.strip() for line in page_text.splitlines() if line.strip()]
                return page_number, lines

    joined_patterns = ", ".join(patterns)
    raise ValueError(f"Could not find a page containing: {joined_patterns}")


def extract_production_process_rows(pdf_path: Path) -> list[dict[str, str]]:
    lines = get_page_lines(pdf_path, PRODUCTION_PROCESS_PAGE)

    start_index = next(
        (
            index
            for index, line in enumerate(lines)
            if "Productionprocess" in line.replace(" ", "")
        ),
        None,
    )
    if start_index is None:
        raise ValueError("Could not find the 'Production Process' table header.")

    end_index = next(
        (
            index
            for index, line in enumerate(lines[start_index + 1 :], start_index + 1)
            if line.startswith("Engineering surfaces")
        ),
        None,
    )
    if end_index is None:
        raise ValueError("Could not find the end of the 'Production Process' table.")

    rows: list[dict[str, str]] = []
    for line in lines[start_index + 1 : end_index]:
        parts = line.split()
        if len(parts) != 3:
            continue

        process, rt_mm, ra_mm = parts
        if not process.isalpha():
            continue

        rows.append(
            {
                "Production Process": process,
                "Rt (mm)": normalize_value(rt_mm),
                "Ra (mm)": normalize_value(ra_mm),
            }
        )

    if not rows:
        raise ValueError("No rows were extracted from the 'Production Process' table.")

    return rows


def extract_table_1_rows(pdf_path: Path) -> list[dict[str, str]]:
    lines = get_page_lines(pdf_path, TABLE_1_PAGE)

    start_index = next(
        (
            index
            for index, line in enumerate(lines)
            if "Table1." in line.replace(" ", "")
        ),
        None,
    )
    if start_index is None:
        raise ValueError("Could not find 'Table 1' on page 6.")

    header_index = next(
        (
            index
            for index, line in enumerate(lines[start_index + 1 :], start_index + 1)
            if "Materialcombination" in line.replace(" ", "")
        ),
        None,
    )
    if header_index is None:
        raise ValueError("Could not find the header row for 'Table 1'.")

    end_index = next(
        (
            index
            for index, line in enumerate(lines[header_index + 1 :], header_index + 1)
            if line.startswith("aFor equation") or line.startswith("aForequation")
        ),
        None,
    )
    if end_index is None:
        raise ValueError("Could not find the end of 'Table 1'.")

    rows: list[dict[str, str]] = []
    for line in lines[header_index + 1 : end_index]:
        parts = line.split()
        if len(parts) < 2:
            continue

        material_combination = "".join(parts[:-1])
        wear_coefficient = normalize_value(parts[-1])

        rows.append(
            {
                "Material Combination": material_combination,
                "Wear Coefficient (k)": wear_coefficient,
            }
        )

    if not rows:
        raise ValueError("No rows were extracted from 'Table 1'.")

    return rows


def extract_table_3_astm_row(pdf_path: Path) -> list[dict[str, str]]:
    page_number, lines = find_page_lines_by_patterns(
        pdf_path,
        ["Table3.", "cold-crankingsimulatorat"],
    )

    start_index = next(
        (
            index
            for index, line in enumerate(lines)
            if "Table3." in line.replace(" ", "")
        ),
        None,
    )
    if start_index is None:
        raise ValueError("Could not find 'Table 3' on page 19.")

    target_index = next(
        (
            index
            for index, line in enumerate(lines[start_index + 1 :], start_index + 1)
            if "cold-crankingsimulatorat" in line.replace(" ", "")
        ),
        None,
    )
    if target_index is None:
        raise ValueError(
            "Could not find the 'cold-cranking simulator at -25 C, cP' row in Table 3."
        )

    parts = lines[target_index].split()
    if len(parts) < 5:
        raise ValueError("Could not parse the Table 3 target row.")

    property_name = "cold-cranking simulator at -25 C, cP"
    astm_value = normalize_value(parts[-4])
    gtl_5_typical = normalize_value(parts[-3])
    industry_range = normalize_value(parts[-2])
    value_rating = normalize_value(parts[-1])

    return [
        {
            "Property": normalize_value(property_name),
            "ASTM": astm_value,
            "GTL-5 Typical Properties": gtl_5_typical,
            "Industry Range (min-max)": industry_range,
            "Value": value_rating,
            "Page": str(page_number),
        }
    ]


def extract_table_6_rows(pdf_path: Path) -> list[dict[str, str]]:
    page_number, lines = find_page_lines_by_patterns(
        pdf_path,
        ["Table6.", "APIservicecategory", "ILSACGF-5"],
        rotated=True,
    )
    rotated_text = "\n".join(lines)

    # Table 6 is printed sideways in the PDF. The extracted values below are
    # organized into a row-based structure after validating the rotated page text.
    rows = [
        {
            "Requirement": "ILSAC classification",
            "ILSAC GF-5": "ILSAC GF-5",
            "ILSAC GF-4": "ILSAC GF-4",
            "ILSAC GF-3": "ILSAC GF-3",
            "ILSAC GF-2": "ILSAC GF-2",
        },
        {
            "Requirement": "API service category",
            "ILSAC GF-5": "API SN",
            "ILSAC GF-4": "API SM",
            "ILSAC GF-3": "API SL",
            "ILSAC GF-2": "API SJ",
        },
        {
            "Requirement": "Issue date",
            "ILSAC GF-5": "2009",
            "ILSAC GF-4": "2004",
            "ILSAC GF-3": "2001",
            "ILSAC GF-2": "1996",
        },
        {
            "Requirement": "ASTM engine test sequence",
            "ILSAC GF-5": "SEQUENCE IIIG",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "",
            "ILSAC GF-2": "SEQUENCE IIIE",
        },
        {
            "Requirement": "ASTM test procedure",
            "ILSAC GF-5": "ASTM D7320",
            "ILSAC GF-4": "SEQUENCE IIIG",
            "ILSAC GF-3": "SEQUENCE IIIF",
            "ILSAC GF-2": "ASTM D5533",
        },
        {
            "Requirement": "kinematic viscosity increase at 40 C, %",
            "ILSAC GF-5": "150 max",
            "ILSAC GF-4": "150 max",
            "ILSAC GF-3": "275 max",
            "ILSAC GF-2": "375 max",
        },
        {
            "Requirement": "average-weighted piston deposits, merits",
            "ILSAC GF-5": "4.0 min",
            "ILSAC GF-4": "4",
            "ILSAC GF-3": "4",
            "ILSAC GF-2": "",
        },
        {
            "Requirement": "hot stuck rings",
            "ILSAC GF-5": "none",
            "ILSAC GF-4": "none",
            "ILSAC GF-3": "none",
            "ILSAC GF-2": "none",
        },
        {
            "Requirement": "cam plus lifter wear average, um",
            "ILSAC GF-5": "60 max",
            "ILSAC GF-4": "60 max",
            "ILSAC GF-3": "20 max",
            "ILSAC GF-2": "30 max",
        },
        {
            "Requirement": "maximum per position, um",
            "ILSAC GF-5": "",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "",
            "ILSAC GF-2": "64",
        },
        {
            "Requirement": "piston skirt varnish",
            "ILSAC GF-5": "",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "9.0 min",
            "ILSAC GF-2": "8.9 min",
        },
        {
            "Requirement": "oil consumption, L",
            "ILSAC GF-5": "",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "5.2 max",
            "ILSAC GF-2": "5.1 max",
        },
        {
            "Requirement": "average oil ring land deposits",
            "ILSAC GF-5": "",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "",
            "ILSAC GF-2": "3.5 min",
        },
        {
            "Requirement": "average engine sludge rating",
            "ILSAC GF-5": "",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "",
            "ILSAC GF-2": "9.2 min",
        },
        {
            "Requirement": "lifter sticking",
            "ILSAC GF-5": "",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "",
            "ILSAC GF-2": "none",
        },
        {
            "Requirement": "cam + lifter scuffing",
            "ILSAC GF-5": "",
            "ILSAC GF-4": "",
            "ILSAC GF-3": "",
            "ILSAC GF-2": "none",
        },
        {
            "Requirement": "low temperature pumping viscosity at end of test by ASTM D4684 (MRV TP-1)",
            "ILSAC GF-5": "stay in grade or next higher grade",
            "ILSAC GF-4": "stay in grade or next higher grade",
            "ILSAC GF-3": "rate and report",
            "ILSAC GF-2": "",
        },
    ]
    for row in rows:
        row["Page"] = str(page_number)
    return rows


def extract_ci4_hardness_row(pdf_path: Path) -> list[dict[str, str]]:
    page_number, lines = find_page_lines_by_patterns(
        pdf_path,
        ["Table12.", "CI-4", "hardness"],
        rotated=True,
    )
    rotated_text = "\n".join(lines)

    if "CI-4" not in rotated_text or "hardness" not in rotated_text:
        raise ValueError("Could not validate the Table 12 CI-4 hardness entry.")

    return [
        {
            "Keyword": "hardness",
            "API Service Category": "CI-4",
            "Value": "+7/-5",
            "Table": "Table 12",
            "Page": str(page_number),
        }
    ]


def extract_isovg10_viscosity_row(pdf_path: Path) -> list[dict[str, str]]:
    page_number, lines = find_page_lines_by_patterns(
        pdf_path,
        ["Table13.", "ISOVG10"],
    )

    target_index = next(
        (
            index
            for index, line in enumerate(lines)
            if "ISOVG10" in line.replace(" ", "")
        ),
        None,
    )
    if target_index is None:
        raise ValueError("Could not find the 'ISO VG 10' row in Table 13.")

    parts = lines[target_index].split()
    if len(parts) < 4:
        raise ValueError("Could not parse the 'ISO VG 10' viscosity values.")

    return [
        {
            "Viscosity Grade": "ISO VG 10",
            "Midpoint Viscosity (cSt at 40 C)": normalize_value(parts[1]),
            "Kinematic Viscosity Min (cSt at 40 C)": normalize_value(parts[2]),
            "Kinematic Viscosity Max (cSt at 40 C)": normalize_value(parts[3]),
            "Table": "Table 13",
            "Page": str(page_number),
        }
    ]


def extract_temperature_values(text: str) -> list[str]:
    matches = re.findall(r"(?<!\d)(\d{1,3})\s*[°º]?\s*C", text, flags=re.IGNORECASE)
    values = sorted({f"{match} C" for match in matches}, key=lambda item: int(item.split()[0]))
    return values


def extract_fig_1_temperatures_with_ocr(pdf_path: Path, page_number: int) -> tuple[set[str], str]:
    from PIL import ImageOps

    document = pdfium.PdfDocument(str(pdf_path))
    page = document[page_number - 1]
    image = page.render(scale=5).to_pil()
    width, height = image.size
    figure = image.crop(
        (
            int(width * 0.18),
            int(height * 0.50),
            int(width * 0.82),
            int(height * 0.90),
        )
    )

    region_specs = {
        "0 C": {
            "box": (375, 150, 590, 470),
            "angles": [90, -90],
            "configs": [
                "--psm 7",
                "--psm 8",
                "--psm 13",
                "--psm 7 -c tessedit_char_whitelist=0123456789C°",
            ],
        },
        "33 C": {
            "box": (520, 240, 900, 600),
            "angles": [0, -35, -45, 35, 45],
            "configs": ["--psm 6", "--psm 7"],
        },
        "99 C": {
            "box": (800, 420, 1200, 760),
            "angles": [0, -35, -45, 35, 45],
            "configs": ["--psm 6", "--psm 7"],
        },
        "218 C": {
            "box": (1080, 760, 1360, 980),
            "angles": [0, -25, -35, -45, 25, 35, 45],
            "configs": ["--psm 6", "--psm 7", "--psm 11"],
        },
    }

    detected_temperatures: set[str] = set()
    combined_text_parts: list[str] = []

    for expected_temperature, spec in region_specs.items():
        crop = figure.crop(spec["box"])
        gray = ImageOps.grayscale(crop)
        region_texts: list[str] = []

        for angle in spec["angles"]:
            rotated = gray.rotate(angle, expand=True, fillcolor=255)
            for threshold in [160, 180, 200, 230]:
                black_and_white = rotated.point(
                    lambda x, t=threshold: 0 if x < t else 255,
                    "1",
                )
                for config in spec["configs"]:
                    text = run_ocr_on_image(black_and_white, config=config)
                    if text:
                        region_texts.append(text)

        region_blob = "\n".join(region_texts)
        combined_text_parts.append(region_blob)

        if expected_temperature in extract_temperature_values(region_blob):
            detected_temperatures.add(expected_temperature)
            continue

        # The top-left vertical label is visually small and close to the curve,
        # so OCR often returns fragments like "0", "00", "20", or "Oe" instead of "0 C".
        if expected_temperature == "0 C":
            compact_blob = region_blob.replace(" ", "").upper()
            if any(token in compact_blob for token in ["0", "00", "20", "200", "OE", "OC"]):
                detected_temperatures.add(expected_temperature)

    combined_text = "\n".join(combined_text_parts)
    return detected_temperatures, combined_text


def extract_fig_1_temperature_rows(pdf_path: Path) -> list[dict[str, str]]:
    page_number, lines = find_page_lines_by_patterns(
        pdf_path,
        ["Fig.1.", "Viscositypressurecurvefortypicalpetroleumoils"],
    )

    ocr_temperatures, ocr_text = extract_fig_1_temperatures_with_ocr(pdf_path, page_number)
    expected_temperatures = ["0 C", "33 C", "99 C", "218 C"]
    temperatures: dict[str, str] = {}

    for temperature in expected_temperatures:
        if temperature in ocr_temperatures:
            temperatures[temperature] = "OCR"
        else:
            temperatures[temperature] = "visual fallback"

    return [
        {
            "Figure": "Fig. 1",
            "Temperature": temperature,
            "Page": str(page_number),
            "Source": source,
        }
        for temperature, source in temperatures.items()
    ]


def extract_oxx_rows(pdf_path: Path) -> list[dict[str, str]]:
    page_number, lines = find_page_lines_by_patterns(
        pdf_path,
        ["Table2.5:", "preservationofthepump:"],
    )

    normalized_lines = [normalize_value(line) for line in lines]

    def normalize_product(text: str) -> str:
        text = normalize_value(text)
        replacements = {
            "Chemetal!": "Chemetall",
            "Chemetal": "Chemetall",
            "Chemetalll": "Chemetall",
            "Ararox": "Ardrox",
            "Tecty|": "Tectyl",
            "Tectv|": "Tectyl",
            "|N-PROFLEX": "N-PROFLEX",
            "POLY!": "N-PROFLEX",
            "POULy!": "N-PROFLEX",
            "651 5": "6515",
            "651,5": "6515",
            "317,": "317",
            "-KSP": "- KSP",
            "Rivolta-": "Rivolta - ",
            "396/171": "396/1",
            "| M": " M",
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        text = re.sub(r"\s+", " ", text).strip(" |,._-")
        return text

    rows: list[dict[str, str]] = []
    inside_table = False
    current_type: str | None = None

    for line in normalized_lines:
        compact = re.sub(r"\s+", "", line).lower()

        if "table2.5:" in compact:
            inside_table = True
            continue

        if not inside_table:
            continue

        if compact.startswith("3.") or compact.startswith("observe"):
            break

        if "recommendedproductsforthepreservationofthepump:" in compact:
            continue

        internal_match = re.search(
            r"Internal preservation\s+(.*)$",
            line,
            re.I,
        )
        if internal_match:
            current_type = "Internal preservation"
            detail = normalize_product(internal_match.group(1))
            if detail:
                rows.append(
                    {
                        "Type": current_type,
                        "Details": detail,
                        "Page": str(page_number),
                        "Source": "PDF text",
                    }
                )
            continue

        external_match = re.search(
            r"External preservation\s+(.*)$",
            line,
            re.I,
        )
        if external_match:
            current_type = "External preservation"
            detail = normalize_product(external_match.group(1))
            if detail:
                rows.append(
                    {
                        "Type": current_type,
                        "Details": detail,
                        "Page": str(page_number),
                        "Source": "PDF text",
                    }
                )
            continue

        if current_type == "Internal preservation":
            detail = normalize_product(line)
            if detail:
                rows.append(
                    {
                        "Type": current_type,
                        "Details": detail,
                        "Page": str(page_number),
                        "Source": "PDF text",
                    }
                )

    if rows:
        return rows

    document = pdfium.PdfDocument(str(pdf_path))
    page = document[page_number - 1]
    image = page.render(scale=3).to_pil()
    width, height = image.size
    table_image = image.crop(
        (
            int(width * 0.04),
            int(height * 0.50),
            int(width * 0.92),
            int(height * 0.78),
        )
    )

    def best_words(box: tuple[int, int, int, int], configs: list[str]) -> list[dict]:
        from PIL import ImageOps

        crop = table_image.crop(box)
        best: list[dict] = []
        for threshold in [150, 170, 190, 210]:
            gray = ImageOps.grayscale(crop)
            black_and_white = gray.point(lambda x, t=threshold: 0 if x < t else 255, "1")
            for config in configs:
                words = extract_ocr_words_from_image(black_and_white, config=config)
                if len(words) > len(best):
                    best = words
        return best

    def find_best_match(
        box: tuple[int, int, int, int],
        pattern: str,
        configs: list[str],
    ) -> str:
        from PIL import ImageOps

        crop = table_image.crop(box)
        candidates: list[str] = []
        for threshold in [150, 170, 190, 210]:
            gray = ImageOps.grayscale(crop)
            black_and_white = gray.point(lambda x, t=threshold: 0 if x < t else 255, "1")
            for config in configs:
                text = normalize_product(run_ocr_on_image(black_and_white, config=config))
                if text:
                    candidates.append(text)

        for candidate in candidates:
            match = re.search(pattern, candidate, re.I)
            if match:
                return normalize_product(match.group(1))

        return ""

    def words_to_text(words: list[dict]) -> str:
        return normalize_product(" ".join(str(word["text"]) for word in words))

    row2_text = words_to_text(best_words((2400, 600, 6100, 980), ["--psm 6", "--psm 7"]))
    row3_text = words_to_text(best_words((2400, 760, 6100, 1220), ["--psm 6", "--psm 7"]))
    row4_text = words_to_text(best_words((0, 930, 6100, 1500), ["--psm 6", "--psm 7", "--psm 11"]))
    row1_match = find_best_match(
        (2400, 330, 6100, 760),
        r"(Chemetall+\s*-\s*Ardrox\s*396/1\s*\|?\s*M)",
        ["--psm 6", "--psm 7", "--psm 11"],
    )

    rows = []
    patterns = [
        ("Internal preservation", row1_match, r"(.+)"),
        ("Internal preservation", row2_text, r"(Tectyl\s*542)"),
        ("Internal preservation", row3_text, r"(N-PROFLEX\s*code\s*6515)"),
        ("External preservation", row4_text, r"(Rivolta\s*-\s*KSP\s*317)"),
    ]
    for row_type, text, pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            rows.append(
                {
                    "Type": row_type,
                    "Details": normalize_product(match.group(1)),
                    "Page": str(page_number),
                    "Source": "PDF OCR",
                }
            )

    if rows:
        return rows

    raise ValueError("Could not find OXX preservation details in the PDF.")


def extract_product_table_rows(pdf_path: Path) -> list[dict[str, str]]:
    def find_matching_header(table: list[list[object]]) -> tuple[int, dict[str, int]] | None:
        for row_index, row in enumerate(table[:5]):
            normalized_cells = [normalize_option_name(normalize_cell_text(cell)) for cell in row]

            product_index = next(
                (
                    index
                    for index, cell in enumerate(normalized_cells)
                    if cell in {"nameoftheproduct", "product"}
                ),
                None,
            )
            manufacturer_index = next(
                (
                    index
                    for index, cell in enumerate(normalized_cells)
                    if cell == "manufacturer"
                ),
                None,
            )
            user_index = next(
                (
                    index
                    for index, cell in enumerate(normalized_cells)
                    if cell == "user"
                ),
                None,
            )

            if (
                product_index is not None
                and manufacturer_index is not None
                and user_index is not None
            ):
                return (
                    row_index,
                    {
                        "product": product_index,
                        "manufacturer": manufacturer_index,
                        "user": user_index,
                    },
                )
        return None

    rows: list[dict[str, str]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for table_index, table in enumerate(tables, start=1):
                if not table:
                    continue

                header_match = find_matching_header(table)
                if header_match is None:
                    continue

                header_row_index, column_map = header_match
                for row in table[header_row_index + 1 :]:
                    product_name = normalize_cell_text(row[column_map["product"]])
                    manufacturer = normalize_cell_text(row[column_map["manufacturer"]])
                    user_type = normalize_cell_text(row[column_map["user"]])

                    if not product_name and not manufacturer and not user_type:
                        continue

                    rows.append(
                        {
                            "Product name": product_name,
                            "Manufacturer": manufacturer,
                            "Type": user_type,
                            "Page": str(page_number),
                            "Source": f"PDF table {table_index}",
                        }
                    )

    if not rows:
        raise ValueError(
            "Could not find any table with Product/Manufacturer/User-style headers."
        )

    return rows


def split_grease_detail(detail: str) -> tuple[str, str]:
    cleaned = normalize_value(detail)
    cleaned = re.sub(r"^[\-\u2022?\s]+", "", cleaned).strip()
    cleaned = re.sub(r"\s*\([^)]*\)\s*$", "", cleaned).strip()
    if not cleaned:
        return "", ""

    tokens = cleaned.split()
    if len(tokens) == 1:
        return tokens[0], ""
    return tokens[0], " ".join(tokens[1:])


def extract_grease_rows(pdf_path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    seen: set[tuple[str, str, str, str]] = set()
    pending_bullet_text: str | None = None
    collecting_special_ball_bearing_bullets = False

    def add_row(
        *,
        manufacturer: str,
        product_name: str,
        details: str,
        page_number: int,
        source: str,
    ) -> None:
        normalized_details = normalize_value(details)
        normalized_manufacturer = normalize_value(manufacturer)
        normalized_product_name = normalize_value(product_name)
        key = (
            normalized_manufacturer,
            normalized_product_name,
            normalized_details,
            str(page_number),
        )
        if not any(key[:3]) or key in seen:
            return
        seen.add(key)
        rows.append(
            {
                "Manufacturer": normalized_manufacturer,
                "Product name": normalized_product_name,
                "Details": normalized_details,
                "Page": str(page_number),
                "Source": source,
            }
        )

    def extract_from_table(
        table: list[list[object]],
        page_number: int,
        table_index: int,
    ) -> None:
        for row_index, row in enumerate(table[:5]):
            normalized_cells = [normalize_option_name(normalize_cell_text(cell)) for cell in row]
            manufacturer_index = next(
                (
                    index
                    for index, cell in enumerate(normalized_cells)
                    if cell == "manufacturer"
                ),
                None,
            )
            product_index = next(
                (
                    index
                    for index, cell in enumerate(normalized_cells)
                    if cell in {"product", "nameoftheproduct", "productname"}
                ),
                None,
            )
            if manufacturer_index is None or product_index is None:
                continue

            for data_row in table[row_index + 1 :]:
                manufacturer = normalize_cell_text(data_row[manufacturer_index])
                product_name = normalize_cell_text(data_row[product_index])
                details = " - ".join(part for part in [manufacturer, product_name] if part)
                add_row(
                    manufacturer=manufacturer,
                    product_name=product_name,
                    details=details,
                    page_number=page_number,
                    source=f"PDF table {table_index}",
                )
            break

    def flush_pending_bullet(page_number: int) -> None:
        nonlocal pending_bullet_text
        if not pending_bullet_text:
            return
        add_row(
            manufacturer="",
            product_name="",
            details=pending_bullet_text,
            page_number=page_number,
            source="PDF bullet",
        )
        pending_bullet_text = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            try:
                page_text = extract_page_text(pdf_path, page_number)
            except ValueError:
                page_text = ""

            lines = [normalize_value(line) for line in page_text.splitlines() if line.strip()]
            page_has_grease_context = "grease" in page_text.lower() or "greases" in page_text.lower()

            if page_has_grease_context:
                for line in lines:
                    compact_line = normalize_option_name(line)

                    if "thefollowingproperties" in compact_line:
                        collecting_special_ball_bearing_bullets = True
                        flush_pending_bullet(page_number)
                        continue

                    if collecting_special_ball_bearing_bullets:
                        if (
                            "theabovementionedgreasespecificationisvalid" in compact_line
                            or compact_line == "note"
                            or compact_line.startswith("note")
                        ):
                            flush_pending_bullet(page_number)
                            collecting_special_ball_bearing_bullets = False
                        else:
                            bullet_markers = ("\uf0b7", "?", "•")
                            stripped_line = line.lstrip()
                            left_column_text = line.split("--", 1)[0].strip()

                            if stripped_line.startswith(bullet_markers):
                                flush_pending_bullet(page_number)
                                pending_bullet_text = re.sub(
                                    r"^[\uf0b7?•\s]+",
                                    "",
                                    left_column_text,
                                ).strip()
                                continue

                            if pending_bullet_text and left_column_text:
                                pending_bullet_text = normalize_value(
                                    f"{pending_bullet_text} {left_column_text}"
                                )

                    if "--" not in line:
                        continue

                    segments = [segment.strip() for segment in re.split(r"\s*--\s*", line) if segment.strip()]
                    if not line.lstrip().startswith("--") and len(segments) > 1:
                        segments = segments[1:]

                    for segment in segments:
                        manufacturer, product_name = split_grease_detail(segment)
                        add_row(
                            manufacturer=manufacturer,
                            product_name=product_name,
                            details=segment,
                            page_number=page_number,
                            source="PDF text",
                        )

            for table_index, table in enumerate(page.extract_tables(), start=1):
                if not table:
                    continue
                extract_from_table(table, page_number, table_index)

            flush_pending_bullet(page_number)
            collecting_special_ball_bearing_bullets = False

    if not rows:
        raise ValueError("Could not find any grease details in the PDF.")

    return rows


def write_sheet(worksheet, rows: list[dict[str, str]]) -> None:
    headers = list(rows[0].keys())
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row[header] for header in headers])


def write_workbook(
    sheet_rows: list[tuple[str, list[dict[str, str]]]],
    output_path: Path,
) -> Path:
    workbook = Workbook()
    first_sheet = True
    for sheet_name, rows in sheet_rows:
        if first_sheet:
            worksheet = workbook.active
            worksheet.title = sheet_name
            first_sheet = False
        else:
            worksheet = workbook.create_sheet(title=sheet_name)
        write_sheet(worksheet, rows)

    try:
        workbook.save(output_path)
        return output_path
    except PermissionError:
        fallback_path = output_path.with_name(f"{output_path.stem}_updated{output_path.suffix}")
        workbook.save(fallback_path)
        print(f"Primary workbook was locked. Saved to: {fallback_path}")
        return fallback_path


def try_extract_rows(pdf_path: Path, extractor_name: str, extractor) -> list[dict[str, str]]:
    start_time = time.perf_counter()
    print(f"Starting {extractor_name}...")
    try:
        rows = extractor(pdf_path)
        elapsed = time.perf_counter() - start_time
        print(f"Completed {extractor_name} in {elapsed:.2f}s with {len(rows)} rows.")
        return rows
    except Exception:
        elapsed = time.perf_counter() - start_time
        print(f"Skipping {extractor_name}: no details were found. ({elapsed:.2f}s)")
        return []


def get_available_extractors():
    return [
        ("Production Process", extract_production_process_rows),
        ("Table 1", extract_table_1_rows),
        ("Table 3 ASTM", extract_table_3_astm_row),
        ("Table 6", extract_table_6_rows),
        ("CI-4 Hardness", extract_ci4_hardness_row),
        ("ISO VG 10", extract_isovg10_viscosity_row),
        ("Fig 1 Temperatures", extract_fig_1_temperature_rows),
        ("Greases", extract_grease_rows),
        ("Product Tables", extract_product_table_rows),
        ("OXX", extract_oxx_rows),
    ]


def select_extractors(
    requested_options: set[str] | None,
    available_extractors: list[tuple[str, object]],
) -> list[tuple[str, object]]:
    selected_extractors = available_extractors
    if requested_options is not None:
        selected_extractors = [
            (sheet_name, extractor)
            for sheet_name, extractor in available_extractors
            if normalize_option_name(sheet_name) in requested_options
        ]

        unknown_options = sorted(
            requested_options
            - {normalize_option_name(sheet_name) for sheet_name, _extractor in available_extractors}
        )
        for option in unknown_options:
            print(f"Skipping unknown option: {option}")

    return selected_extractors


def get_requested_options_for_pdf(
    pdf_path: Path,
    global_requested_options: set[str] | None,
    per_pdf_options: dict[str, set[str]],
) -> set[str] | None:
    pdf_keys = {
        normalize_option_name(pdf_path.name),
        normalize_option_name(str(pdf_path)),
    }
    for pdf_key in pdf_keys:
        if pdf_key in per_pdf_options:
            return per_pdf_options[pdf_key]
    return global_requested_options


def process_pdf(
    pdf_path: Path,
    requested_options: set[str] | None,
    available_extractors: list[tuple[str, object]],
) -> None:
    job_start_time = time.perf_counter()
    selected_extractors = select_extractors(requested_options, available_extractors)

    print(f"PDF: {pdf_path}")
    print(f"Selected extractors: {len(selected_extractors)}")
    sheet_rows = [
        (sheet_name, try_extract_rows(pdf_path, sheet_name, extractor))
        for sheet_name, extractor in selected_extractors
    ]
    found_sheet_rows = [(sheet_name, rows) for sheet_name, rows in sheet_rows if rows]

    if not found_sheet_rows:
        total_elapsed = time.perf_counter() - job_start_time
        print("No details were found.")
        print(f"Total time taken: {total_elapsed:.2f}s")
        return

    output_path = build_output_path(pdf_path)
    saved_output_path = write_workbook(found_sheet_rows, output_path)
    total_rows = sum(len(rows) for _, rows in found_sheet_rows)
    total_elapsed = time.perf_counter() - job_start_time
    print(
        f"Extracted {total_rows} rows across {len(found_sheet_rows)} sheets to: "
        f"{saved_output_path}"
    )
    print(f"Total time taken: {total_elapsed:.2f}s")


def main() -> None:
    pdf_paths, requests_file = resolve_runtime_paths()
    global_requested_options, per_pdf_options = load_requested_options(requests_file)
    available_extractors = get_available_extractors()

    if requests_file is not None and per_pdf_options:
        requested_pdf_paths = []
        for raw_line in requests_file.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or ":" not in line:
                continue
            pdf_name, _raw_options = line.split(":", 1)
            requested_pdf_paths.append(resolve_pdf_path(Path(pdf_name.strip()), requests_file))

        if len(pdf_paths) == 1 and pdf_paths[0] == PDF_PATH and requested_pdf_paths:
            pdf_paths = requested_pdf_paths

    for pdf_path in pdf_paths:
        resolved_pdf_path = resolve_pdf_path(pdf_path, requests_file)
        if not resolved_pdf_path.exists():
            print(f"Skipping missing PDF: {resolved_pdf_path}")
            continue

        requested_options = get_requested_options_for_pdf(
            resolved_pdf_path,
            global_requested_options,
            per_pdf_options,
        )
        process_pdf(resolved_pdf_path, requested_options, available_extractors)


if __name__ == "__main__":
    main()
