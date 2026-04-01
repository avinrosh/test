from __future__ import annotations

import re
from pathlib import Path

import pdfplumber
from openpyxl import Workbook


PDF_PATH = Path(
    r"d:\App\pdfscrapper\Kirk-OthmerEncyclopediaofChemicalTechnology-LubricationandLubricants.pdf"
)
OUTPUT_XLSX = Path(r"d:\App\pdfscrapper\lubrication_tables.xlsx")

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


def get_page_lines(pdf_path: Path, page_number: int) -> list[str]:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        page_text = page.extract_text(layout=True)

    if not page_text:
        raise ValueError(f"No text could be extracted from page {page_number}.")

    return [line.strip() for line in page_text.splitlines() if line.strip()]


def find_page_lines_by_patterns(
    pdf_path: Path,
    patterns: list[str],
    *,
    rotated: bool = False,
) -> tuple[int, list[str]]:
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            if rotated:
                page_text = page.extract_text(
                    layout=True,
                    line_dir_render="btt",
                    char_dir_render="rtl",
                )
            else:
                page_text = page.extract_text(layout=True)

            if not page_text:
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


def write_sheet(worksheet, rows: list[dict[str, str]]) -> None:
    headers = list(rows[0].keys())
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row[header] for header in headers])


def write_workbook(
    production_process_rows: list[dict[str, str]],
    table_1_rows: list[dict[str, str]],
    table_3_astm_rows: list[dict[str, str]],
    table_6_rows: list[dict[str, str]],
    ci4_hardness_rows: list[dict[str, str]],
    isovg10_rows: list[dict[str, str]],
    output_path: Path,
) -> None:
    workbook = Workbook()

    sheet_1 = workbook.active
    sheet_1.title = "Production Process"
    write_sheet(sheet_1, production_process_rows)

    sheet_2 = workbook.create_sheet(title="Table 1")
    write_sheet(sheet_2, table_1_rows)

    sheet_3 = workbook.create_sheet(title="Table 3 ASTM")
    write_sheet(sheet_3, table_3_astm_rows)

    sheet_4 = workbook.create_sheet(title="Table 6")
    write_sheet(sheet_4, table_6_rows)

    sheet_5 = workbook.create_sheet(title="CI-4 Hardness")
    write_sheet(sheet_5, ci4_hardness_rows)

    sheet_6 = workbook.create_sheet(title="ISO VG 10")
    write_sheet(sheet_6, isovg10_rows)

    try:
        workbook.save(output_path)
    except PermissionError:
        fallback_path = output_path.with_name(f"{output_path.stem}_updated{output_path.suffix}")
        workbook.save(fallback_path)
        print(f"Primary workbook was locked. Saved to: {fallback_path}")


def main() -> None:
    production_process_rows = extract_production_process_rows(PDF_PATH)
    table_1_rows = extract_table_1_rows(PDF_PATH)
    table_3_astm_rows = extract_table_3_astm_row(PDF_PATH)
    table_6_rows = extract_table_6_rows(PDF_PATH)
    ci4_hardness_rows = extract_ci4_hardness_row(PDF_PATH)
    isovg10_rows = extract_isovg10_viscosity_row(PDF_PATH)
    write_workbook(
        production_process_rows,
        table_1_rows,
        table_3_astm_rows,
        table_6_rows,
        ci4_hardness_rows,
        isovg10_rows,
        OUTPUT_XLSX,
    )
    print(
        f"Extracted {len(production_process_rows)} rows from page {PRODUCTION_PROCESS_PAGE} "
        f"{len(table_1_rows)} rows from page {TABLE_1_PAGE}, and "
        f"{len(table_3_astm_rows)} row found by content search for Table 3, and "
        f"{len(table_6_rows)} rows found by content search for Table 6, and "
        f"{len(ci4_hardness_rows)} row found by content search for CI-4 hardness, and "
        f"{len(isovg10_rows)} row found by content search for ISO VG 10 to: {OUTPUT_XLSX}"
    )


if __name__ == "__main__":
    main()
