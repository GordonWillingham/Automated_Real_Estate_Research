import zipfile
from xml.sax.saxutils import escape

HEADERS = [
    "Address",
    "Price",
    "Beds",
    "Baths",
    "Sqft",
    "Lot size",
    "Year built",
    "HOA",
    "Days on market",
    "Listing URL",
    "Source",
    "Notes",
    "Change",
    "Resale Value Score",
]

ROWS = [
    [
        "N/A – live search unavailable in offline environment",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        (
            "Unable to collect Charlotte listings (north of I-485, $300k-$450k, 3bd/3ba+) because "
            "internet access is blocked in this environment; cannot compute resale value heuristics."
        ),
        "",
        "0 (placeholder)",
    ]
]


def col_letter(idx: int) -> str:
    """Convert zero-based column index to Excel column letters."""
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def build_sheet_xml(headers, rows) -> str:
    cells = []
    # Header row
    row_idx = 1
    row_cells = []
    for col_idx, header in enumerate(headers):
        cell_ref = f"{col_letter(col_idx)}{row_idx}"
        row_cells.append(f"<c r=\"{cell_ref}\" t=\"inlineStr\"><is><t>{escape(header)}</t></is></c>")
    cells.append(f"<row r=\"{row_idx}\">{''.join(row_cells)}</row>")

    # Data rows
    for data in rows:
        row_idx += 1
        row_cells = []
        for col_idx, value in enumerate(data):
            cell_ref = f"{col_letter(col_idx)}{row_idx}"
            text = escape(str(value)) if value is not None else ""
            row_cells.append(f"<c r=\"{cell_ref}\" t=\"inlineStr\"><is><t>{text}</t></is></c>")
        cells.append(f"<row r=\"{row_idx}\">{''.join(row_cells)}</row>")

    sheet_data = "".join(cells)
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        f"<sheetData>{sheet_data}</sheetData>"
        "</worksheet>"
    )


def write_xlsx(path: str, headers, rows) -> None:
    content_types = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "</Types>"
    )

    rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>"
    )

    workbook_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        "</Relationships>"
    )

    workbook = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Listings\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
        "</workbook>"
    )

    sheet = build_sheet_xml(headers, rows)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet)


if __name__ == "__main__":
    write_xlsx("charlotte_listings.xlsx", HEADERS, ROWS)
    print("Wrote charlotte_listings.xlsx with placeholder row.")
