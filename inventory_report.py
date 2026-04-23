from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

# Ruta principal en Windows y ruta secundaria actual.
BASE_PATH_CANDIDATES = [
    Path("C:/Users/cesar/OneDrive/Documentos"),
    Path("/Users/pgcesaare/OneDrive/Documentos"),
]

RANCH_FILES = {
    "Gold Star Cattle": "Gold Star Inventory.xlsx",
    "La Esperanza Ranch": "Inventory at Dominguez - Guess Cattle.xlsx",
    "Frias Ranch": "Inventory at Frias - Guess Cattle.xlsx",
}

COLUMNS = [
    ("Breed", 28),
    ("Quantity", 14),
    ("Avg. Price", 14),
    ("Avg. DOF", 14),
    ("Min Date", 14),
    ("Max Date", 14),
    ("Total", 14),
]


def resolve_base_path() -> Path:
    for base_path in BASE_PATH_CANDIDATES:
        if base_path.exists():
            return base_path

    return BASE_PATH_CANDIDATES[0]


BASE_PATH = resolve_base_path()
OUTPUT_DIR = BASE_PATH / "Inventory Reports"


def load_ranch_file(filename: str) -> pd.DataFrame:
    return pd.read_excel(BASE_PATH / filename)


def filter_inventory(df: pd.DataFrame) -> pd.DataFrame:
    mask = (df["Ownership"] == "Brandao Cattle") & (df["Status"] == "Feeding")
    return df.loc[mask].copy()


def build_inventory(df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        df.groupby(by="Breed")
        .agg(
            quantity=("Breed", "size"),
            avg_price=("Purchase Price", "mean"),
            avg_DOF=("DOF", "mean"),
            min_date=("Date In", "min"),
            max_date=("Date In", "max"),
            total=("Purchase Price", "sum"),
        )
        .sort_values(by="total", ascending=False)
    )

    return summary


def load_inventory_assignments() -> dict[str, pd.DataFrame]:
    inventories = {}

    for ranch_name, filename in RANCH_FILES.items():
        ranch_df = load_ranch_file(filename)
        filtered_df = filter_inventory(ranch_df)
        inventories[ranch_name] = build_inventory(filtered_df)

    return inventories


def build_output_path() -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    report_date = pd.Timestamp.today().strftime("%m.%d.%Y")
    filename = f"brandao cattle inventory report {report_date}.xlsx"
    return OUTPUT_DIR / filename


def apply_sheet_styles(ws) -> None:
    for column_letter, (_, width) in zip("ABCDEFG", COLUMNS):
        ws.column_dimensions[column_letter].width = width

    ws.sheet_view.showGridLines = False


def write_headers(ws) -> None:
    ws["A1"] = "BRANDAO CATTLE"
    ws["A1"].font = Font(bold=True, size=15)

    ws["A2"] = "INVENTORY REPORT"
    ws["A2"].font = Font(size=13)

    ws["A3"] = '="DATE: " & TEXT(TODAY(), "mm/dd/yyyy")'
    ws["A3"].font = Font(size=12)
    ws["B3"] = None


def write_table_header(ws, row_number: int, ranch_name: str) -> int:
    thin_gray = Side(style="thin", color="808080")

    ws.cell(row=row_number, column=1, value=ranch_name).font = Font(bold=False, size=13)

    header_row = row_number + 1

    for column_index, (header, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=column_index, value=header)
        cell.font = Font(bold=True, size=12)
        if header == "Breed":
            alignment = "left"
        elif header == "Total":
            alignment = "right"
        else:
            alignment = "center"
        cell.alignment = Alignment(horizontal=alignment)
        cell.border = Border(bottom=thin_gray)

    return header_row


def write_inventory_rows(ws, start_row: int, inventory_df: pd.DataFrame) -> int:
    current_row = start_row

    for breed, values in inventory_df.iterrows():
        ws.cell(row=current_row, column=1, value=breed)
        ws.cell(row=current_row, column=2, value=int(values["quantity"]))
        ws.cell(row=current_row, column=3, value=float(values["avg_price"]))
        ws.cell(row=current_row, column=4, value=float(values["avg_DOF"]))
        ws.cell(row=current_row, column=5, value=values["min_date"])
        ws.cell(row=current_row, column=6, value=values["max_date"])
        ws.cell(row=current_row, column=7, value=float(values["total"]))
        current_row += 1

    return current_row


def format_data_rows(ws, first_row: int, last_row: int) -> None:
    if last_row < first_row:
        return

    for row in range(first_row, last_row + 1):
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="left")
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=3).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=5).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=6).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=7).alignment = Alignment(horizontal="right")

        ws.cell(row=row, column=2).number_format = "#,##0"
        ws.cell(row=row, column=3).number_format = "$#,##0.00"
        ws.cell(row=row, column=4).number_format = "0.00"
        ws.cell(row=row, column=5).number_format = "mm/dd/yyyy"
        ws.cell(row=row, column=6).number_format = "mm/dd/yyyy"
        ws.cell(row=row, column=7).number_format = "$#,##0.00"


def write_table_totals(ws, total_row: int, data_start_row: int, data_end_row: int) -> int:
    if data_end_row >= data_start_row:
        quantity_formula = f"=SUM(B{data_start_row}:B{data_end_row})"
        total_formula = f"=SUM(G{data_start_row}:G{data_end_row})"
    else:
        quantity_formula = "=0"
        total_formula = "=0"

    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=total_row, column=2, value=quantity_formula).font = Font(bold=True)
    ws.cell(row=total_row, column=7, value=total_formula).font = Font(bold=True)

    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left")
    ws.cell(row=total_row, column=2).alignment = Alignment(horizontal="center")
    ws.cell(row=total_row, column=7).alignment = Alignment(horizontal="right")

    ws.cell(row=total_row, column=2).number_format = "#,##0"
    ws.cell(row=total_row, column=7).number_format = "$#,##0.00"

    return total_row


def write_ranch_section(ws, start_row: int, ranch_name: str, inventory_df: pd.DataFrame) -> tuple[int, int]:
    header_row = write_table_header(ws, start_row, ranch_name)
    data_start_row = header_row + 1
    next_row = write_inventory_rows(ws, data_start_row, inventory_df)
    data_end_row = next_row - 1

    format_data_rows(ws, data_start_row, data_end_row)

    total_row = next_row if not inventory_df.empty else data_start_row
    write_table_totals(ws, total_row, data_start_row, data_end_row)

    return total_row + 4, total_row


def write_global_total(ws, row_number: int, total_rows: list[int]) -> None:
    total_formula = "=" + "+".join(f"G{row}" for row in total_rows) if total_rows else "=0"

    ws.cell(row=row_number, column=5, value="TOTAL").font = Font(bold=True, size=14)
    ws.cell(row=row_number, column=7, value=total_formula).font = Font(bold=True, size=14)
    ws.cell(row=row_number, column=7).number_format = "$#,##0.00"


def generate_inventory_report(inventories: dict[str, pd.DataFrame], output_path: Path | None = None) -> Path:
    if output_path is None:
        output_path = build_output_path()

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Inventory Report"

    apply_sheet_styles(worksheet)
    write_headers(worksheet)

    current_row = 8
    total_rows = []

    for ranch_name, inventory_df in inventories.items():
        current_row, total_row = write_ranch_section(worksheet, current_row, ranch_name, inventory_df)
        total_rows.append(total_row)

    write_global_total(worksheet, current_row, total_rows)
    workbook.save(output_path)

    return output_path


inventory_assignments = load_inventory_assignments()

# Variables finales para usar en otros scripts.
gold_star_inv = inventory_assignments["Gold Star Cattle"]
la_esperanza_inv = inventory_assignments["La Esperanza Ranch"]
frias_ranch_inv = inventory_assignments["Frias Ranch"]


if __name__ == "__main__":
    report_path = generate_inventory_report(inventory_assignments)
    print(f"Reporte creado en: {report_path}")
