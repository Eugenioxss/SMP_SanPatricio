import pandas as pd
import re
from pathlib import Path


def _safe_name(value):
    return "".join(character if character.isalnum() or character in ("-", "_") else "_" for character in str(value)).strip("_")


def _row_blocks(sheet_df):
    blocks = []
    start = None

    for row_index, has_data in sheet_df.notna().any(axis=1).items():
        if has_data and start is None:
            start = row_index
        elif not has_data and start is not None:
            blocks.append((start, row_index - 1))
            start = None

    if start is not None:
        blocks.append((start, len(sheet_df) - 1))

    return blocks


def _is_header_row(row, min_header_cells=2):
    non_empty = [value for value in row.tolist() if pd.notna(value)]

    if len(non_empty) < min_header_cells:
        return False

    for value in non_empty:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return False

    return True


def _is_two_column_summary_block(block):
    non_empty_counts = block.notna().sum(axis=1)

    if non_empty_counts.empty:
        return False

    if non_empty_counts.max() > 2:
        return False

    two_cell_rows = block[non_empty_counts == 2]
    if len(two_cell_rows) < 2:
        return False

    numeric_total_rows = 0
    for _, row in two_cell_rows.iterrows():
        values = [value for value in row.tolist() if pd.notna(value)]
        if len(values) != 2:
            continue

        total_value = values[1]
        if isinstance(total_value, (int, float)) and not isinstance(total_value, bool):
            numeric_total_rows += 1

    return numeric_total_rows >= 2 and numeric_total_rows == len(two_cell_rows)


def _is_title_text(value: object) -> bool:
    if not isinstance(value, str):
        return False
    v = value.strip()
    if not v:
        return False
    if re.match(r"^\s*\d+\.\s*", v):
        return True
    if "resumen" in v.lower():
        return True
    return False


def extract_tables_from_sheet(sheet_df, sheet_name, min_header_cells=2):
    tables = []
    canonical_columns = None

    for block_number, (start_row, end_row) in enumerate(_row_blocks(sheet_df), start=1):
        block = sheet_df.iloc[start_row : end_row + 1].reset_index(drop=True)

        header_row_index = None
        for row_index in range(len(block)):
            if _is_header_row(block.iloc[row_index], min_header_cells=min_header_cells):
                header_row_index = row_index
                break

        if header_row_index is None:
            if _is_two_column_summary_block(block):
                data = block.iloc[:, :2].copy()
                data = data.dropna(how="all")

                if data.empty:
                    continue

                # remove title-like rows (e.g. '5. Resumen general del mes') when the second col is empty
                first_col = data.iloc[:, 0]
                second_col = data.iloc[:, 1]
                title_mask = first_col.apply(_is_title_text) & (
                    second_col.isna() | (second_col.astype(str).str.strip() == "")
                )
                if title_mask.any():
                    data = data.loc[~title_mask].copy()

                if data.empty:
                    continue

                data.columns = ["concept", "total"]
                data = data.reset_index(drop=True)

                tables.append(
                    {
                        "sheet_name": sheet_name,
                        "block_number": block_number,
                        "start_row": start_row + 1,
                        "end_row": end_row + 1,
                        "data": data,
                    }
                )
                continue

            if canonical_columns is None:
                continue

            data = block.dropna(how="all").copy()
            if data.empty:
                continue

            data = data.iloc[:, : len(canonical_columns)]
            data.columns = canonical_columns[: len(data.columns)]
            data = data.reset_index(drop=True)

            tables.append(
                {
                    "sheet_name": sheet_name,
                    "block_number": block_number,
                    "start_row": start_row + 1,
                    "end_row": end_row + 1,
                    "data": data,
                }
            )
            continue

        if header_row_index >= len(block) - 1:
            continue

        header = block.iloc[header_row_index]
        header_mask = header.notna()

        if header_mask.sum() < min_header_cells:
            continue

        data = block.iloc[header_row_index + 1 :, header_mask].copy()
        data = data.dropna(how="all")

        if data.empty:
            continue

        data.columns = [str(value).strip() for value in header[header_mask].tolist()]
        data = data.reset_index(drop=True)

        if canonical_columns is None:
            canonical_columns = list(data.columns)

        tables.append(
            {
                "sheet_name": sheet_name,
                "block_number": block_number,
                "start_row": start_row + 1,
                "end_row": end_row + 1,
                "data": data,
            }
        )

    return tables


def extract_data(file_path):
    try:
        sheets = pd.read_excel(file_path, sheet_name=None, header=None)
        workbook_tables = {}

        for sheet_name, sheet_df in sheets.items():
            workbook_tables[sheet_name] = extract_tables_from_sheet(sheet_df, sheet_name)

        print(f"Data successfully extracted from {file_path}")
        return workbook_tables
    except Exception as e:
        print(f"An error occurred while extracting data: {e}")
        return None
        # finished exporting CSVs


if __name__ == "__main__":
    file_path = "San_Patricio_Cierre_Marzo_2025.xlsx"
    output_dir = Path("extracted_tables")
    output_dir.mkdir(exist_ok=True)

    data = extract_data(file_path)

    if data:
        for sheet_name, tables in data.items():
            print(f"Hoja: {sheet_name} -> {len(tables)} tablas detectadas")

            safe_sheet_name = _safe_name(sheet_name)
            for index, table_info in enumerate(tables, start=1):
                table = table_info["data"]
                # skip truly empty tables (all NaN) or tables with no meaningful cells
                if table.dropna(how="all").empty:
                    print(f"  Skipping empty tabla {index} for hoja {sheet_name}")
                    # remove existing file if any
                    existing = output_dir / f"{safe_sheet_name}__tabla_{index}.csv"
                    try:
                        if existing.exists():
                            existing.unlink()
                    except Exception:
                        pass
                    continue

                # skip tables that are just a single note row with no numeric content
                table_df = pd.DataFrame(table)
                numeric_df = table_df.apply(lambda col: pd.to_numeric(col, errors="coerce"))
                numeric_cells = int(numeric_df.notna().sum().sum())
                non_empty_rows = int(table_df.dropna(how="all").shape[0])
                if numeric_cells == 0 and non_empty_rows <= 1:
                    print(f"  Skipping note-like tabla {index} for hoja {sheet_name} (no numeric content)")
                    existing = output_dir / f"{safe_sheet_name}__tabla_{index}.csv"
                    try:
                        if existing.exists():
                            existing.unlink()
                    except Exception:
                        pass
                    continue

                output_path = output_dir / f"{safe_sheet_name}__tabla_{index}.csv"
                table.to_csv(output_path, index=False)
                print(
                    f"  Tabla {index}: filas {table_info['start_row']}-{table_info['end_row']} -> {output_path}"
                )

        # finished exporting CSVs