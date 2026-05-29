import unicodedata
from pathlib import Path

import pandas as pd

from extract import extract_data


def _normalize_text(value):
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    return text.lower().strip()


def _normalized_columns(df):
    return {_normalize_text(column): column for column in df.columns}


def _find_table(workbook_tables, sheet_name, required_columns):
    required = {_normalize_text(column) for column in required_columns}
    for table_info in workbook_tables.get(sheet_name, []):
        df = table_info["data"]
        columns = set(_normalized_columns(df).keys())
        if required.issubset(columns):
            return df
    return None


def _get_column(df, column_name):
    columns = _normalized_columns(df)
    return columns.get(_normalize_text(column_name))


def _text_mask(series, pattern):
    return series.astype(str).str.contains(pattern, case=False, na=False)


def _numeric_series(df, column_name):
    column = _get_column(df, column_name)
    if column is None:
        return pd.Series(dtype="float64")
    return pd.to_numeric(df[column], errors="coerce")


def _single_value(df, row_mask, column_name):
    column = _get_column(df, column_name)
    if column is None:
        return None
    values = pd.to_numeric(df.loc[row_mask, column], errors="coerce")
    if values.empty:
        return None
    return values.iloc[0]


def _table_folio_name_grade_fee(df):
    required = {"folio", "nombre del alumno", "grado", "mensualidad"}
    columns = set(_normalized_columns(df).keys())
    if not required.issubset(columns):
        return pd.DataFrame()

    result = pd.DataFrame(
        {
            "Folio": df[_get_column(df, "Folio")],
            "Nombre del alumno": df[_get_column(df, "Nombre del alumno")],
            "Grado": df[_get_column(df, "Grado")],
            "Mensualidad": pd.to_numeric(df[_get_column(df, "Mensualidad")], errors="coerce"),
        }
    )
    return result.dropna(how="all")


def generate_discrepancy_report(workbook_tables):
    lines = []

    balanza_marzo = _find_table(workbook_tables, "Balanza Marzo", ["Cuenta", "Descripción", "Saldo final"])
    banco_marzo = _find_table(workbook_tables, "Banco Marzo", ["Fecha", "Concepto", "Depósito", "Retiro", "Saldo"])
    operativo_niveles = _find_table(workbook_tables, "Operativo Marzo", ["Nivel", "Facturado del mes", "Cobrado del mes"])
    antiguedad = _find_table(workbook_tables, "Operativo Marzo", ["Antigüedad", "Saldo marzo"])
    estatus_alumno = _find_table(workbook_tables, "Operativo Marzo", ["Folio", "Nombre del alumno", "Grado", "Mensualidad"])

    # Bank reconciliation
    if balanza_marzo is not None and banco_marzo is not None:
        saldo_final_col = _get_column(balanza_marzo, "Saldo final")
        cargos_col = _get_column(balanza_marzo, "Cargos del período")
        abonos_col = _get_column(balanza_marzo, "Abonos del período")
        descripcion_col = _get_column(balanza_marzo, "Descripción")

        bank_saldo_col = _get_column(banco_marzo, "Saldo")
        bank_deposito_col = _get_column(banco_marzo, "Depósito")
        bank_retiro_col = _get_column(banco_marzo, "Retiro")

        balanza_bancos_mask = _text_mask(balanza_marzo[descripcion_col], r"bancos|1110")
        balanza_saldo = pd.to_numeric(balanza_marzo.loc[balanza_bancos_mask, saldo_final_col], errors="coerce").iloc[0]
        balanza_cargos = pd.to_numeric(balanza_marzo.loc[balanza_bancos_mask, cargos_col], errors="coerce").iloc[0]
        balanza_abonos = pd.to_numeric(balanza_marzo.loc[balanza_bancos_mask, abonos_col], errors="coerce").iloc[0]

        banco_saldo = pd.to_numeric(banco_marzo[bank_saldo_col], errors="coerce").dropna().iloc[-1]
        banco_depositos = pd.to_numeric(banco_marzo[bank_deposito_col], errors="coerce").sum(min_count=1)
        banco_retiros = pd.to_numeric(banco_marzo[bank_retiro_col], errors="coerce").sum(min_count=1)

        lines.append("BANK RECONCILIATION:")
        lines.append(f"  Balanza reported saldo final: {balanza_saldo}")
        lines.append(f"  Bank statement saldo final: {banco_saldo}")
        lines.append(f"  Balanza cargos (entradas): {balanza_cargos}")
        lines.append(f"  Bank deposits: {banco_depositos}")
        lines.append(f"  Balanza abonos (salidas): {balanza_abonos}")
        lines.append(f"  Bank withdrawals: {banco_retiros}")
        lines.append(f"  Difference (bank - balanza): {banco_saldo - balanza_saldo}")
        lines.append("")
    else:
        lines.append("BANK RECONCILIATION: insufficient data to perform check")
        lines.append("")

    # Secondary income discrepancy
    if operativo_niveles is not None and balanza_marzo is not None:
        nivel_col = _get_column(operativo_niveles, "Nivel")
        facturado_col = _get_column(operativo_niveles, "Facturado del mes")
        cobrado_col = _get_column(operativo_niveles, "Cobrado del mes")

        secundaria_mask = _text_mask(operativo_niveles[nivel_col], r"^secundaria$")
        secundaria_facturado = pd.to_numeric(operativo_niveles.loc[secundaria_mask, facturado_col], errors="coerce").iloc[0]
        secundaria_cobrado = pd.to_numeric(operativo_niveles.loc[secundaria_mask, cobrado_col], errors="coerce").iloc[0]

        descripcion_col = _get_column(balanza_marzo, "Descripción")
        abonos_col = _get_column(balanza_marzo, "Abonos del período")
        colegiatura_mask = _text_mask(balanza_marzo[descripcion_col], r"colegiaturas secundaria")
        secundaria_abono = pd.to_numeric(balanza_marzo.loc[colegiatura_mask, abonos_col], errors="coerce").iloc[0]

        lines.append("SECONDARY INCOME CHECK:")
        lines.append(f"  Operative facturado del mes: {secundaria_facturado}")
        lines.append(f"  Operative cobrado del mes: {secundaria_cobrado}")
        lines.append(f"  Accounting abono in Balanza Marzo: {secundaria_abono}")
        lines.append(f"  Difference (facturado - abono): {secundaria_facturado - secundaria_abono}")
        lines.append("")
    else:
        lines.append("SECONDARY INCOME CHECK: insufficient data to perform check")
        lines.append("")

    # Client portfolio (Cartera)
    if balanza_marzo is not None and antiguedad is not None:
        descripcion_col = _get_column(balanza_marzo, "Descripción")
        saldo_final_col = _get_column(balanza_marzo, "Saldo final")
        cartera_mask = _text_mask(balanza_marzo[descripcion_col], r"clientes \(cartera de alumnos\)")
        cartera_balanza = pd.to_numeric(balanza_marzo.loc[cartera_mask, saldo_final_col], errors="coerce").iloc[0]

        antiguedad_label_col = _get_column(antiguedad, "Antigüedad")
        saldo_marzo_col = _get_column(antiguedad, "Saldo marzo")
        total_mask = _text_mask(antiguedad[antiguedad_label_col], r"^total$")
        cartera_operativa = pd.to_numeric(antiguedad.loc[total_mask, saldo_marzo_col], errors="coerce").iloc[0]

        lines.append("CLIENTS (CARTERA) CHECK:")
        lines.append(f"  Balanza Marzo client saldo final: {cartera_balanza}")
        lines.append(f"  Operativo Marzo antiguedad TOTAL saldo marzo: {cartera_operativa}")
        lines.append(f"  Difference (operativo - balanza): {cartera_operativa - cartera_balanza}")
        lines.append("")
    else:
        lines.append("CLIENTS (CARTERA) CHECK: insufficient data to perform check")
        lines.append("")

    # Month-over-month balance comparison
    balanza_febrero = _find_table(workbook_tables, "Balanza Febrero", ["Cuenta", "Descripción", "Saldo final"])
    if balanza_febrero is not None and balanza_marzo is not None:
        cuenta_feb_col = _get_column(balanza_febrero, "Cuenta")
        descripcion_feb_col = _get_column(balanza_febrero, "Descripción")
        saldo_feb_col = _get_column(balanza_febrero, "Saldo final")

        cuenta_mar_col = _get_column(balanza_marzo, "Cuenta")
        descripcion_mar_col = _get_column(balanza_marzo, "Descripción")
        saldo_mar_col = _get_column(balanza_marzo, "Saldo final")

        feb_base = balanza_febrero.loc[:, [cuenta_feb_col, descripcion_feb_col, saldo_feb_col]].rename(
            columns={
                cuenta_feb_col: "Cuenta",
                descripcion_feb_col: "Descripción Febrero",
                saldo_feb_col: "Saldo febrero",
            }
        )
        mar_base = balanza_marzo.loc[:, [cuenta_mar_col, descripcion_mar_col, saldo_mar_col]].rename(
            columns={
                cuenta_mar_col: "Cuenta",
                descripcion_mar_col: "Descripción Marzo",
                saldo_mar_col: "Saldo marzo",
            }
        )

        feb_base["Saldo febrero"] = pd.to_numeric(feb_base["Saldo febrero"], errors="coerce")
        mar_base["Saldo marzo"] = pd.to_numeric(mar_base["Saldo marzo"], errors="coerce")

        merged_balance = feb_base.merge(mar_base, on="Cuenta", how="outer")
        merged_balance["Descripción"] = merged_balance["Descripción Marzo"].combine_first(merged_balance["Descripción Febrero"])
        merged_balance["Saldo febrero"] = merged_balance["Saldo febrero"].fillna(0)
        merged_balance["Saldo marzo"] = merged_balance["Saldo marzo"].fillna(0)
        merged_balance["Variación absoluta"] = merged_balance["Saldo marzo"] - merged_balance["Saldo febrero"]
        merged_balance["Variación porcentual"] = merged_balance["Variación absoluta"].div(
            merged_balance["Saldo febrero"].replace(0, pd.NA)
        ).mul(100)

        new_balance_mask = (merged_balance["Saldo febrero"] == 0) & (merged_balance["Saldo marzo"] > 0)
        disappeared_balance_mask = (merged_balance["Saldo febrero"] > 0) & (merged_balance["Saldo marzo"] == 0)
        atypical_balance_mask = (
            merged_balance["Saldo febrero"].ne(0)
            & merged_balance["Variación absoluta"].abs().gt(15000)
            & merged_balance["Variación porcentual"].abs().gt(20)
        )

        def _format_money(value):
            if pd.isna(value):
                return "0"
            return f"{value:,.0f}"

        def _format_pct(value):
            if pd.isna(value):
                return "n/a"
            return f"{value:.2f}%"

        lines.append("MONTH-OVER-MONTH BALANCE CHECK:")

        new_rows = merged_balance.loc[new_balance_mask].sort_values("Saldo marzo", ascending=False)
        if new_rows.empty:
            lines.append("  No new balances detected.")
        else:
            lines.append("  Saldos nuevos:")
            for row in new_rows.to_dict(orient="records"):
                lines.append(
                    "  - Cuenta {cuenta} | {descripcion} | Febrero: {febrero} | Marzo: {marzo} | Variación: {variacion} | %: {porcentaje}".format(
                        cuenta=row["Cuenta"],
                        descripcion=row["Descripción"],
                        febrero=_format_money(row["Saldo febrero"]),
                        marzo=_format_money(row["Saldo marzo"]),
                        variacion=_format_money(row["Variación absoluta"]),
                        porcentaje=_format_pct(row["Variación porcentual"]),
                    )
                )

        disappeared_rows = merged_balance.loc[disappeared_balance_mask].sort_values("Saldo febrero", ascending=False)
        if disappeared_rows.empty:
            lines.append("  No disappeared balances detected.")
        else:
            lines.append("  Saldos desaparecidos:")
            for row in disappeared_rows.to_dict(orient="records"):
                lines.append(
                    "  - Cuenta {cuenta} | {descripcion} | Febrero: {febrero} | Marzo: {marzo} | Variación: {variacion} | %: {porcentaje}".format(
                        cuenta=row["Cuenta"],
                        descripcion=row["Descripción"],
                        febrero=_format_money(row["Saldo febrero"]),
                        marzo=_format_money(row["Saldo marzo"]),
                        variacion=_format_money(row["Variación absoluta"]),
                        porcentaje=_format_pct(row["Variación porcentual"]),
                    )
                )

        atypical_rows = merged_balance.loc[atypical_balance_mask].assign(
            abs_variacion=lambda frame: frame["Variación absoluta"].abs()
        ).sort_values("abs_variacion", ascending=False)
        if atypical_rows.empty:
            lines.append("  No atypical variations detected.")
        else:
            lines.append("  Variaciones atípicas:")
            for row in atypical_rows.to_dict(orient="records"):
                lines.append(
                    "  - Cuenta {cuenta} | {descripcion} | Febrero: {febrero} | Marzo: {marzo} | Variación: {variacion} | %: {porcentaje}".format(
                        cuenta=row["Cuenta"],
                        descripcion=row["Descripción"],
                        febrero=_format_money(row["Saldo febrero"]),
                        marzo=_format_money(row["Saldo marzo"]),
                        variacion=_format_money(row["Variación absoluta"]),
                        porcentaje=_format_pct(row["Variación porcentual"]),
                    )
                )

        lines.append("")
    else:
        lines.append("MONTH-OVER-MONTH BALANCE CHECK: insufficient data to perform check")
        lines.append("")

    # Student-level fee validation
    if estatus_alumno is not None:
        validation_levels = pd.DataFrame(
            {
                "Nivel esperable": ["Preescolar", "Primaria", "Secundaria", "Prepa"],
                "Cuota esperada": [7500, 9500, 11500, 13500],
            }
        )

        students = _table_folio_name_grade_fee(estatus_alumno)
        students["Nivel base"] = (
            students["Grado"].astype(str).str.extract(r"^(Preescolar|Primaria|Secundaria|Prepa|Preparatoria)", expand=False)
            .replace({"Preparatoria": "Prepa"})
        )

        merged = students.merge(validation_levels, left_on="Nivel base", right_on="Nivel esperable", how="left")
        merged["Mensualidad"] = pd.to_numeric(merged["Mensualidad"], errors="coerce")
        mismatches = merged.loc[
            merged["Mensualidad"].notna()
            & merged["Cuota esperada"].notna()
            & (merged["Mensualidad"] != merged["Cuota esperada"])
        ].copy()

        lines.append("STUDENT-LEVEL FEE CHECK:")
        if mismatches.empty:
            lines.append("  No cobro errors detected in student monthly fees.")
        else:
            lines.append("  Students with mismatched monthly fee:")
            lines.extend(
                mismatches.loc[:, ["Folio", "Nombre del alumno", "Grado", "Mensualidad", "Cuota esperada"]]
                .assign(
                    Mensualidad=lambda frame: frame["Mensualidad"].astype("Int64"),
                    **{"Cuota esperada": lambda frame: frame["Cuota esperada"].astype("Int64")},
                )
                .to_string(index=False)
                .splitlines()
            )
        lines.append("")
    else:
        lines.append("STUDENT-LEVEL FEE CHECK: insufficient data to perform check")
        lines.append("")

    return "\n".join(lines)


if __name__ == "__main__":
    file_path = "San_Patricio_Cierre_Marzo_2025.xlsx"
    output_dir = Path("discrepancies")
    output_dir.mkdir(exist_ok=True)

    data = extract_data(file_path)

    if data:
        report = generate_discrepancy_report(data)
        report_path = output_dir / "discrepancy_report.txt"
        with report_path.open("w", encoding="utf-8") as f:
            f.write(report)
        print(f"Discrepancy report written to {report_path}")
