# San Patricio Dashboard

Este proyecto genera un entregable final en Excel + Dashboard a partir de tablas extraídas desde un archivo fuente de cierre.

## Archivos principales

- `extract.py`: extrae tablas del XLSX original y genera CSVs en `extracted_tables/`.
- `check_discrepancies.py`: genera `discrepancies/discrepancy_report.txt`.
- `excel_compiler.py`: compila `San_Patricio_Dashboard_Marzo_2025.xlsx` leyendo los CSVs y el reporte.

## Requisitos

- Python 3.12+
- `pandas`
- `openpyxl`

## Uso

```powershell
python extract.py
python check_discrepancies.py
python excel_compiler.py
```

El archivo final se guarda como `San_Patricio_Dashboard_Marzo_2025.xlsx` en la raíz del proyecto.
