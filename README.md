# San Patricio Financial Dashboard

Este proyecto implementa un pipeline de datos automatizado para procesar el cierre financiero mensual del Colegio San Patricio. A partir de la extracción de datos en bruto, el sistema genera dos entregables de nivel ejecutivo: un reporte estructurado en Excel y un dashboard web interactivo con proyecciones de riesgo.

## Características principales

- **ETL automatizado:** Extracción y limpieza de datos desde los archivos fuente del cierre operativo y contable.
- **Doble entregable:** Reporte `.xlsx` formateado (listo para auditoría) y una aplicación web interactiva.
- **Seguridad:** Autenticación en el dashboard para proteger la confidencialidad financiera.
- **Modelo predictivo:** Heurística basada en Cadenas de Markov para proyecciones de cartera vencida (30/60/90+ días).
- **Visualización:** Interfaz moderna con gráficos en Plotly.

## Arquitectura de archivos

- `extract.py` — Extrae las tablas operativas y contables y genera los CSV en `extracted_tables/`.
- `check_discrepancies.py` — Cruza información contable vs. operativa y escribe `discrepancies/discrepancy_report.txt`.
- `excel_compiler.py` — Compila KPIs y genera `San_Patricio_Dashboard_Marzo_2025.xlsx`.
- `app.py` — Interfaz web (Streamlit) que consume los datos procesados.

## Requisitos

- Python 3.12+
- pandas
- openpyxl
- streamlit
- plotly

Las dependencias exactas están en `requirements.txt`.

## Uso (local)

1. Ejecuta el pipeline de procesamiento en este orden (desde PowerShell o la terminal de tu preferencia):

```powershell
python extract.py
python check_discrepancies.py
python excel_compiler.py
```

Nota: El archivo final de Excel se guardará en la raíz del proyecto.

2. Para ejecutar el dashboard localmente:

```powershell
streamlit run app.py
```

### Credenciales de acceso (entorno de desarrollo)

Por motivos de seguridad, la aplicación requiere autenticación. Credenciales de ejemplo para la vista de `Controller`:

- Usuario: `controller`
- Contraseña: `SanPatricio2025`

Si vas a publicar el repositorio o desplegar en producción, considera eliminar o rotar estas credenciales y usar variables de entorno o un secret manager.

## Despliegue en producción

El proyecto está preparado para despliegue (CI/CD) en plataformas como Render. Comando de arranque recomendado:

```bash
streamlit run app.py
```

## Pasos para actualizar en GitHub / Render

1. Añade el archivo modificado y crea el commit:

```bash
git add README.md
git commit -m "Actualizar README con formato y guía de uso"
git push origin main
```

2. En Render o la plataforma que uses, asegúrate de que el comando de start sea `streamlit run app.py` y que las dependencias de `requirements.txt` estén instaladas.

---

Si quieres, aplico cambios adicionales (ej.: quitar credenciales, añadir badges, o traducción completa al inglés). Indica qué prefieres.