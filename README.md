# HelioScope Tools

GUI con 3 herramientas:

1. **Reporte unificado desde PDFs**
- Ejecuta `generar_reporte_unificado.py`
- Genera un único `.xlsx` con hojas:
  - `consolidado`
  - `temperature_model_parameters`
  - `warnings` (si hay)

2. **Unificar CSVs por carpeta**
- Ejecuta `unificar_csvs_por_carpeta.py`
- Busca subcarpetas con `*.csv`
- Genera un `.xlsx` con una hoja por carpeta

3. **Tabla de generación mensual por sitio**
- Ejecuta `tabla_generacion_mensual.py`
- Usa el `.xlsx` unificado de CSVs (una hoja por sitio)
- Detecta columnas de mes y generación/energía
- Genera un `.xlsx` con hojas:
  - `tabla_mensual` (Mes x Sitio)
  - `detalle_largo`
  - `warnings` (si hay)

## Ejecutar GUI

```bash
cd /home/panda/Apps/Helioscope_Tools
python3 helioscope_tools_gui.py
```

## Requisitos

- Python 3.10+
- `pandas`
- `xlsxwriter`
- `PyMuPDF` (`fitz`)
- `pdftotext` instalado en el sistema

## Uso rápido

1. Abre la GUI.
2. Selecciona la carpeta raíz de trabajo.
3. Para PDFs: define nombre de salida y clic en **Generar reporte unificado**.
4. Para CSVs: define nombre de salida y clic en **Unificar CSVs**.
5. Para tabla mensual: define XLSX de entrada/salida y clic en **Generar tabla mensual**.
6. Revisa el log en la parte inferior.
