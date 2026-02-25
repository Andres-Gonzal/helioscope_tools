# HelioScope Tools

GUI con 2 herramientas:

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
5. Revisa el log en la parte inferior.
