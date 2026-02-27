#!/usr/bin/env python3
"""Construye una tabla de generación mensual por sitio desde un XLSX unificado."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List

import pandas as pd

MONTH_NAME_BY_NUM = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

MONTH_TEXT_MAP = {
    "1": 1,
    "01": 1,
    "ene": 1,
    "enero": 1,
    "jan": 1,
    "january": 1,
    "2": 2,
    "02": 2,
    "feb": 2,
    "febrero": 2,
    "february": 2,
    "3": 3,
    "03": 3,
    "mar": 3,
    "marzo": 3,
    "march": 3,
    "4": 4,
    "04": 4,
    "abr": 4,
    "abril": 4,
    "apr": 4,
    "april": 4,
    "5": 5,
    "05": 5,
    "may": 5,
    "mayo": 5,
    "6": 6,
    "06": 6,
    "jun": 6,
    "junio": 6,
    "june": 6,
    "7": 7,
    "07": 7,
    "jul": 7,
    "julio": 7,
    "july": 7,
    "8": 8,
    "08": 8,
    "ago": 8,
    "agosto": 8,
    "aug": 8,
    "august": 8,
    "9": 9,
    "09": 9,
    "sep": 9,
    "sept": 9,
    "septiembre": 9,
    "september": 9,
    "10": 10,
    "oct": 10,
    "octubre": 10,
    "october": 10,
    "11": 11,
    "nov": 11,
    "noviembre": 11,
    "november": 11,
    "12": 12,
    "dic": 12,
    "diciembre": 12,
    "dec": 12,
    "december": 12,
}

MONTH_COL_KEYWORDS = ("month", "mes", "period", "periodo")
TIME_COL_KEYWORDS = ("timestamp", "datetime", "date", "fecha", "time", "hora")
GEN_COL_KEYWORDS = (
    "grid power",
    "ac power",
    "generation",
    "generacion",
    "energ",
    "production",
    "produccion",
    "yield",
    "ac energy",
    "kwh",
    "mwh",
)


def _normalize_text(text: str) -> str:
    return " ".join(str(text).strip().lower().replace("_", " ").split())


def _detect_month_column(df: pd.DataFrame) -> str | None:
    for col in df.columns:
        name = _normalize_text(col)
        if any(keyword in name for keyword in MONTH_COL_KEYWORDS):
            return col
    return None


def _detect_time_column(df: pd.DataFrame) -> str | None:
    for col in df.columns:
        name = _normalize_text(col)
        if any(keyword in name for keyword in TIME_COL_KEYWORDS):
            return col

    for col in df.columns:
        parsed = pd.to_datetime(df[col], errors="coerce")
        if parsed.notna().mean() > 0.8:
            return col
    return None


def _detect_generation_column(df: pd.DataFrame) -> str | None:
    # Prioridad explícita para perfiles HelioScope de potencia horaria.
    priority_tokens = ("grid power", "grid_power", "ac power", "ac_power")
    for token in priority_tokens:
        for col in df.columns:
            if token in _normalize_text(col):
                return col

    best_col = None
    best_score = -1
    for col in df.columns:
        name = _normalize_text(col)
        score = sum(1 for keyword in GEN_COL_KEYWORDS if keyword in name)
        if score > best_score:
            best_score = score
            best_col = col

    if best_col is not None and best_score > 0:
        return best_col

    numeric_candidates = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
    if len(numeric_candidates) == 1:
        return numeric_candidates[0]
    return None


def _kwh_factor_from_column_name(col: str) -> float:
    name = _normalize_text(col)
    if "mwh" in name:
        return 1000.0
    if "kwh" in name:
        return 1.0
    if "kw" in name and "kwh" not in name:
        return 1.0
    if "watt" in name or "power" in name or name == "w":
        return 1.0 / 1000.0
    return 1.0


def _parse_month(value) -> int | None:
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        month_num = int(value)
        return month_num if 1 <= month_num <= 12 else None

    text = _normalize_text(str(value))
    if not text:
        return None

    if text in MONTH_TEXT_MAP:
        return MONTH_TEXT_MAP[text]

    dt = pd.to_datetime(text, errors="coerce")
    if not pd.isna(dt):
        return int(dt.month)

    return None


def build_monthly_generation_table(input_xlsx: Path) -> tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    xls = pd.ExcelFile(input_xlsx)
    long_rows: List[Dict] = []
    warnings: List[str] = []

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
        except Exception as exc:
            warnings.append(f"No se pudo leer hoja '{sheet}': {exc}")
            continue

        if df.empty:
            warnings.append(f"Hoja '{sheet}' vacía; se omite")
            continue

        month_col = _detect_month_column(df)
        time_col = _detect_time_column(df) if month_col is None else None
        if month_col is None and time_col is None:
            warnings.append(f"Hoja '{sheet}' sin columna de mes/timestamp reconocible; se omite")
            continue

        generation_col = _detect_generation_column(df)
        if generation_col is None:
            warnings.append(f"Hoja '{sheet}' sin columna de generación/energía reconocible; se omite")
            continue

        if month_col is not None:
            temp = df[[month_col, generation_col]].copy()
            temp["month_num"] = temp[month_col].apply(_parse_month)
        else:
            temp = df[[time_col, generation_col]].copy()
            temp["_dt"] = pd.to_datetime(temp[time_col], errors="coerce")
            temp["month_num"] = temp["_dt"].dt.month

        raw_generation = pd.to_numeric(
            temp[generation_col].astype(str).str.replace(",", "", regex=False),
            errors="coerce",
        )
        factor = _kwh_factor_from_column_name(generation_col)
        temp["generation"] = raw_generation * factor

        temp = temp.dropna(subset=["month_num", "generation"])

        if temp.empty:
            warnings.append(f"Hoja '{sheet}' no tiene filas válidas de mes+generación; se omite")
            continue

        grouped = temp.groupby("month_num", as_index=False)["generation"].sum()
        site_name = str(sheet).strip()
        for _, row in grouped.iterrows():
            month_num = int(row["month_num"])
            long_rows.append(
                {
                    "Sitio": site_name,
                    "Mes_Num": month_num,
                    "Mes": MONTH_NAME_BY_NUM[month_num],
                    "Generacion_kWh": float(row["generation"]),
                }
            )

    long_df = pd.DataFrame(long_rows)
    if long_df.empty:
        raise RuntimeError("No se pudo construir la tabla mensual. Revisa que el XLSX tenga columnas de mes/timestamp y generación.")

    month_order = pd.DataFrame(
        {
            "Mes_Num": list(MONTH_NAME_BY_NUM.keys()),
            "Mes": list(MONTH_NAME_BY_NUM.values()),
        }
    )

    pivot_df = (
        long_df.pivot_table(index="Mes_Num", columns="Sitio", values="Generacion_kWh", aggfunc="sum")
        .reset_index()
        .merge(month_order, on="Mes_Num", how="left")
    )

    site_columns = sorted([c for c in pivot_df.columns if c not in {"Mes_Num", "Mes"}])
    pivot_df = pivot_df[["Mes_Num", "Mes", *site_columns]]
    pivot_df = pivot_df.sort_values("Mes_Num").reset_index(drop=True)

    return pivot_df, long_df.sort_values(["Mes_Num", "Sitio"]).reset_index(drop=True), warnings


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Genera tabla de generación mensual por sitio desde un XLSX unificado de CSVs."
    )
    parser.add_argument("--input", type=Path, required=True, help="Ruta del XLSX unificado de entrada.")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("generacion_mensual_por_sitio.xlsx"),
        help="Ruta del XLSX de salida.",
    )
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    output_path = args.output.expanduser().resolve()

    if not input_path.exists() or not input_path.is_file():
        raise FileNotFoundError(f"No existe el archivo de entrada: {input_path}")

    pivot_df, long_df, warnings = build_monthly_generation_table(input_path)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        pivot_df.to_excel(writer, sheet_name="tabla_mensual", index=False)
        long_df.to_excel(writer, sheet_name="detalle_largo", index=False)
        if warnings:
            pd.DataFrame({"warning": warnings}).to_excel(writer, sheet_name="warnings", index=False)

    print(f"Excel generado: {output_path}")
    print(f"Filas tabla mensual: {len(pivot_df)}")
    print(f"Filas detalle: {len(long_df)}")
    if warnings:
        print("Incluye hoja: warnings")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
