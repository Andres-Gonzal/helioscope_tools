#!/usr/bin/env python3
"""Genera un reporte consolidado HelioScope en un solo XLSX."""
from __future__ import annotations

import argparse
import json
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator, List, Sequence, Tuple

import fitz  # PyMuPDF
import pandas as pd


# ---------- Constantes y regex ----------
RE_KW = re.compile(r"([\d.,]+)\s*kW", re.IGNORECASE)
RE_NAME_COUNT = re.compile(
    r"(?P<name>.*?)(?P<count>\d[\d,]*)\s*\((?P<kw>[\d.,]+)(?:\s*\([^)]+\))?\s*kW\)",
    re.IGNORECASE,
)
RE_REMOVE_NODIGIT_PARENS = re.compile(r"\([^0-9)]*\)")
RE_NUMBER = re.compile(r"(-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?)")
RE_PERCENT = re.compile(r"(-?\d+(?:\.\d+)?)\s*%")
RE_TEMP_DELTA = re.compile(r"(-?\d+(?:\.\d+)?)\s*°?\s*C", re.IGNORECASE)

STOP_INVERTERS = {"AC Panels", "AC Home", "Strings", "Module", "AC", "Runs"}
STOP_MODULES = {
    "Description",
    "Racking",
    "Orientation",
    "Tilt",
    "Azimuth",
    "Intrarow",
    "Frame",
    "Frames",
    "Modules",
    "Power",
    "Field Segment",
    "©",
}
IGNORED_NAME_KEYWORDS = {
    "wiring",
    "zone",
    "flush",
    "mount",
    "landscape",
    "portrait",
    "racking",
    "orientation",
    "tilt",
    "azimuth",
    "spacing",
    "frame",
    "modules",
    "power",
    "segment",
    "inverters",
}
IGNORED_CHARS = set("0123456789-./°")

DESIRED_ORDER = [
    "Suburbia Lindavista",
    "Suburbia Torres Lindavista",
    "Suburbia Toluca Gran Plaza",
    "Suburbia Ermita Iztapalapa",
    "Suburbia Coacalco",
    "Suburbia Puerta Texcoco",
    "Suburbia Ciudad Victoria",
    "Suburbia Monterrey Santa Catarina",
    "Suburbia Cuautitlán Centella",
    "Suburbia Paseo Gomez Palacio",
    "Suburbia Tuxtepec",
    "Suburbia Monterrey Paseo Juarez",
    "Suburbia Tizayuca Tizara",
    "Suburbia Campeche",
    "Suburbia Los Reyes Tepozan",
]
ORDER_INDEX = {name: idx for idx, name in enumerate(DESIRED_ORDER)}


@dataclass(frozen=True)
class TemperatureModelRow:
    rack_type: str
    a: float
    b: float
    temperature_delta: str


# ---------- Utilidades ----------
def _find_pdfs(root: Path) -> list[Path]:
    return sorted(p.resolve() for p in root.rglob("*.pdf") if p.is_file())


def _run_pdftotext(pdf: Path) -> Sequence[str]:
    result = subprocess.run(
        ["pdftotext", str(pdf), "-"],
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout.splitlines()


def _normalise_project_name(raw: str) -> str:
    text = raw.strip()
    if not text:
        return text
    lower = text.lower()
    if lower.startswith("suburbia"):
        rest = text[len("suburbia") :].strip()
        text = "Suburbia"
        if rest:
            text += " "
            text += rest.title() if rest.isupper() else rest
    elif text.isupper():
        text = text.title()
    return " ".join(text.split())


def _parse_first_number(text: str) -> float | None:
    for match in RE_NUMBER.finditer(text):
        start, end = match.span()
        if start > 0 and text[start - 1].isalpha():
            continue
        if end < len(text) and text[end].isalpha():
            continue
        raw = match.group(0).replace(",", "")
        try:
            return float(raw)
        except ValueError:
            continue
    return None


def _extract_project_name(lines: Sequence[str], pdf: Path, warnings: List[str]) -> str:
    for idx, line in enumerate(lines):
        if line.strip().lower() == "project name":
            for next_line in lines[idx + 1 :]:
                candidate = next_line.strip()
                if candidate:
                    return _normalise_project_name(candidate)
            break
    warnings.append(f"No se pudo detectar el nombre del proyecto en {pdf}")
    return _normalise_project_name(pdf.stem)


def _extract_coordinates(lines: Sequence[str], pdf: Path, warnings: List[str]) -> str | None:
    for idx, line in enumerate(lines):
        label = line.strip().lower()
        if label == "project address" or (
            label == "project"
            and idx + 1 < len(lines)
            and lines[idx + 1].strip().lower() == "address"
        ):
            start_idx = idx + 1 if label == "project address" else idx + 2
            coords: List[float] = []
            for candidate in lines[start_idx : start_idx + 6]:
                value = _parse_first_number(candidate)
                if value is None:
                    continue
                coords.append(value)
                if len(coords) == 2:
                    return f"{coords[0]}, {coords[1]}"
            warnings.append(f"No se pudieron leer coordenadas en {pdf}")
            return None
    return None


# ---------- Bloque concentrado ----------
def _canonical_installation_type(raw: str) -> str:
    text = " ".join(raw.lower().split())
    if "flush" in text and "mount" in text:
        return "Flush Mount"
    if "fixed" in text and "tilt" in text:
        return "Fixed Tilt"
    if "carport" in text:
        return "Carport"
    if "east" in text and "west" in text:
        return "East-West"
    return ""


def _extract_installation_types(lines: Sequence[str]) -> List[str]:
    start = None
    for idx, line in enumerate(lines):
        if "field segments" in line.strip().lower():
            start = idx
            break
    if start is None:
        return []

    section: List[str] = []
    for line in lines[start + 1 :]:
        lower = line.strip().lower()
        if "detailed layout" in lower:
            break
        section.append(line)
    section_text = " ".join(section)

    candidates = [
        _canonical_installation_type("Flush Mount")
        if re.search(r"flush\s+mount", section_text, re.IGNORECASE)
        else "",
        _canonical_installation_type("Fixed Tilt")
        if re.search(r"fixed\s+tilt", section_text, re.IGNORECASE)
        else "",
        _canonical_installation_type("Carport")
        if re.search(r"carport", section_text, re.IGNORECASE)
        else "",
        _canonical_installation_type("East-West")
        if re.search(r"east\s*-\s*west", section_text, re.IGNORECASE)
        else "",
    ]
    if not candidates[0] and re.search(
        r"field\s+segment(?:\s+\d+)?\s+flush(?:\s+\d+)?(?:\s+mount)?",
        section_text,
        re.IGNORECASE,
    ):
        candidates[0] = "Flush Mount"
    if not candidates[1] and re.search(
        r"field\s+segment(?:\s+\d+)?\s+fixed(?:\s+\d+)?\s+tilt",
        section_text,
        re.IGNORECASE,
    ):
        candidates[1] = "Fixed Tilt"

    unique: List[str] = []
    for item in candidates:
        if item and item not in unique:
            unique.append(item)
    return unique


def _slice_component_table(lines: Sequence[str]) -> Sequence[str]:
    table_start = None
    for idx, line in enumerate(lines):
        if line.strip() == "Component":
            context = "".join(lines[max(0, idx - 5) : idx + 1])
            if "Components" in context:
                table_start = idx
                break
    if table_start is None:
        return []

    table = list(lines[table_start:])
    for idx, line in enumerate(table):
        stripped = line.strip()
        if stripped.startswith("© 2025") or "Detailed Layout" in stripped:
            return table[:idx]
    return table


def _collect_blocks(table_lines: Sequence[str], label: str, stop_tokens: Iterable[str]) -> List[str]:
    stop_tokens_lower = tuple(token.lower() for token in stop_tokens)
    blocks: List[str] = []
    i = 0
    n = len(table_lines)
    while i < n:
        if table_lines[i].strip() == label:
            i += 1
            while i < n and table_lines[i].strip() == "":
                i += 1
            while i < n:
                current = table_lines[i].strip()
                if not current:
                    i += 1
                    continue
                lower = current.lower()
                if lower.startswith("") or any(lower.startswith(token) for token in stop_tokens_lower):
                    return blocks
                block_lines: List[str] = []
                while i < n and table_lines[i].strip():
                    block_lines.append(table_lines[i].strip())
                    i += 1
                block_text = " ".join(block_lines).strip()
                if block_text:
                    blocks.append(block_text)
                while i < n and table_lines[i].strip() == "":
                    i += 1
            break
        i += 1
    return blocks


def _should_ignore_name_block(text_lower: str) -> bool:
    if not text_lower:
        return True
    if all(char in IGNORED_CHARS for char in text_lower.replace(" ", "")):
        return True
    return any(keyword in text_lower for keyword in IGNORED_NAME_KEYWORDS)


def _normalise_blocks(blocks: Sequence[str], pdf: Path, warnings: List[str]) -> List[dict]:
    normalised: List[dict] = []
    pending_name_parts: List[str] = []

    for block in blocks:
        text = " ".join(block.split())
        if "kW" not in text:
            if _should_ignore_name_block(text.lower()):
                continue
            pending_name_parts.append(text)
            continue

        name_prefix = " ".join(pending_name_parts).strip()
        pending_name_parts = []
        sanitised = RE_REMOVE_NODIGIT_PARENS.sub("", text)
        candidate = sanitised if not name_prefix else f"{name_prefix} {sanitised}"
        match = RE_NAME_COUNT.search(candidate)
        if not match:
            warnings.append(f'No se pudo interpretar el bloque "{block}" en {pdf}')
            continue

        normalised.append(
            {
                "name": match.group("name").strip(),
                "count_text": f"{match.group('count')} ({match.group('kw')} kW)",
                "count": int(match.group("count").replace(",", "")),
            }
        )

    return normalised


def _extract_concentrado_row(pdf: Path, warnings: List[str]) -> dict | None:
    try:
        lines = _run_pdftotext(pdf)
    except subprocess.CalledProcessError as exc:
        warnings.append(f"pdftotext falló en {pdf}: {exc.stderr.strip()}")
        return None

    project_name = _extract_project_name(lines, pdf, warnings)
    capacity_kwp = None
    for idx, line in enumerate(lines):
        if "Module DC" in line:
            for next_line in lines[idx : idx + 25]:
                match = RE_KW.search(next_line)
                if match:
                    capacity_kwp = float(match.group(1).replace(",", ""))
                    break
            if capacity_kwp is not None:
                break
    if capacity_kwp is None:
        warnings.append(f"No se encontró la capacidad Module DC Nameplate en {pdf}")

    coordinates = _extract_coordinates(lines, pdf, warnings)
    installation_types = "; ".join(_extract_installation_types(lines))

    table_lines = _slice_component_table(lines)
    if not table_lines:
        warnings.append(f"No se localizó la tabla de componentes en {pdf}")
        return {
            "Project Name": project_name,
            "Capacity (kWp)": capacity_kwp,
            "Project Coordinates": coordinates,
            "Installation Types": installation_types,
            "Inverter Names": "",
            "Inverter Counts": "",
            "Total Inverters": None,
            "Module Names": "",
            "Module Counts": "",
            "Total Modules": None,
            "PDF Path": str(pdf),
        }

    inverters = _normalise_blocks(_collect_blocks(table_lines, "Inverters", STOP_INVERTERS), pdf, warnings)
    modules = _normalise_blocks(_collect_blocks(table_lines, "Module", STOP_MODULES), pdf, warnings)

    return {
        "Project Name": project_name,
        "Capacity (kWp)": capacity_kwp,
        "Project Coordinates": coordinates,
        "Installation Types": installation_types,
        "Inverter Names": "; ".join(entry["name"] for entry in inverters if entry["name"]),
        "Inverter Counts": "; ".join(entry["count_text"] for entry in inverters if entry["count_text"]),
        "Total Inverters": sum(entry["count"] for entry in inverters if entry["count"] is not None),
        "Module Names": "; ".join(entry["name"] for entry in modules if entry["name"]),
        "Module Counts": "; ".join(entry["count_text"] for entry in modules if entry["count_text"]),
        "Total Modules": sum(entry["count"] for entry in modules if entry["count"] is not None),
        "PDF Path": str(pdf),
    }


def build_concentrado(root: Path) -> Tuple[pd.DataFrame, List[str]]:
    warnings: List[str] = []
    rows: List[dict] = []
    for pdf in _find_pdfs(root):
        data = _extract_concentrado_row(pdf, warnings)
        if data:
            rows.append(data)
    if not rows:
        raise RuntimeError("No se pudo extraer información de los PDFs")
    rows.sort(key=lambda row: row["Project Name"])
    return pd.DataFrame(rows), warnings


# ---------- Bloque GHI ----------
def _find_annual_ghi(lines: Sequence[str], start_idx: int) -> float | None:
    for line in lines[start_idx : start_idx + 20]:
        value = _parse_first_number(line)
        if value is not None:
            return value
    return None


def _extract_soiling_average(lines: Sequence[str], pdf: Path, warnings: List[str]) -> float | None:
    for idx, line in enumerate(lines):
        if "soiling (%)" not in line.lower():
            continue
        values: List[float] = []
        collecting = False
        for candidate in lines[idx + 1 :]:
            if len(values) >= 12:
                break
            stripped = candidate.strip()
            if not stripped:
                continue
            if not collecting:
                if stripped.upper() == "J":
                    collecting = True
                continue
            if len(stripped) == 1 and stripped.upper() in "JFMAMJJASOND":
                continue
            value = _parse_first_number(stripped)
            if value is not None:
                values.append(value)
        if len(values) >= 12:
            values = values[:12]
            return sum(values) / len(values)
        warnings.append(f"Solo se encontraron {len(values)} valores de soiling en {pdf} (se esperaban 12)")
        return None
    warnings.append(f"No se localizó la gráfica de soiling en {pdf}")
    return None


def _extract_ghi_row(pdf: Path, warnings: List[str]) -> dict | None:
    try:
        lines = _run_pdftotext(pdf)
    except subprocess.CalledProcessError as exc:
        warnings.append(f"pdftotext falló en {pdf}: {exc.stderr.strip()}")
        return None

    project_name = _extract_project_name(lines, pdf, warnings)
    ghi_value = None
    for idx, line in enumerate(lines):
        if "annual global horizontal" in line.strip().lower():
            ghi_value = _find_annual_ghi(lines, idx)
            if ghi_value is not None:
                break
    if ghi_value is None:
        warnings.append(f"No se encontró la Annual Global Horizontal Irradiance en {pdf}")

    return {
        "Project Name": project_name,
        "Annual Global Horizontal Irradiance (kWh/m2)": ghi_value,
        "Average Soiling (%)": _extract_soiling_average(lines, pdf, warnings),
        "Project Coordinates": _extract_coordinates(lines, pdf, warnings),
        "PDF Path": str(pdf),
    }


def build_ghi_concentrado(root: Path) -> Tuple[pd.DataFrame, List[str]]:
    warnings: List[str] = []
    rows: List[dict] = []
    for pdf in _find_pdfs(root):
        data = _extract_ghi_row(pdf, warnings)
        if data:
            rows.append(data)
    if not rows:
        raise RuntimeError("No se pudo extraer información de los PDFs")
    rows.sort(key=lambda row: (ORDER_INDEX.get(row["Project Name"], len(DESIRED_ORDER)), row["Project Name"]))
    return pd.DataFrame(rows), warnings


# ---------- Bloque métricas (PyMuPDF) ----------
def _iter_pdf_lines(pdf: Path) -> Iterator[str]:
    doc = fitz.open(pdf)
    try:
        for page in doc:
            text = page.get_text("text") or ""
            for line in text.splitlines():
                yield line
    finally:
        doc.close()


def _parse_percent(text: str) -> float | None:
    match = RE_PERCENT.search(text)
    if not match:
        return None
    try:
        return float(match.group(1))
    except ValueError:
        return None


def _find_value_after_label(lines: Sequence[str], label_predicate, max_lookahead: int = 20) -> float | None:
    for idx, line in enumerate(lines):
        if label_predicate(line):
            for candidate in lines[idx + 1 : idx + 1 + max_lookahead]:
                value = _parse_first_number(candidate)
                if value is not None:
                    return value
    return None


def _extract_poa_irradiance(lines: Sequence[str]) -> float | None:
    for idx, line in enumerate(lines):
        lower = line.strip().lower()
        if "poa irradiance" in lower:
            for candidate in lines[idx + 1 : idx + 21]:
                value = _parse_first_number(candidate)
                if value is not None:
                    return value
            continue
        if lower == "poa":
            j = idx + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines) and lines[j].strip().lower() == "irradiance":
                for candidate in lines[j + 1 : j + 21]:
                    value = _parse_first_number(candidate)
                    if value is not None:
                        return value
    return None


def _extract_performance_ratio(lines: Sequence[str]) -> float | None:
    for idx, line in enumerate(lines):
        if line.strip().lower() != "performance":
            continue
        j = idx + 1
        while j < len(lines) and not lines[j].strip():
            j += 1
        if j >= len(lines) or lines[j].strip().lower() != "ratio":
            continue
        k = j + 1
        while k < len(lines) and not lines[k].strip():
            k += 1
        if k < len(lines):
            value = _parse_percent(lines[k])
            if value is not None:
                return value
    for idx, line in enumerate(lines):
        if "performance ratio" in line.lower():
            value = _parse_percent(line)
            if value is not None:
                return value
            for candidate in lines[idx + 1 : idx + 6]:
                value = _parse_percent(candidate)
                if value is not None:
                    return value
    return None


def _extract_temperature_model_parameters(lines: Sequence[str]) -> List[TemperatureModelRow]:
    stop_tokens = (
        "soiling",
        "irradiation variance",
        "cell temperature spread",
        "module & component",
        "condition set",
        "annual production report",
        "©",
    )
    for idx, line in enumerate(lines):
        if line.strip().lower() != "rack type":
            continue
        header_end = None
        for j in range(idx, min(len(lines), idx + 12)):
            if lines[j].strip().lower() == "temperature delta":
                header_end = j
                break
        if header_end is None:
            continue
        block: List[str] = []
        for raw in lines[header_end + 1 :]:
            cleaned = raw.strip()
            if not cleaned:
                continue
            lower = cleaned.lower()
            if any(lower.startswith(token) for token in stop_tokens):
                break
            block.append(cleaned)
        rows: List[TemperatureModelRow] = []
        i = 0
        while i + 3 < len(block):
            rack_type = block[i].strip()
            a_val = _parse_first_number(block[i + 1])
            b_val = _parse_first_number(block[i + 2])
            delta_text = block[i + 3]
            if a_val is not None and b_val is not None and RE_TEMP_DELTA.search(delta_text):
                rows.append(
                    TemperatureModelRow(
                        rack_type=rack_type,
                        a=a_val,
                        b=b_val,
                        temperature_delta=delta_text,
                    )
                )
                i += 4
                continue
            i += 1
        if rows:
            return rows
    return []


def extract_metrics(pdf: Path) -> tuple[dict, List[TemperatureModelRow], List[str]]:
    warnings: List[str] = []
    lines = list(_iter_pdf_lines(pdf))
    project_name = _extract_project_name(lines, pdf, warnings)
    annual_ghi = _find_value_after_label(
        lines,
        lambda l: "annual global horizontal" in l.lower() or "annual global horizonal" in l.lower(),
    )
    if annual_ghi is None:
        warnings.append("No se encontró Annual Global Horizontal Irradiance")
    poa = _extract_poa_irradiance(lines)
    if poa is None:
        warnings.append("No se encontró POA Irradiance")
    performance_ratio = _extract_performance_ratio(lines)
    if performance_ratio is None:
        warnings.append("No se encontró Performance Ratio")
    temp_rows = _extract_temperature_model_parameters(lines)
    if not temp_rows:
        warnings.append("No se encontró la tabla Temperature Model Parameters")

    return (
        {
            "pdf": str(pdf),
            "project_name": project_name,
            "annual_global_horizontal_irradiance_kwh_m2": annual_ghi,
            "poa_irradiance_kwh_m2": poa,
            "performance_ratio_percent": performance_ratio,
        },
        temp_rows,
        warnings,
    )


def _build_metrics_tables(root: Path) -> tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    metrics_rows: List[dict] = []
    temp_rows: List[dict] = []
    warnings: List[str] = []

    for pdf in _find_pdfs(root):
        metrics, temps, local_warnings = extract_metrics(pdf)
        metrics_rows.append(
            {
                "PDF Path": str(pdf),
                "Project Name": metrics.get("project_name"),
                "Annual GHI (PyMuPDF) (kWh/m2)": metrics.get("annual_global_horizontal_irradiance_kwh_m2"),
                "POA Irradiance (kWh/m2)": metrics.get("poa_irradiance_kwh_m2"),
                "Performance Ratio (%)": metrics.get("performance_ratio_percent"),
            }
        )
        for row in temps:
            temp_rows.append(
                {
                    "PDF Path": str(pdf),
                    "Project Name": metrics.get("project_name"),
                    "rack_type": row.rack_type,
                    "a": row.a,
                    "b": row.b,
                    "temperature_delta": row.temperature_delta,
                }
            )
        for warn in local_warnings:
            warnings.append(f"{pdf}: {warn}")

    return pd.DataFrame(metrics_rows), pd.DataFrame(temp_rows), warnings


# ---------- Merge final ----------
def _coalesce_project_name(df: pd.DataFrame) -> pd.DataFrame:
    name_cols = [col for col in df.columns if col.startswith("Project Name")]
    if "Project Name" not in df.columns:
        df["Project Name"] = None
    for col in name_cols:
        if col == "Project Name":
            continue
        df["Project Name"] = df["Project Name"].combine_first(df[col])
    drop_cols = [col for col in name_cols if col != "Project Name"]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


def _coalesce_coordinates(df: pd.DataFrame) -> pd.DataFrame:
    if "Project Coordinates" not in df.columns:
        df["Project Coordinates"] = None
    if "Project Coordinates (GHI)" in df.columns:
        df["Project Coordinates"] = df["Project Coordinates"].combine_first(df["Project Coordinates (GHI)"])
        df = df.drop(columns=["Project Coordinates (GHI)"])
    return df


def _build_temperature_for_installation(merged: pd.DataFrame, temp_df: pd.DataFrame) -> pd.DataFrame:
    if merged.empty:
        merged["Temperature Params (Selected Installation Types)"] = ""
        return merged
    if temp_df.empty:
        merged["Temperature Params (Selected Installation Types)"] = ""
        return merged

    temp_df = temp_df.copy()
    temp_df["Project Name"] = temp_df["Project Name"].astype(str)
    temp_df["rack_type_norm"] = temp_df["rack_type"].astype(str).str.strip().str.lower()

    params_col: List[str] = []
    for _, row in merged.iterrows():
        project_name = str(row.get("Project Name") or "").strip()
        selected_types = [
            part.strip()
            for part in str(row.get("Installation Types") or "").split(";")
            if part.strip()
        ]
        selected_norm = {x.lower() for x in selected_types}

        project_rows = temp_df[temp_df["Project Name"] == project_name]
        if project_rows.empty and row.get("PDF Path"):
            project_rows = temp_df[temp_df["PDF Path"] == row["PDF Path"]]
        if project_rows.empty:
            params_col.append("")
            continue
        if selected_norm:
            project_rows = project_rows[project_rows["rack_type_norm"].isin(selected_norm)]

        unique_items = []
        seen = set()
        for _, trow in project_rows.iterrows():
            key = (
                str(trow["rack_type"]).strip(),
                trow["a"],
                trow["b"],
                str(trow["temperature_delta"]).strip(),
            )
            if key in seen:
                continue
            seen.add(key)
            unique_items.append(f"{key[0]}: a={key[1]}, b={key[2]}, delta={key[3]}")
        params_col.append(" | ".join(unique_items))

    merged["Temperature Params (Selected Installation Types)"] = params_col
    return merged


def build_unified_report(root: Path) -> tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    concentrado_df, concentrado_warnings = build_concentrado(root)
    ghi_df, ghi_warnings = build_ghi_concentrado(root)
    metrics_df, temp_df, metrics_warnings = _build_metrics_tables(root)

    merged = concentrado_df.merge(ghi_df, on="PDF Path", how="outer", suffixes=("", " (GHI)"))
    merged = merged.merge(metrics_df, on="PDF Path", how="outer", suffixes=("", " (Metrics)"))
    merged = _coalesce_project_name(merged)
    merged = _coalesce_coordinates(merged)
    merged = _build_temperature_for_installation(merged, temp_df)

    warnings = sorted(set(concentrado_warnings + ghi_warnings + metrics_warnings))
    return merged, temp_df, warnings


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Consolida extracción HelioScope en un solo archivo Excel.")
    parser.add_argument("--root", type=Path, default=Path("."), help="Carpeta raíz para buscar PDFs.")
    parser.add_argument("--out-prefix", type=str, default="helioscope_unificado", help="Prefijo del Excel de salida.")
    parser.add_argument("--json", action="store_true", help="Imprime resumen en JSON.")
    args = parser.parse_args(argv)

    root = args.root.resolve()
    merged_df, temp_df, warnings = build_unified_report(root)
    excel_path = root / f"{args.out_prefix}.xlsx"

    with pd.ExcelWriter(excel_path) as writer:
        merged_df.to_excel(writer, sheet_name="consolidado", index=False)
        temp_df.to_excel(writer, sheet_name="temperature_model_parameters", index=False)
        if warnings:
            pd.DataFrame({"warning": warnings}).to_excel(writer, sheet_name="warnings", index=False)

    if args.json:
        print(
            json.dumps(
                {
                    "excel": str(excel_path),
                    "rows": len(merged_df),
                    "temperature_rows": len(temp_df),
                    "warnings": len(warnings),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    else:
        print(f"Excel consolidado: {excel_path}")
        print(f"Filas consolidado: {len(merged_df)}")
        print(f"Filas temperatura: {len(temp_df)}")
        if warnings:
            print("Incluye hoja: warnings")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
