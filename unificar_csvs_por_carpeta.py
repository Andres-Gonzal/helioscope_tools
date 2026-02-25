#!/usr/bin/env python3
"""Unifica CSVs por subcarpeta en un solo XLSX (una hoja por carpeta)."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import List

import pandas as pd


def _sanitize_sheet_name(name: str, used: set[str]) -> str:
    invalid = set('[]:*?/\\')
    cleaned = ''.join('_' if c in invalid else c for c in name).strip() or 'Sheet'
    cleaned = cleaned[:31]
    base = cleaned
    i = 2
    while cleaned in used:
        suffix = f'_{i}'
        cleaned = (base[: 31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else base + suffix
        i += 1
    used.add(cleaned)
    return cleaned


def build_excel_from_csv_folders(root: Path, output_xlsx: Path) -> tuple[int, list[str]]:
    root = root.resolve()
    subfolders = sorted([p for p in root.iterdir() if p.is_dir()])
    warnings: List[str] = []
    sheet_count = 0
    used_sheets: set[str] = set()

    with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
        for folder in subfolders:
            csv_files = sorted(folder.glob('*.csv'))
            if not csv_files:
                continue

            frames = []
            for csv_path in csv_files:
                try:
                    frames.append(pd.read_csv(csv_path))
                except Exception as exc:
                    warnings.append(f'No se pudo leer {csv_path}: {exc}')

            if not frames:
                continue

            # Si hay varios CSV en la misma carpeta, se concatenan.
            df = pd.concat(frames, ignore_index=True)
            sheet_name = _sanitize_sheet_name(folder.name, used_sheets)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet_count += 1

    return sheet_count, warnings


def main() -> int:
    parser = argparse.ArgumentParser(description='Unifica CSVs por subcarpeta en un solo archivo Excel.')
    parser.add_argument('--root', type=Path, default=Path('.'), help='Carpeta raíz que contiene subcarpetas con CSVs.')
    parser.add_argument('--output', type=Path, default=Path('P70 NREL.xlsx'), help='Ruta del XLSX de salida.')
    args = parser.parse_args()

    sheet_count, warnings = build_excel_from_csv_folders(args.root, args.output)
    print(f'Excel generado: {args.output.resolve()}')
    print(f'Hojas creadas: {sheet_count}')
    if warnings:
        print('Advertencias:')
        for w in warnings:
            print(f'  - {w}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
