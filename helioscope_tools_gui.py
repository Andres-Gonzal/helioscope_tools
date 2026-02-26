#!/usr/bin/env python3
"""GUI para herramientas HelioScope.

Incluye:
- Generar reporte unificado desde PDFs.
- Unificar CSVs por subcarpeta en un solo XLSX.
"""
from __future__ import annotations

import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


class HelioscopeToolsGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("HelioScope Tools")
        self.root.geometry("900x620")

        self.base_dir = Path(__file__).resolve().parent
        self.python_exe = sys.executable
        self.queue: queue.Queue[tuple[str, str]] = queue.Queue()

        self._build_ui()
        self._poll_queue()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        path_frame = ttk.LabelFrame(main, text="Carpeta de trabajo", padding=10)
        path_frame.pack(fill="x")

        self.root_var = tk.StringVar(value=str(Path.cwd()))
        ttk.Label(path_frame, text="Ruta raíz:").grid(row=0, column=0, sticky="w")
        ttk.Entry(path_frame, textvariable=self.root_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(path_frame, text="Seleccionar", command=self._select_root).grid(row=1, column=1, sticky="ew")
        path_frame.columnconfigure(0, weight=1)

        pdf_frame = ttk.LabelFrame(main, text="1) Reporte unificado (PDF -> XLSX)", padding=10)
        pdf_frame.pack(fill="x", pady=(10, 0))

        self.unified_prefix_var = tk.StringVar(value="helioscope_unificado")
        self.sort_options = {
            "Project Name (A-Z)": "project_name_asc",
            "Project Name (Z-A)": "project_name_desc",
            "Ruta PDF (A-Z)": "pdf_path_asc",
            "Sin ordenar": "none",
        }
        self.sort_label_var = tk.StringVar(value="Project Name (A-Z)")
        ttk.Label(pdf_frame, text="Nombre de salida (sin .xlsx):").grid(row=0, column=0, sticky="w")
        ttk.Entry(pdf_frame, textvariable=self.unified_prefix_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.btn_run_unified = ttk.Button(pdf_frame, text="Generar reporte unificado", command=self._run_unified)
        self.btn_run_unified.grid(row=1, column=1, sticky="ew")
        ttk.Label(pdf_frame, text="Orden en consolidado:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.sort_combo = ttk.Combobox(
            pdf_frame,
            textvariable=self.sort_label_var,
            values=list(self.sort_options.keys()),
            state="readonly",
        )
        self.sort_combo.grid(row=3, column=0, sticky="ew", padx=(0, 8))
        pdf_frame.columnconfigure(0, weight=1)

        csv_frame = ttk.LabelFrame(main, text="2) Unificar CSVs por carpeta", padding=10)
        csv_frame.pack(fill="x", pady=(10, 0))

        self.csv_output_var = tk.StringVar(value="P70 NREL.xlsx")
        ttk.Label(csv_frame, text="Archivo de salida (.xlsx):").grid(row=0, column=0, sticky="w")
        ttk.Entry(csv_frame, textvariable=self.csv_output_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.btn_run_csv = ttk.Button(csv_frame, text="Unificar CSVs", command=self._run_csv)
        self.btn_run_csv.grid(row=1, column=1, sticky="ew")
        csv_frame.columnconfigure(0, weight=1)

        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=(10, 0))
        self.btn_clear = ttk.Button(actions, text="Limpiar log", command=self._clear_log)
        self.btn_clear.pack(side="left")

        log_frame = ttk.LabelFrame(main, text="Log", padding=10)
        log_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.log = tk.Text(log_frame, wrap="word", height=20)
        self.log.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scroll.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=scroll.set)

    def _select_root(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.root_var.get() or str(Path.cwd()))
        if selected:
            self.root_var.set(selected)

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def _clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.btn_run_unified.configure(state=state)
        self.btn_run_csv.configure(state=state)

    def _validate_root(self) -> Path | None:
        root_path = Path(self.root_var.get()).expanduser().resolve()
        if not root_path.exists() or not root_path.is_dir():
            messagebox.showerror("Ruta inválida", f"La carpeta no existe:\n{root_path}")
            return None
        return root_path

    def _run_unified(self) -> None:
        root_path = self._validate_root()
        if root_path is None:
            return

        prefix = self.unified_prefix_var.get().strip()
        if not prefix:
            messagebox.showerror("Dato faltante", "Escribe un nombre de salida para el reporte unificado.")
            return

        script = self.base_dir / "generar_reporte_unificado.py"
        sort_by = self.sort_options.get(self.sort_label_var.get(), "project_name_asc")
        cmd = [
            self.python_exe,
            str(script),
            "--root",
            str(root_path),
            "--out-prefix",
            prefix,
            "--sort-by",
            sort_by,
        ]
        self._start_command(cmd, root_path)

    def _run_csv(self) -> None:
        root_path = self._validate_root()
        if root_path is None:
            return

        output_name = self.csv_output_var.get().strip()
        if not output_name:
            messagebox.showerror("Dato faltante", "Escribe el nombre del archivo XLSX de salida para CSVs.")
            return

        output_path = (root_path / output_name).resolve()
        script = self.base_dir / "unificar_csvs_por_carpeta.py"
        cmd = [self.python_exe, str(script), "--root", str(root_path), "--output", str(output_path)]
        self._start_command(cmd, root_path)

    def _start_command(self, cmd: list[str], cwd: Path) -> None:
        self._set_running(True)
        self._append_log("$ " + " ".join(cmd))

        def worker() -> None:
            try:
                process = subprocess.Popen(
                    cmd,
                    cwd=str(cwd),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
                assert process.stdout is not None
                for line in process.stdout:
                    self.queue.put(("log", line.rstrip("\n")))
                code = process.wait()
                if code == 0:
                    self.queue.put(("done", "Proceso terminado correctamente."))
                else:
                    self.queue.put(("error", f"Proceso terminó con código {code}."))
            except Exception as exc:
                self.queue.put(("error", f"Error al ejecutar proceso: {exc}"))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, msg = self.queue.get_nowait()
                if kind == "log":
                    self._append_log(msg)
                elif kind == "done":
                    self._append_log(msg)
                    self._set_running(False)
                elif kind == "error":
                    self._append_log(msg)
                    self._set_running(False)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)


def main() -> int:
    root = tk.Tk()
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
    app = HelioscopeToolsGUI(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
