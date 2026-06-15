"""
RPA — Distribución Petroplazas | Recepción de Facturas
GUI built with customtkinter. Playwright automation runs in a background thread.
"""

import argparse
import asyncio
import csv
import os
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

# ── CLI args ─────────────────────────────────────────────
_parser = argparse.ArgumentParser(add_help=False)
_parser.add_argument("--max", type=int, default=None, metavar="N",
                     help="Limitar a los primeros N folios del Excel")
_args, _remaining = _parser.parse_known_args()
sys.argv = [sys.argv[0]] + _remaining   # drop --max so tkinter doesn't complain
MAX_RECORDS: int | None = _args.max

# ── Appearance ──────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Brand colours
C_HEADER   = "#00264d"
C_CARD_TOT = "#154360"
C_CARD_OK  = "#1e5e38"
C_CARD_ERR = "#7b241c"
C_CARD_PEN = "#7d6608"
C_BTN_RUN  = "#1a5e35"
C_BTN_STOP = "#7b0000"


# ════════════════════════════════════════════════════════
class RPAApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RPA — Recepción de Facturas  |  SIPP Petroplazas")
        self.geometry("1020x760")
        self.minsize(900, 680)
        self.resizable(True, True)

        self._log_q: queue.Queue = queue.Queue()
        self._is_running = False
        self._rpa_thread: threading.Thread | None = None
        self._start_ts: datetime | None = None
        self._total = 0
        self._processed = 0
        self._errors = 0

        self._build_ui()
        self._set_window_icon()
        self._poll_logs()
        threading.Thread(target=self._check_updates, daemon=True).start()

    # ────────────────────────────────────────────────────
    # UI construction
    # ────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)   # log area expands

        self._build_header()
        self._build_update_banner()   # row 1 — oculto hasta que haya actualizaciones
        self._build_config()
        self._build_stats()
        self._build_progress()
        self._build_logs()
        self._build_buttons()

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=(C_HEADER, C_HEADER), corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(2, weight=1)   # subtitle expands

        # ── Logo ──
        _LOGO = Path(__file__).parent / "Logo_Petroil.png"
        try:
            from PIL import Image as _PILImage
            _img = ctk.CTkImage(_PILImage.open(_LOGO), size=(40, 40))
            ctk.CTkLabel(hdr, image=_img, text="").grid(
                row=0, column=0, padx=(12, 0), pady=7)
        except Exception:
            ctk.CTkLabel(hdr, text="⛽", font=ctk.CTkFont(size=22),
                         text_color="white").grid(row=0, column=0, padx=(14, 0), pady=7)

        # ── Título ──
        ctk.CTkLabel(
            hdr,
            text="  RPA — Recepción de Facturas  |  SIPP Petroplazas",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="white",
        ).grid(row=0, column=1, padx=(6, 16), pady=11, sticky="w")

        # ── Subtítulo ──
        ctk.CTkLabel(
            hdr,
            text="Automatización de captura CC + Observaciones OC",
            font=ctk.CTkFont(size=11),
            text_color="#99bbcc",
        ).grid(row=0, column=2, padx=16, pady=11, sticky="e")

        # ── Botón de actualización manual ──
        self._btn_check_update = ctk.CTkButton(
            hdr, text="🔄", width=36, height=36,
            font=ctk.CTkFont(size=16),
            fg_color="transparent", hover_color="#1a3a5c",
            border_width=0,
            command=self._manual_check_updates,
            cursor="hand2",
        )
        self._btn_check_update.grid(row=0, column=3, padx=(0, 10), pady=6)

    def _set_window_icon(self):
        """Ícono de la ventana en taskbar/titlebar (PNG → PhotoImage)."""
        logo_path = Path(__file__).parent / "Logo_Petroil.png"
        if not logo_path.exists():
            return
        try:
            from PIL import Image as _PILImage, ImageTk as _ImageTk
            img = _PILImage.open(logo_path).resize((32, 32), _PILImage.LANCZOS)
            self._win_icon = _ImageTk.PhotoImage(img)
            self.iconphoto(True, self._win_icon)
        except Exception:
            pass

    def _build_config(self):
        cfg = ctk.CTkFrame(self)
        cfg.grid(row=2, column=0, sticky="ew", padx=12, pady=(10, 4))
        cfg.grid_columnconfigure(1, weight=1)

        # ── File picker ──
        ctk.CTkLabel(
            cfg, text="Archivo Excel:", width=105,
            font=ctk.CTkFont(weight="bold"), anchor="w",
        ).grid(row=0, column=0, padx=(12, 6), pady=(12, 4), sticky="w")

        self._file_var = ctk.StringVar()
        self._file_entry = ctk.CTkEntry(
            cfg,
            textvariable=self._file_var,
            placeholder_text="Seleccionar archivo .xlsx ...",
            state="disabled",
        )
        self._file_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=(12, 4))

        ctk.CTkButton(
            cfg, text="Examinar", width=100,
            command=self._browse_file,
        ).grid(row=0, column=2, padx=(0, 12), pady=(12, 4))

        # ── Credentials ──
        cred = ctk.CTkFrame(cfg, fg_color="transparent")
        cred.grid(row=1, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12))
        cred.grid_columnconfigure((1, 3), weight=1)

        ctk.CTkLabel(cred, text="Usuario:", width=65,
                     font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=(0, 6), sticky="w")
        self._user_var = ctk.StringVar(value="afuentes")
        ctk.CTkEntry(cred, textvariable=self._user_var, width=150).grid(
            row=0, column=1, sticky="ew", padx=(0, 22))

        ctk.CTkLabel(cred, text="Contraseña:", width=80,
                     font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=2, padx=(0, 6), sticky="w")
        self._pass_var = ctk.StringVar(value="")
        ctk.CTkEntry(cred, textvariable=self._pass_var, show="●", width=150).grid(
            row=0, column=3, sticky="ew", padx=(0, 22))

        self._headless_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            cred, text="Modo silencioso (headless)",
            variable=self._headless_var,
        ).grid(row=0, column=4, padx=(0, 4))

        # ── Workers ──
        wrk = ctk.CTkFrame(cfg, fg_color="transparent")
        wrk.grid(row=2, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12))

        ctk.CTkLabel(
            wrk, text="Sesiones paralelas:",
            font=ctk.CTkFont(weight="bold"),
        ).pack(side="left", padx=(0, 8))

        self._workers_seg = ctk.CTkSegmentedButton(
            wrk, values=["1×", "2×", "4×", "8×"], width=230,
        )
        self._workers_seg.set("1×")
        self._workers_seg.pack(side="left")

        ctk.CTkLabel(
            wrk,
            text="  ·  más sesiones = más velocidad (cada una inicia login en SIPP)",
            font=ctk.CTkFont(size=10), text_color="gray55",
        ).pack(side="left", padx=(8, 0))

    def _build_stats(self):
        stats = ctk.CTkFrame(self, fg_color="transparent")
        stats.grid(row=3, column=0, sticky="ew", padx=12, pady=4)
        for c in range(4):
            stats.grid_columnconfigure(c, weight=1)

        cards = [
            ("TOTAL REGISTROS", C_CARD_TOT),
            ("PROCESADOS OK",   C_CARD_OK),
            ("ERRORES",         C_CARD_ERR),
            ("PENDIENTES",      C_CARD_PEN),
        ]
        self._stat_vals: list[ctk.CTkLabel] = []
        for i, (label, color) in enumerate(cards):
            f = ctk.CTkFrame(stats, fg_color=(color, color), corner_radius=8)
            f.grid(row=0, column=i, sticky="ew", padx=5, pady=6)
            ctk.CTkLabel(
                f, text=label,
                font=ctk.CTkFont(size=9, weight="bold"),
                text_color="#ccdde8",
            ).pack(pady=(10, 0))
            lbl = ctk.CTkLabel(
                f, text="0",
                font=ctk.CTkFont(size=28, weight="bold"),
                text_color="white",
            )
            lbl.pack(pady=(0, 10))
            self._stat_vals.append(lbl)

    def _build_progress(self):
        prog = ctk.CTkFrame(self, fg_color="transparent")
        prog.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 4))
        prog.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(prog, fg_color="transparent")
        top.grid(row=0, column=0, columnspan=2, sticky="ew")
        top.grid_columnconfigure(0, weight=1)

        self._status_lbl = ctk.CTkLabel(
            top, text="En espera — selecciona un archivo.",
            text_color="gray", font=ctk.CTkFont(size=11),
        )
        self._status_lbl.grid(row=0, column=0, sticky="w")

        self._elapsed_lbl = ctk.CTkLabel(
            top, text="", font=ctk.CTkFont(size=11), text_color="gray"
        )
        self._elapsed_lbl.grid(row=0, column=1, sticky="e")

        bar_row = ctk.CTkFrame(prog, fg_color="transparent")
        bar_row.grid(row=1, column=0, sticky="ew")
        bar_row.grid_columnconfigure(0, weight=1)

        self._progress = ctk.CTkProgressBar(bar_row, height=14)
        self._progress.grid(row=0, column=0, sticky="ew")
        self._progress.set(0)

        self._pct_lbl = ctk.CTkLabel(
            bar_row, text="0%", width=42,
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        self._pct_lbl.grid(row=0, column=1, padx=(8, 0))

    def _build_logs(self):
        wrap = ctk.CTkFrame(self)
        wrap.grid(row=5, column=0, sticky="nsew", padx=12, pady=4)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(1, weight=1)

        hdr = ctk.CTkFrame(wrap, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 0))
        hdr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            hdr, text="Registro de actividad",
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            hdr, text="Limpiar", width=72, height=26,
            fg_color="gray35", hover_color="gray25",
            command=self._clear_logs,
        ).grid(row=0, column=1)

        self._log_box = ctk.CTkTextbox(
            wrap, state="disabled",
            font=ctk.CTkFont(family="Courier New", size=11),
        )
        self._log_box.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)

    def _build_buttons(self):
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.grid(row=6, column=0, sticky="ew", padx=12, pady=(4, 14))

        self._btn_run = ctk.CTkButton(
            row, text="▶  EJECUTAR RPA", width=190, height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=(C_BTN_RUN, C_BTN_RUN),
            hover_color=("#0f3d22", "#0f3d22"),
            command=self._start,
        )
        self._btn_run.pack(side="left", padx=(0, 10))

        self._btn_cancel = ctk.CTkButton(
            row, text="⏹  CANCELAR", width=160, height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=(C_BTN_STOP, C_BTN_STOP),
            hover_color=("#550000", "#550000"),
            command=self._cancel,
            state="disabled",
        )
        self._btn_cancel.pack(side="left")

        # ── Separator ──
        ctk.CTkLabel(row, text="│", text_color="gray40",
                     font=ctk.CTkFont(size=28)).pack(side="left", padx=16)

        # ── Compare button ──
        self._btn_compare = ctk.CTkButton(
            row, text="📊  COMPARAR CSV", width=200, height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#1a4a6e", "#1a4a6e"),
            hover_color=("#0d2e47", "#0d2e47"),
            command=self._start_compare,
        )
        self._btn_compare.pack(side="left")

        ctk.CTkLabel(row, text="│", text_color="gray40",
                     font=ctk.CTkFont(size=28)).pack(side="left", padx=16)

        # ── Catalogs button ──
        self._btn_catalogs = ctk.CTkButton(
            row, text="📋  CATÁLOGOS", width=180, height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#3b2a6b", "#3b2a6b"),
            hover_color=("#251a46", "#251a46"),
            command=self._open_catalogs,
        )
        self._btn_catalogs.pack(side="left")

    # ────────────────────────────────────────────────────
    # Auto-update
    # ────────────────────────────────────────────────────
    def _build_update_banner(self):
        """Banner verde entre header y config. Oculto hasta que haya commits nuevos."""
        self._update_frame = ctk.CTkFrame(
            self, fg_color=("#1b5e33", "#1b5e33"), corner_radius=0,
        )
        # No se agrega al grid todavía — se muestra solo si hay actualizaciones.

        inner = ctk.CTkFrame(self._update_frame, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=7)
        inner.grid_columnconfigure(0, weight=1)

        self._update_lbl = ctk.CTkLabel(
            inner, text="",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="white",
        )
        self._update_lbl.grid(row=0, column=0, sticky="w")

        self._btn_update = ctk.CTkButton(
            inner, text="⬇  Actualizar y reiniciar", width=200, height=28,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="white", text_color="#1b5e33",
            hover_color="#d4f0d4",
            command=self._do_update,
        )
        self._btn_update.grid(row=0, column=1, padx=(14, 0))

        ctk.CTkButton(
            inner, text="✕", width=28, height=28,
            fg_color="transparent", text_color="#aaddaa",
            hover_color="#0f3d22",
            command=self._dismiss_update_banner,
        ).grid(row=0, column=2, padx=(6, 0))

    def _check_updates(self):
        """Hilo de fondo: fetch → compara HEAD local vs remoto."""
        root = str(Path(__file__).parent)
        try:
            fetch = subprocess.run(
                ["git", "fetch", "origin", "main", "--quiet"],
                cwd=root, capture_output=True, timeout=15,
            )
            if fetch.returncode != 0:
                return   # sin internet, sin git, sin remote — ignorar silenciosamente

            local = subprocess.run(
                ["git", "rev-parse", "HEAD"],
                cwd=root, capture_output=True,
            ).stdout.decode().strip()

            remote = subprocess.run(
                ["git", "rev-parse", "origin/main"],
                cwd=root, capture_output=True,
            ).stdout.decode().strip()

            if local == remote:
                return   # ya está al día — no mostrar nada

            behind = subprocess.run(
                ["git", "rev-list", "--count", "HEAD..origin/main"],
                cwd=root, capture_output=True,
            ).stdout.decode().strip()

            last_msg = subprocess.run(
                ["git", "log", "-1", "--pretty=%s", "origin/main"],
                cwd=root, capture_output=True,
            ).stdout.decode().strip()

            self.after(0, lambda: self._show_update_banner(behind, last_msg))

        except Exception:
            pass   # el check de actualizaciones nunca debe interrumpir la app

    def _show_update_banner(self, behind: str, last_msg: str):
        n = behind or "?"
        self._update_lbl.configure(
            text=f"  🔄  {n} actualización(es) disponible(s)  ·  \"{last_msg}\""
        )
        self._update_frame.grid(row=1, column=0, sticky="ew")
        self._log(f"Actualización disponible ({n} commit(s)): {last_msg}", "warn")

    def _dismiss_update_banner(self):
        self._update_frame.grid_remove()

    def _do_update(self):
        self._btn_update.configure(state="disabled", text="Actualizando…")

        def _pull():
            root = str(Path(__file__).parent)
            try:
                res = subprocess.run(
                    ["git", "pull", "origin", "main"],
                    cwd=root, capture_output=True, text=True, timeout=60,
                )
                if res.returncode == 0:
                    self.after(0, self._restart_app)
                else:
                    err = res.stderr.strip() or res.stdout.strip()
                    self.after(0, lambda: messagebox.showerror(
                        "Error al actualizar",
                        f"git pull falló:\n\n{err}",
                        parent=self,
                    ))
                    self.after(0, lambda: self._btn_update.configure(
                        state="normal", text="⬇  Actualizar y reiniciar",
                    ))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror(
                    "Error", str(exc), parent=self,
                ))

        threading.Thread(target=_pull, daemon=True).start()

    def _restart_app(self):
        self.destroy()
        if sys.platform == "win32":
            # os.execv no reemplaza el proceso en Windows — usar Popen + exit
            subprocess.Popen([sys.executable] + sys.argv)
            sys.exit(0)
        else:
            os.execv(sys.executable, [sys.executable] + sys.argv)

    def _manual_check_updates(self):
        """Triggered by the 🔄 button in the header — runs the same check in background."""
        self._btn_check_update.configure(state="disabled", text="⏳")
        self._log("Comprobando actualizaciones…", "info")

        def _check_and_restore():
            self._check_updates()
            self.after(0, lambda: self._btn_check_update.configure(
                state="normal", text="🔄",
            ))

        threading.Thread(target=_check_and_restore, daemon=True).start()

    # ────────────────────────────────────────────────────
    # File handling
    # ────────────────────────────────────────────────────
    def _browse_file(self):
        initial = str(Path(__file__).parent / "Recepcion_Facturas")
        path = filedialog.askopenfilename(
            title="Seleccionar Excel de Recepción de Facturas",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
            initialdir=initial,
        )
        if not path:
            return
        # Enable the entry widget to update its value
        self._file_entry.configure(state="normal")
        self._file_var.set(path)
        self._file_entry.configure(state="disabled")
        self._load_file_info(path)

    def _load_file_info(self, path: str):
        try:
            from rpa.excel_handler import ExcelHandler
            handler = ExcelHandler(path)
            folios = handler.get_folios()
            n = len(folios)
            self._total = n
            self._stat_vals[0].configure(text=str(n))
            self._stat_vals[3].configure(text=str(n))
            self._status_lbl.configure(
                text=f"Listo — {n} folio(s) en '{Path(path).name}'"
            )
            self._log(f"Archivo cargado: {Path(path).name}  ({n} folios)", "info")
        except Exception as exc:
            self._log(f"Error leyendo Excel: {exc}", "error")

    # ────────────────────────────────────────────────────
    # Logging
    # ────────────────────────────────────────────────────
    def _log(self, msg: str, level: str = "info"):
        self._log_q.put((level, msg))

    def _poll_logs(self):
        try:
            while True:
                level, msg = self._log_q.get_nowait()
                icon = {"ok": "✓", "error": "✗", "warn": "⚠", "info": "·"}.get(
                    level, "·"
                )
                ts = datetime.now().strftime("%H:%M:%S")
                line = f"[{ts}] {icon}  {msg}\n"
                self._log_box.configure(state="normal")
                self._log_box.insert("end", line)
                self._log_box.see("end")
                self._log_box.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_logs)

    def _clear_logs(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    # ────────────────────────────────────────────────────
    # Progress helpers
    # ────────────────────────────────────────────────────
    def _on_progress(self, processed: int, errors: int, folio: str):
        """Called from background thread — use after() for thread safety."""
        self._processed = processed
        self._errors = errors
        self.after(0, self._refresh_stats, processed, errors, folio)

    def _refresh_stats(self, processed: int, errors: int, folio: str = ""):
        pending = max(0, self._total - processed - errors)
        self._stat_vals[1].configure(text=str(processed))
        self._stat_vals[2].configure(text=str(errors))
        self._stat_vals[3].configure(text=str(pending))
        if self._total > 0:
            pct = (processed + errors) / self._total
            self._progress.set(pct)
            self._pct_lbl.configure(text=f"{int(pct * 100)}%")
        if folio:
            self._status_lbl.configure(text=f"Procesando: {folio}")

    def _tick(self):
        if self._is_running and self._start_ts:
            secs = int((datetime.now() - self._start_ts).total_seconds())
            m, s = divmod(secs, 60)
            self._elapsed_lbl.configure(text=f"Tiempo: {m:02d}:{s:02d}")
            self.after(1000, self._tick)

    # ────────────────────────────────────────────────────
    # RPA lifecycle
    # ────────────────────────────────────────────────────
    def _start(self):
        if not self._file_var.get():
            messagebox.showwarning("Sin archivo", "Por favor selecciona un archivo Excel.")
            return
        if not self._user_var.get() or not self._pass_var.get():
            messagebox.showwarning("Sin credenciales", "Ingresa usuario y contraseña.")
            return

        self._is_running = True
        self._processed = 0
        self._errors = 0
        self._start_ts = datetime.now()

        # Reset UI
        for lbl in self._stat_vals[1:]:
            lbl.configure(text="0")
        self._progress.set(0)
        self._pct_lbl.configure(text="0%")
        self._btn_run.configure(state="disabled")
        self._btn_cancel.configure(state="normal")
        self._status_lbl.configure(text="Iniciando proceso...")
        self._log("─" * 62, "info")
        self._log(
            f"Inicio: {self._start_ts.strftime('%d/%m/%Y %H:%M:%S')}  |  "
            f"Archivo: {Path(self._file_var.get()).name}",
            "info",
        )

        self._tick()
        self._rpa_thread = threading.Thread(target=self._thread_run, daemon=True)
        self._rpa_thread.start()

    def _cancel(self):
        if self._is_running:
            self._is_running = False
            self._log("Cancelación solicitada — terminando folio actual...", "warn")
            self._status_lbl.configure(text="Cancelando...")

    def _thread_run(self):
        # Playwright requiere ProactorEventLoop en Windows (Python ≤ 3.9)
        if sys.platform == "win32":
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        asyncio.run(self._async_run())

    async def _async_run(self):
        from rpa.automation import RPAAutomation
        from rpa.excel_handler import ExcelHandler

        try:
            handler = ExcelHandler(self._file_var.get())
            folio_rows = handler.get_folios()
            if MAX_RECORDS is not None:
                folio_rows = folio_rows[:MAX_RECORDS]
                self._log(f"[--max {MAX_RECORDS}] Limitando a los primeros {len(folio_rows)} folio(s).", "warn")
            self._total = len(folio_rows)
            self.after(0, lambda: self._stat_vals[0].configure(text=str(self._total)))

            # ── Número de workers ──
            n_workers = int(self._workers_seg.get()[0])   # "4×" → 4
            n_workers = min(n_workers, self._total) or 1  # no más workers que folios

            self._log(
                f"Iniciando RPA — {self._total} folio(s) en {n_workers} sesión(es) paralela(s)…",
                "info",
            )

            # ── Dividir folios en chunks contiguos ──
            chunk_sz = (self._total + n_workers - 1) // n_workers
            chunks   = [
                folio_rows[i * chunk_sz:(i + 1) * chunk_sz]
                for i in range(n_workers)
                if folio_rows[i * chunk_sz:(i + 1) * chunk_sz]
            ]
            actual_workers = len(chunks)

            # ── Contadores de progreso por worker ──
            proc_counts = [0] * actual_workers
            err_counts  = [0] * actual_workers

            def make_progress_cb(idx: int):
                def _cb(processed: int, errors: int, folio: str):
                    proc_counts[idx] = processed
                    err_counts[idx]  = errors
                    label = f"W{idx+1}›{folio}" if actual_workers > 1 else folio
                    self._on_progress(sum(proc_counts), sum(err_counts), label)
                return _cb

            # ── Lanzar workers en paralelo ──
            async def run_worker(idx: int, chunk):
                prefix = f"[W{idx+1}] " if actual_workers > 1 else ""

                def _log(msg: str, level: str = "info"):
                    self._log(f"{prefix}{msg}", level)

                rpa = RPAAutomation(
                    username=self._user_var.get(),
                    password=self._pass_var.get(),
                    headless=self._headless_var.get(),
                    log_fn=_log,
                    cancel_fn=lambda: not self._is_running,
                )
                worker_results = await rpa.run(chunk, on_progress=make_progress_cb(idx))
                return worker_results, rpa

            all_returns = await asyncio.gather(
                *[run_worker(i, chunk) for i, chunk in enumerate(chunks)],
                return_exceptions=True,
            )

            # ── Fusionar resultados ──
            results:    list = []
            all_skipped: list = []
            all_not_found: list = []

            for i, ret in enumerate(all_returns):
                if isinstance(ret, BaseException):
                    import traceback as _tb
                    self._log(f"[W{i+1}] Error fatal: {ret}", "error")
                    self._log(_tb.format_exception(type(ret), ret, ret.__traceback__)[-1], "error")
                else:
                    w_results, rpa = ret
                    results.extend(w_results)
                    all_skipped.extend(getattr(rpa, "skipped", []))
                    all_not_found.extend(getattr(rpa, "not_found", []))

            results.sort(key=lambda r: r[0])  # ordenar por row_num

            # ── Guardar en Excel ──
            if results:
                self._log("Guardando resultados en Excel…", "info")
                handler.ensure_headers()
                for row_num, cc, obs, subtotal, descuento, iva, gastos_envio, total_oc in results:
                    handler.write_result(row_num, cc, obs, subtotal, descuento, iva, gastos_envio, total_oc)
                handler.save()
                self._log(f"Excel actualizado: {Path(self._file_var.get()).name}", "ok")

            # ── Resumen final ──
            ok_count = sum(1 for r in results if r[1])   # r[1] = cc
            elapsed  = int((datetime.now() - self._start_ts).total_seconds())
            m, s     = divmod(elapsed, 60)

            self._log("═" * 62, "info")
            self._log("RESUMEN FINAL", "info")
            self._log(f"  Total folios:        {self._total}", "info")
            self._log(f"  Sesiones usadas:     {actual_workers}", "info")
            self._log(f"  Con CC extraído:     {ok_count}", "ok")
            self._log(f"  Sin OC / con error:  {self._total - ok_count}", "warn")
            if all_skipped:
                self._log(f"  Duplicados omitidos: {len(all_skipped)}", "warn")
                for f in all_skipped:
                    self._log(f"    · {f}", "warn")
            if all_not_found:
                self._log(f"  No encontrados:      {len(all_not_found)}", "warn")
                for f in all_not_found:
                    self._log(f"    · {f}", "warn")
            self._log(f"  Tiempo total:        {m}m {s}s", "info")
            self._log("═" * 62, "info")

            self.after(
                0,
                lambda: self._status_lbl.configure(
                    text=f"Completado — {ok_count}/{self._total} con CC extraído  ({m}m {s}s)"
                ),
            )

        except Exception as exc:
            import traceback
            self._log(f"Error crítico: {exc}", "error")
            self._log(traceback.format_exc(), "error")

        finally:
            self._is_running = False
            self.after(
                0,
                lambda: [
                    self._btn_run.configure(state="normal"),
                    self._btn_cancel.configure(state="disabled"),
                ],
            )


    # ────────────────────────────────────────────────────
    # Catálogos de distribución
    # ────────────────────────────────────────────────────
    def _open_catalogs(self):
        win = CatalogEditorWindow(self, Path(__file__).parent / "Distribucion")
        win.focus()

    # ────────────────────────────────────────────────────
    # Comparar sucursales
    # ────────────────────────────────────────────────────
    def _start_compare(self):
        initial = str(Path(__file__).parent / "Recepcion_Facturas")
        path = filedialog.askopenfilename(
            title="Seleccionar CSV para comparar sucursales",
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
            initialdir=initial,
        )
        if not path:
            return

        self._log("─" * 62, "info")
        self._log(f"Comparando sucursales: {Path(path).name}", "info")
        self._status_lbl.configure(text=f"Comparando: {Path(path).name}…")
        self._btn_compare.configure(state="disabled")

        t = threading.Thread(target=self._compare_thread, args=(path,), daemon=True)
        t.start()

    def _compare_thread(self, csv_path: str):
        try:
            import compare_sucursales as cs
            from openpyxl import Workbook

            self._log("  Cargando registros…", "info")
            _, data = cs.load_csv(csv_path)
            total = len(data)
            self._log(f"  {total} registros cargados.", "info")

            catalog = cs.load_catalogs(cs.DISTRIBUCION_DIR)
            if catalog:
                self._log(f"  {len(catalog)} claves de catálogo de distribución cargadas.", "info")
            else:
                self._log("  [AVISO] Carpeta Distribucion/ no encontrada — sin cálculo de montos.", "warn")

            wb           = Workbook()
            ws_main      = wb.active
            ws_sum       = wb.create_sheet()
            ws_suc       = wb.create_sheet()
            ws_dist      = wb.create_sheet()
            ws_dist_calc = wb.create_sheet()

            counts, details = cs.build_main_sheet(ws_main, data)
            cs.build_summary_sheet(ws_sum, counts, total)
            cs.build_sucursal_detail_sheet(ws_suc, details)
            cs.build_distribucion_sheet(ws_dist, details)
            cs.build_distribucion_calculada_sheet(ws_dist_calc, details, catalog)

            out = Path(csv_path).parent / (Path(csv_path).stem + "_comparacion.xlsx")
            wb.save(str(out))

            match_pct = counts["MATCH"] / total * 100 if total else 0
            mis_pct   = counts["MISMATCH"] / total * 100 if total else 0
            dis_pct   = counts["DISTRIBUCIÓN"] / total * 100 if total else 0
            sin_pct   = counts["SIN SUCURSAL"] / total * 100 if total else 0

            self._log("═" * 62, "info")
            self._log("RESUMEN COMPARACIÓN DE SUCURSALES", "info")
            self._log(f"  Total registros  : {total}", "info")
            self._log(f"  MATCH ✓          : {counts['MATCH']}  ({match_pct:.1f}%)", "ok")
            self._log(f"  MISMATCH ✗       : {counts['MISMATCH']}  ({mis_pct:.1f}%)", "error")
            self._log(f"  DISTRIBUCIÓN     : {counts['DISTRIBUCIÓN']}  ({dis_pct:.1f}%)", "warn")
            self._log(f"  Sin sucursal     : {counts['SIN SUCURSAL']}  ({sin_pct:.1f}%)", "info")
            self._log(f"  Archivo generado : {out.name}", "ok")
            self._log("═" * 62, "info")

            self.after(0, lambda: self._status_lbl.configure(
                text=f"Comparación lista — {counts['MATCH']} MATCH  |  "
                     f"{counts['MISMATCH']} MISMATCH  |  "
                     f"{counts['DISTRIBUCIÓN']} DISTRIBUCIÓN"
            ))

        except Exception as exc:
            import traceback
            self._log(f"Error en comparación: {exc}", "error")
            self._log(traceback.format_exc(), "error")
            self.after(0, lambda: self._status_lbl.configure(text="Error en comparación."))

        finally:
            self.after(0, lambda: self._btn_compare.configure(state="normal"))


# ════════════════════════════════════════════════════════
# Catálogos de distribución — ventana de edición
# ════════════════════════════════════════════════════════

_CATALOG_COLS = ["(GCC)", "ZONA(CC)", "ESTACION", "PORCENTAJE"]
_CATALOG_FILES = {
    "Mazatlán General": "Mazatlan_General.csv",
    "Corporativo":      "Corporativo.csv",
    "Zonas":            "Zonas.csv",
}


class CatalogEditorWindow(ctk.CTkToplevel):
    """Ventana modal para ver, editar, agregar y eliminar entradas de los
    catálogos de distribución (CSV en la carpeta Distribucion/)."""

    def __init__(self, master, distribucion_dir: Path):
        super().__init__(master)
        self.title("Catálogos de Distribución")
        self.geometry("860x580")
        self.minsize(700, 460)
        self._dir = distribucion_dir
        self._trees: dict = {}   # tab_name → (Treeview, cols)
        self._apply_tree_style()
        self._build_ui()

    # ── ttk dark style ───────────────────────────────────────────────────────
    def _apply_tree_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Cat.Treeview",
            background="#2b2b2b", foreground="#e0e0e0",
            fieldbackground="#2b2b2b", borderwidth=0, rowheight=24,
            font=("Courier New", 10))
        style.configure("Cat.Treeview.Heading",
            background="#00264d", foreground="white",
            relief="flat", font=("Helvetica", 10, "bold"))
        style.map("Cat.Treeview",
            background=[("selected", "#1a5e8a")],
            foreground=[("selected", "white")])
        style.configure("Cat.Scrollbar",
            background="#3a3a3a", troughcolor="#1e1e1e",
            arrowcolor="#aaaaaa", borderwidth=0)

    # ── UI principal ─────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        hdr = ctk.CTkFrame(self, fg_color=(C_HEADER, C_HEADER), corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            hdr, text="  Catálogos de Distribución",
            font=ctk.CTkFont(size=14, weight="bold"), text_color="white",
        ).pack(side="left", padx=12, pady=9)
        ctk.CTkLabel(
            hdr, text="Doble clic en una fila para editar",
            font=ctk.CTkFont(size=10), text_color="#99bbcc",
        ).pack(side="right", padx=14, pady=9)

        # Tabs
        tabs = ctk.CTkTabview(self, anchor="nw")
        tabs.grid(row=1, column=0, sticky="nsew", padx=10, pady=(8, 4))
        for tab_name in _CATALOG_FILES:
            tabs.add(tab_name)
            self._build_tab(tabs.tab(tab_name), tab_name)

        # Bottom bar
        bot = ctk.CTkFrame(self, fg_color="transparent")
        bot.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        ctk.CTkButton(
            bot, text="✕  Cerrar", width=110, height=38,
            fg_color="gray30", hover_color="gray20",
            command=self.destroy,
        ).pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            bot, text="💾  Guardar todos", width=170, height=38,
            fg_color=("#1a5e35", "#1a5e35"), hover_color=("#0f3d22", "#0f3d22"),
            command=self._save_all,
        ).pack(side="right")

    # ── Tab con tabla ─────────────────────────────────────────────────────────
    def _build_tab(self, parent, tab_name):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        fname = _CATALOG_FILES[tab_name]
        rows  = self._load_csv(fname)

        # Frame con treeview + scrollbar
        frm = ctk.CTkFrame(parent, corner_radius=6)
        frm.grid(row=0, column=0, sticky="nsew")
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(0, weight=1)

        tree = ttk.Treeview(
            frm, columns=_CATALOG_COLS, show="headings",
            style="Cat.Treeview", selectmode="browse",
        )
        col_widths = [160, 130, 180, 90]
        for col, w in zip(_CATALOG_COLS, col_widths):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor="w", stretch=True)

        vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
        vsb.grid(row=0, column=1, sticky="ns", pady=4, padx=(0, 4))

        for row in rows:
            tree.insert("", "end", values=[row.get(c, "") for c in _CATALOG_COLS])

        tree.bind("<Double-1>", lambda e, tn=tab_name: self._edit_row(tn))

        self._trees[tab_name] = tree

        # Botones de fila
        act = ctk.CTkFrame(parent, fg_color="transparent")
        act.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ctk.CTkButton(
            act, text="＋  Agregar fila", width=140, height=30,
            command=lambda tn=tab_name: self._add_row(tn),
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            act, text="✎  Editar", width=100, height=30,
            fg_color="gray35", hover_color="gray25",
            command=lambda tn=tab_name: self._edit_row(tn),
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            act, text="✕  Eliminar", width=110, height=30,
            fg_color="#7b241c", hover_color="#5a1a14",
            command=lambda tn=tab_name: self._del_row(tn),
        ).pack(side="left")

    # ── Row actions ───────────────────────────────────────────────────────────
    def _add_row(self, tab_name: str):
        self._show_row_dialog(tab_name, None)

    def _edit_row(self, tab_name: str):
        tree = self._trees[tab_name]
        sel  = tree.selection()
        if not sel:
            messagebox.showinfo("Seleccionar fila",
                                "Haz clic en una fila primero.", parent=self)
            return
        self._show_row_dialog(tab_name, sel[0])

    def _del_row(self, tab_name: str):
        tree = self._trees[tab_name]
        sel  = tree.selection()
        if not sel:
            return
        if messagebox.askyesno("Eliminar fila",
                               "¿Eliminar la fila seleccionada?", parent=self):
            tree.delete(sel[0])

    def _show_row_dialog(self, tab_name: str, item_id):
        tree   = self._trees[tab_name]
        is_new = item_id is None
        cur    = list(tree.item(item_id, "values")) if not is_new else [""] * len(_CATALOG_COLS)

        dlg = ctk.CTkToplevel(self)
        dlg.title("Nueva fila" if is_new else "Editar fila")
        dlg.geometry("420x230")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.grid_columnconfigure(1, weight=1)

        entries: list[ctk.CTkEntry] = []
        for i, (col, val) in enumerate(zip(_CATALOG_COLS, cur)):
            ctk.CTkLabel(
                dlg, text=col + ":", width=110, anchor="w",
                font=ctk.CTkFont(weight="bold"),
            ).grid(row=i, column=0, padx=(18, 8), pady=7, sticky="w")
            ent = ctk.CTkEntry(dlg, width=240)
            ent.insert(0, str(val))
            ent.grid(row=i, column=1, padx=(0, 18), pady=7, sticky="ew")
            entries.append(ent)

        def _ok():
            vals = tuple(e.get().strip() for e in entries)
            if is_new:
                tree.insert("", "end", values=vals)
            else:
                tree.item(item_id, values=vals)
            dlg.destroy()

        btns = ctk.CTkFrame(dlg, fg_color="transparent")
        btns.grid(row=len(_CATALOG_COLS), column=0, columnspan=2, pady=(6, 14))
        ctk.CTkButton(btns, text="Aceptar", width=110, command=_ok).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Cancelar", width=110,
                      fg_color="gray30", hover_color="gray20",
                      command=dlg.destroy).pack(side="left", padx=6)

        entries[0].focus()
        dlg.bind("<Return>", lambda e: _ok())
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    # ── CSV helpers ───────────────────────────────────────────────────────────
    def _load_csv(self, fname: str) -> list[dict]:
        path = self._dir / fname
        if not path.exists():
            return []
        with open(path, encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))

    def _save_all(self):
        try:
            self._dir.mkdir(parents=True, exist_ok=True)
            for tab_name, fname in _CATALOG_FILES.items():
                tree = self._trees[tab_name]
                rows = [
                    {c: v for c, v in zip(_CATALOG_COLS, tree.item(iid, "values"))}
                    for iid in tree.get_children()
                ]
                with open(self._dir / fname, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.DictWriter(f, fieldnames=_CATALOG_COLS)
                    writer.writeheader()
                    writer.writerows(rows)
            messagebox.showinfo("Guardado", "Catálogos guardados correctamente.", parent=self)
        except Exception as exc:
            messagebox.showerror("Error al guardar", str(exc), parent=self)


# ════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = RPAApp()
    app.mainloop()
