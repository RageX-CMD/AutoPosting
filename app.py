#!/usr/bin/env python3
"""Windows-Oberfläche für den Event-Post-Generator. Kein Python, keine Konsole."""

from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

from forum_poster import (
    BOARD_NAME,
    CATEGORY_NAME,
    ForumPoster,
    load_posting_state,
)
from generate_event_docx import (
    app_base_dir,
    format_week_label,
    generate_documents,
    generate_missing_weeks,
    list_excel_files,
    load_weeks,
    week_already_generated,
)


APP_TITLE = f"Event-Post Generator – {BOARD_NAME}"
ACCENT = "#d35400"
ACCENT_HOVER = "#e67e22"
BG = "#f4f1ec"
CARD = "#ffffff"
TEXT = "#2c3e50"
MUTED = "#7f8c8d"


def enable_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            from ctypes import windll

            windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def open_path(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606 - lokaler Explorer/Word-Start
        return
    os.system(f'xdg-open "{path}"')  # noqa: S605


class EventPostApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(bg=BG)
        self.root.minsize(820, 650)
        self.root.geometry("940x740")

        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(app_base_dir() / "output"))
        self.week_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Excel-Datei wählen oder die Datei neben das Programm legen.")
        self.busy = False
        self.week_map: dict[str, str] = {}
        self.last_result_dir: Path | None = None
        self.last_combined: Path | None = None
        self.last_copy_view: Path | None = None
        self.active_posters: list[ForumPoster] = []

        self._build_style()
        self._build_ui()
        self._autoload_excel()

    def _build_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")

        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI Semibold", 18))
        style.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("TEntry", font=("Segoe UI", 10), padding=4)
        style.configure("TCombobox", font=("Segoe UI", 10), padding=4)
        style.configure(
            "Accent.TButton",
            font=("Segoe UI Semibold", 12),
            padding=(18, 10),
        )
        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 6))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=24)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Event-Post Generator", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Liest Spalte C der Wochen-Excel und erzeugt Forum-DOCX-Dateien zum Copy-Paste.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 18))

        card = tk.Frame(outer, bg=CARD, highlightbackground="#e0d8ce", highlightthickness=1)
        card.pack(fill="x")
        inner = ttk.Frame(card, style="Card.TFrame", padding=18)
        inner.pack(fill="x")
        inner.columnconfigure(1, weight=1)

        self._row(inner, 0, "Excel-Datei", self.excel_var, self._browse_excel)
        ttk.Label(inner, text="Woche", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(12, 0))
        self.week_combo = ttk.Combobox(inner, textvariable=self.week_var, state="readonly")
        self.week_combo.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(12, 0), pady=(12, 0))
        self._row(inner, 2, "Zielordner", self.output_var, self._browse_output)

        actions = ttk.Frame(outer, style="App.TFrame")
        actions.pack(fill="x", pady=18)
        self.generate_btn = tk.Button(
            actions,
            text="DOCX erstellen",
            command=self._on_generate,
            bg=ACCENT,
            fg="white",
            activebackground=ACCENT_HOVER,
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            font=("Segoe UI Semibold", 12),
            padx=22,
            pady=10,
        )
        self.generate_btn.pack(side="left")
        self.generate_all_btn = tk.Button(
            actions,
            text="Alle fehlenden Wochen",
            command=self._on_generate_all,
            bg="#1e2a39",
            fg="white",
            activebackground="#2c3e50",
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            font=("Segoe UI Semibold", 11),
            padx=16,
            pady=10,
        )
        self.generate_all_btn.pack(side="left", padx=(10, 0))

        self.open_folder_btn = ttk.Button(actions, text="Ordner öffnen", command=self._open_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=(10, 0))
        self.open_docx_btn = ttk.Button(actions, text="DOCX öffnen", command=self._open_docx, state="disabled")
        self.open_docx_btn.pack(side="left", padx=(8, 0))
        self.open_copy_btn = ttk.Button(
            actions,
            text="Kopieransicht öffnen",
            command=self._open_copy_view,
            state="disabled",
        )
        self.open_copy_btn.pack(side="left", padx=(8, 0))

        post_actions = ttk.Frame(outer, style="App.TFrame")
        post_actions.pack(fill="x", pady=(0, 18))
        self.auto_post_btn = tk.Button(
            post_actions,
            text="Ausgewählte Woche automatisch posten",
            command=self._on_auto_post,
            bg="#218c53",
            fg="white",
            activebackground="#2da866",
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            font=("Segoe UI Semibold", 12),
            padx=22,
            pady=10,
        )
        self.auto_post_btn.pack(side="left")
        ttk.Label(
            post_actions,
            text=(
                f"Ziel: {BOARD_NAME} → {CATEGORY_NAME} · "
                "Login immer manuell im geöffneten Edge"
            ),
            style="Sub.TLabel",
        ).pack(side="left", padx=(14, 0))

        ttk.Label(outer, text="Protokoll", style="Sub.TLabel").pack(anchor="w")
        log_wrap = tk.Frame(outer, bg="#1e2a39")
        log_wrap.pack(fill="both", expand=True, pady=(8, 0))
        self.log = tk.Text(
            log_wrap,
            height=12,
            wrap="word",
            bg="#1e2a39",
            fg="#f4f6f8",
            insertbackground="#f4f6f8",
            relief="flat",
            font=("Consolas", 10),
            padx=12,
            pady=10,
        )
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")
        self._log(self.status_var.get())

    def _row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, browse) -> None:
        ttk.Label(parent, text=label, style="Card.TLabel").grid(row=row, column=0, sticky="w", pady=(12 if row else 0, 0))
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(12, 8), pady=(12 if row else 0, 0))
        ttk.Button(parent, text="Durchsuchen…", command=browse).grid(
            row=row, column=2, sticky="e", pady=(12 if row else 0, 0)
        )

    def _log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.status_var.set(message)

    def _autoload_excel(self) -> None:
        files = list_excel_files(app_base_dir())
        if not files:
            return
        self.excel_var.set(str(files[0]))
        self._refresh_weeks()

    def _browse_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Excel-Datei wählen",
            initialdir=str(app_base_dir()),
            filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if not path:
            return
        self.excel_var.set(path)
        self._refresh_weeks()

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="Zielordner wählen", initialdir=self.output_var.get() or str(app_base_dir()))
        if path:
            self.output_var.set(path)

    def _refresh_weeks(self) -> None:
        excel = self.excel_var.get().strip()
        self.week_map.clear()
        self.week_combo["values"] = []
        self.week_var.set("")
        if not excel:
            return
        path = Path(excel)
        if not path.exists():
            self._log(f"Datei nicht gefunden: {path}")
            return
        try:
            weeks = load_weeks(path)
        except Exception as exc:
            self._log(f"Wochenblätter konnten nicht gelesen werden: {exc}")
            return

        if not weeks:
            self._log("Keine Wochenblätter im Format YYYYMMDD gefunden.")
            return

        newest = max(weeks)
        output_dir = Path(self.output_var.get().strip() or app_base_dir() / "output")
        labels = []
        existing_count = 0
        for name in reversed(weeks):
            extra = "  (neueste)" if name == newest else ""
            if week_already_generated(output_dir, name):
                extra += "  — vorhanden"
                existing_count += 1
            label = f"{name}  —  {format_week_label(name)}{extra}"
            self.week_map[label] = name
            labels.append(label)
        self.week_combo["values"] = labels
        self.week_var.set(labels[0])
        missing = len(weeks) - existing_count
        self._log(
            f"{path.name}: {len(weeks)} Wochen gefunden, {existing_count} vorhanden, "
            f"{missing} fehlen. Ausgewählt {newest}."
        )

    def _selected_sheet(self) -> str | None:
        label = self.week_var.get().strip()
        return self.week_map.get(label)

    def _on_generate(self) -> None:
        if self.busy:
            return
        paths = self._selected_paths()
        if paths is None:
            return
        excel_path, output_dir = paths
        self._set_busy(True, "Bitte warten…")
        self._log("DOCX wird erstellt…")
        threading.Thread(
            target=self._generate_worker,
            args=(excel_path, output_dir, self._selected_sheet()),
            daemon=True,
        ).start()

    def _on_generate_all(self) -> None:
        if self.busy:
            return
        paths = self._selected_paths()
        if paths is None:
            return
        excel_path, output_dir = paths
        self._set_busy(True, "Wochen werden geprüft…")
        self._log("Alle fehlenden Wochen werden erzeugt. Vorhandene Ordner bleiben unverändert.")
        threading.Thread(
            target=self._generate_all_worker,
            args=(excel_path, output_dir),
            daemon=True,
        ).start()

    def _on_auto_post(self) -> None:
        if self.busy:
            return
        paths = self._selected_paths()
        sheet = self._selected_sheet()
        if paths is None or not sheet:
            return
        excel_path, output_dir = paths
        self._set_busy(True, "Vorbereitung …")
        self._log("Bereite die ausgewählte Woche für das automatische Posting vor …")
        threading.Thread(
            target=self._prepare_post_worker,
            args=(excel_path, output_dir, sheet),
            daemon=True,
        ).start()

    def _prepare_post_worker(
        self,
        excel_path: Path,
        output_dir: Path,
        sheet: str,
    ) -> None:
        try:
            generated = generate_documents(
                excel=excel_path,
                output_dir=output_dir,
                sheet=sheet,
            )
            state = load_posting_state(generated.week_dir)
        except Exception as exc:
            detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.root.after(0, lambda: self._posting_failed(detail))
            return
        self.root.after(0, lambda: self._confirm_auto_post(generated, state))

    def _confirm_auto_post(self, generated, state) -> None:
        if state is not None and state.complete:
            self._set_busy(False)
            messagebox.showwarning(
                APP_TITLE,
                "Diese Woche wurde bereits vollständig automatisch gepostet.\n\n"
                f"{state.thread_url}",
            )
            return

        resume = False
        if state is not None and state.posted_segments > 0:
            resume = messagebox.askyesno(
                APP_TITLE,
                "Es existiert ein unvollständiger Posting-Lauf.\n\n"
                f"Bereits gepostet: {state.posted_segments}/{state.total_segments}\n"
                f"Thread: {state.thread_url}\n\n"
                "Soll bei der nächsten fehlenden Antwort fortgesetzt werden?",
            )
            if not resume:
                self._set_busy(False)
                self._log("Automatisches Posting wurde vor der Wiederaufnahme abgebrochen.")
                return

        confirmation = messagebox.askyesno(
            "Forum-Beiträge wirklich absenden?",
            (
                f"Bereich: {BOARD_NAME}\n"
                f"Kategorie: {CATEGORY_NAME}\n"
                f"Titel: {generated.post.title}\n"
                f"Beiträge: {len(generated.chunks)}\n\n"
                "Die sichtbaren [Event-Bild hier einfügen]-Platzhalter werden "
                "mitgepostet und können anschließend bearbeitet werden.\n\n"
                "Jetzt verbindlich im Forum posten?"
            ),
            icon="warning",
        )
        if not confirmation:
            self._set_busy(False)
            self._log("Automatisches Posting wurde vor dem Absenden abgebrochen.")
            return

        self._log(
            f"Starte bestätigtes Posting: {BOARD_NAME} → {CATEGORY_NAME} → "
            f"{generated.post.title}"
        )
        threading.Thread(
            target=self._post_worker,
            args=(generated, resume),
            daemon=True,
        ).start()

    def _post_worker(self, generated, resume: bool) -> None:
        poster = ForumPoster(
            progress=lambda message: self.root.after(
                0, lambda m=message: self._log(m)
            ),
            manual_action=self._manual_action_from_worker,
        )
        self.active_posters.append(poster)
        try:
            result = poster.post(generated, resume=resume)
        except Exception as exc:
            detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.root.after(0, lambda: self._posting_failed(detail))
            return
        self.root.after(0, lambda: self._posting_done(result))

    def _manual_action_from_worker(self, message: str) -> bool:
        finished = threading.Event()
        answer = {"value": False}

        def ask() -> None:
            answer["value"] = messagebox.askokcancel(
                "Manueller Login-Schritt",
                f"{message}\n\nKlicke danach hier auf OK, um fortzufahren.",
                icon="info",
            )
            finished.set()

        self.root.after(0, ask)
        finished.wait()
        return answer["value"]

    def _posting_failed(self, detail: str) -> None:
        self._set_busy(False)
        self._log(f"Automatisches Posting gestoppt: {detail}")
        messagebox.showerror(
            APP_TITLE,
            "Das automatische Posting wurde gestoppt.\n"
            "Der Browser bleibt zur Kontrolle geöffnet.\n\n"
            f"{detail}",
        )

    def _posting_done(self, result) -> None:
        self._set_busy(False)
        self._log(f"Forum-Thread vollständig gepostet: {result.state.thread_url}")
        messagebox.showinfo(
            APP_TITLE,
            "Automatisches Posting abgeschlossen.\n\n"
            f"Beiträge: {result.state.posted_segments}/{result.state.total_segments}\n"
            f"Browser: {result.browser_name}\n"
            f"Thread: {result.state.thread_url}\n\n"
            "Der Browser bleibt geöffnet, damit Bilder nachgetragen werden können.",
        )

    def _selected_paths(self) -> tuple[Path, Path] | None:
        excel = self.excel_var.get().strip()
        output = self.output_var.get().strip()
        if not excel:
            messagebox.showwarning(APP_TITLE, "Bitte zuerst eine Excel-Datei wählen.")
            return None
        excel_path = Path(excel)
        if not excel_path.exists():
            messagebox.showerror(APP_TITLE, f"Excel-Datei nicht gefunden:\n{excel_path}")
            return None
        output_dir = Path(output) if output else app_base_dir() / "output"
        return excel_path, output_dir

    def _set_busy(self, busy: bool, generate_text: str = "DOCX erstellen") -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.generate_btn.configure(state=state, text=generate_text if busy else "DOCX erstellen")
        self.generate_all_btn.configure(state=state)
        self.auto_post_btn.configure(state=state)

    def _generate_worker(self, excel: Path, output_dir: Path, sheet: str | None) -> None:
        try:
            result = generate_documents(excel=excel, output_dir=output_dir, sheet=sheet)
        except Exception as exc:
            detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.root.after(0, lambda: self._generate_failed(detail))
            return
        self.root.after(0, lambda: self._generate_done(result))

    def _generate_all_worker(self, excel: Path, output_dir: Path) -> None:
        try:
            batch = generate_missing_weeks(
                excel=excel,
                output_dir=output_dir,
                progress=lambda message: self.root.after(0, lambda m=message: self._log(m)),
            )
        except Exception as exc:
            detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.root.after(0, lambda: self._generate_failed(detail))
            return
        self.root.after(0, lambda: self._generate_all_done(batch, output_dir))

    def _generate_failed(self, detail: str) -> None:
        self._set_busy(False)
        self._log(f"Fehler: {detail}")
        messagebox.showerror(APP_TITLE, f"Die Datei konnte nicht erstellt werden.\n\n{detail}")

    def _generate_done(self, result) -> None:
        self._set_busy(False)
        self.last_result_dir = result.week_dir
        self.last_combined = result.combined
        self.last_copy_view = result.copy_view
        self.open_folder_btn.configure(state="normal")
        self.open_docx_btn.configure(state="normal")
        self.open_copy_btn.configure(state="normal")

        week = result.post.source_sheet
        self._log(f"Fertig: {result.post.title}")
        self._log(f"Woche {week} ({format_week_label(week)}), {len(result.post.events)} Events")
        self._log(f"Ordner: {result.week_dir}")
        for chunk, path in zip(result.chunks, result.chunk_files, strict=True):
            extra = ""
            if chunk.event_from and chunk.event_to:
                extra = f", Events {chunk.event_from}–{chunk.event_to}"
            self._log(f"  Beitrag {chunk.index}/{chunk.total}: {chunk.char_count} Zeichen{extra} → {path.name}")
        self._log(f"Gesamt-DOCX: {result.combined.name}")
        self._log(f"Kopieransicht: {result.copy_view.name}")
        self._refresh_weeks()
        messagebox.showinfo(
            APP_TITLE,
            f"Fertig.\n\n{len(result.chunk_files)} Forum-Beiträge + 1 Gesamt-DOCX\nim Ordner:\n{result.week_dir}",
        )

    def _generate_all_done(self, batch, output_dir: Path) -> None:
        self._set_busy(False)
        self.last_result_dir = output_dir
        if batch.created:
            self.last_combined = batch.created[-1].combined
            self.last_copy_view = batch.created[-1].copy_view
            self.open_docx_btn.configure(state="normal")
            self.open_copy_btn.configure(state="normal")
        self.open_folder_btn.configure(state="normal")
        self._log(
            f"Fertig: {len(batch.created)} neu erstellt, "
            f"{len(batch.skipped)} übersprungen, "
            f"{len(batch.failed)} Fehler"
        )
        if batch.failed:
            for sheet_name, error in batch.failed[:8]:
                self._log(f"  {format_week_label(sheet_name)}: {error}")
            if len(batch.failed) > 8:
                self._log(f"  … und {len(batch.failed) - 8} weitere Fehler")
        self._refresh_weeks()
        messagebox.showinfo(
            APP_TITLE,
            (
                f"Alle fehlenden Wochen fertig.\n\n"
                f"Neu erstellt: {len(batch.created)}\n"
                f"Schon vorhanden: {len(batch.skipped)}\n"
                f"Fehler: {len(batch.failed)}\n\n"
                f"Ordner:\n{output_dir}"
            ),
        )

    def _open_folder(self) -> None:
        if self.last_result_dir and self.last_result_dir.exists():
            open_path(self.last_result_dir)

    def _open_docx(self) -> None:
        if self.last_combined and self.last_combined.exists():
            open_path(self.last_combined)

    def _open_copy_view(self) -> None:
        if self.last_copy_view and self.last_copy_view.exists():
            open_path(self.last_copy_view)


def main() -> None:
    enable_dpi_awareness()
    root = tk.Tk()
    EventPostApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
