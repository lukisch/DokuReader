#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dokumentenbibliothek mit Themen, Vorschau, Gelesen/ungelesen, Doppelklick-Öffnen
und Sammel-PDF-Export (alle/gelesene/ungelesene) mit vielen Fallbacks.

Features:
- GUI mit Tkinter (eine einzelne .py-Datei)
- Themen anlegen/umbenennen/löschen
- Dateien pro Thema verwalten (nur Verweise, Originaldateien bleiben unberührt)
- Drag & Drop hinzufügen (optional: tkinterdnd2)
- Rechtsklick: als gelesen markieren (grün + ✓), Markierung entfernen, aus Bibliothek entfernen
- Doppelklick: Datei im Standardprogramm des OS öffnen (Windows/macOS/Linux)
- Vorschau:
  * Bilder (Pillow)
  * Textdateien (UTF-8 -> Latin-1 -> Hexdump-Fallback)
  * PDF (pdf2image oder PyMuPDF; sonst Metadaten)
  * DOCX (python-docx), ODT (odfpy); sonst Metadaten
- Export Sammel-PDF auf Desktop: Thema_alle.pdf / Thema_gelesene.pdf / Thema_ungelesene.pdf
  * PDFs direkt
  * TXT/Bilder -> PDF (ReportLab und/oder Pillow)
  * DOC/DOCX/ODT/RTF -> PDF via LibreOffice (headless) oder Word COM (Windows, pywin32)
  * Merge via pypdf oder PyPDF2
- Persistenz: JSON im Home-Verzeichnis (.dokubibliothek_state.json)
"""

import os
import sys
import json
import shutil
import subprocess
import tempfile
import platform
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# Optionale Bibliotheken
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.units import cm
    from reportlab.lib.utils import ImageReader
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

try:
    from pypdf import PdfMerger
except Exception:
    try:
        from PyPDF2 import PdfMerger
    except Exception:
        PdfMerger = None

# PDF-Vorschau optional
try:
    from pdf2image import convert_from_path
    PDF2IMG_AVAILABLE = True
except Exception:
    PDF2IMG_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except Exception:
    PYMUPDF_AVAILABLE = False

# Office-Textauszug optional
try:
    import docx
    DOCPREVIEW_AVAILABLE = True
except Exception:
    DOCPREVIEW_AVAILABLE = False

try:
    from odf import text as odf_text, teletype
    from odf.opendocument import load as odf_load
    ODFPREVIEW_AVAILABLE = True
except Exception:
    ODFPREVIEW_AVAILABLE = False

# Drag&Drop optional
try:
    import tkinterdnd2 as tkdnd
    TKDND_AVAILABLE = True
except Exception:
    TKDND_AVAILABLE = False


APP_NAME = "Dokumentenbibliothek"
STATE_FILE = str(Path.home() / ".dokubibliothek_state.json")

SUPPORTED_EXTS = {
    ".txt", ".doc", ".docx", ".pdf", ".odt", ".rtf",
    ".jpg", ".jpeg", ".gif", ".png"
}
IMAGE_EXTS = {".jpg", ".jpeg", ".gif", ".png"}
WORD_EXTS = {".doc", ".docx", ".odt", ".rtf"}
TXT_EXTS = {".txt"}
PDF_EXTS = {".pdf"}

# Vorschau-Einstellungen
TXT_PREVIEW_CHARS = 5000
OFFICE_PREVIEW_PARAGRAPHS = 30


def human_size(num_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} TB"


def desktop_path() -> Path:
    p = Path.home() / "Desktop"
    return p if p.exists() else Path.home()


class State:
    """
    Verwaltet den Anwendungszustand (Themen, Dokumente, Gelesen-Status).

    Attributes:
        topics: dict[str, list[dict]] - Themen mit zugeordneten Dokumenten
        current_topic: str | None - Aktuell ausgewähltes Thema
    """
    def __init__(self):
        self.topics: dict[str, list[dict]] = {}
        self.current_topic: str | None = None

    def load(self):
        """Lädt den Zustand aus der JSON-Datei (~/.dokubibliothek_state.json)."""
        if os.path.isfile(STATE_FILE):
            try:
                with open(STATE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.topics = data.get("topics", {})
                    self.current_topic = data.get("current_topic")
            except (OSError, json.JSONDecodeError):
                # Ignorieren, falls Zustand nicht gelesen werden kann
                pass

    def save(self):
        """Speichert den aktuellen Zustand in die JSON-Datei."""
        try:
            with open(STATE_FILE, "w", encoding="utf-8") as f:
                json.dump({"topics": self.topics, "current_topic": self.current_topic},
                          f, ensure_ascii=False, indent=2)
        except (OSError, TypeError):
            pass

    def ensure_topic(self, topic: str):
        """
        Stellt sicher, dass ein Thema existiert (erstellt leere Liste falls nicht vorhanden).

        Args:
            topic: Name des Themas
        """
        if topic not in self.topics:
            self.topics[topic] = []

    def add_docs(self, topic: str, paths) -> int:
        """
        Fügt Dokumente zu einem Thema hinzu (nur unterstützte Dateitypen, keine Duplikate).

        Args:
            topic: Name des Themas
            paths: Liste von Dateipfaden

        Returns:
            Anzahl der tatsächlich hinzugefügten Dokumente
        """
        self.ensure_topic(topic)
        known = {d["path"] for d in self.topics[topic]}
        added = 0
        for p in paths:
            if os.path.isfile(p) and Path(p).suffix.lower() in SUPPORTED_EXTS and p not in known:
                self.topics[topic].append({"path": p, "read": False})
                added += 1
        return added

    def remove_doc(self, topic: str, path: str):
        """
        Entfernt ein Dokument aus einem Thema.

        Args:
            topic: Name des Themas
            path: Pfad des zu entfernenden Dokuments
        """
        self.topics[topic] = [d for d in self.topics.get(topic, []) if d["path"] != path]

    def set_read(self, topic: str, path: str, is_read: bool):
        """
        Setzt den Gelesen-Status eines Dokuments.

        Args:
            topic: Name des Themas
            path: Pfad des Dokuments
            is_read: True = gelesen, False = ungelesen
        """
        for d in self.topics.get(topic, []):
            if d["path"] == path:
                d["read"] = is_read

    def list_docs(self, topic: str):
        """
        Gibt alle Dokumente eines Themas zurück.

        Args:
            topic: Name des Themas

        Returns:
            Liste von Dokumenten (dicts mit 'path' und 'read')
        """
        return self.topics.get(topic, [])


class App(tk.Tk if not TKDND_AVAILABLE else tkdnd.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1200x700")
        self.minsize(1000, 600)

        self.state_model = State()
        self.state_model.load()

        self._build_ui()
        self._reload_topics()
        if self.state_model.current_topic:
            self._select_topic(self.state_model.current_topic)

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # Linke Spalte: Themen
        left = ttk.Frame(paned)
        paned.add(left, weight=1)
        ttk.Label(left, text="Themen").pack(anchor="w", padx=8, pady=(8, 2))
        self.topic_list = tk.Listbox(left)
        self.topic_list.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))
        self.topic_list.bind("<<ListboxSelect>>", self.on_topic_select)
        btns = ttk.Frame(left)
        btns.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(btns, text="Neu", command=self.add_topic).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Umbenennen", command=self.rename_topic).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Löschen", command=self.delete_topic).pack(side=tk.LEFT, padx=2)

        # Mitte: Dokumente
        center = ttk.Frame(paned)
        paned.add(center, weight=3)
        ttk.Label(center, text="Dokumente im Thema").pack(anchor="w", padx=8, pady=(8, 2))
        self.doc_tree = ttk.Treeview(center, columns=("typ", "größe"), show="tree headings", selectmode="browse")
        self.doc_tree.heading("#0", text="Name", anchor="w")
        self.doc_tree.heading("typ", text="Typ")
        self.doc_tree.heading("größe", text="Größe")
        self.doc_tree.column("#0", width=550, anchor="w")
        self.doc_tree.column("typ", width=100, anchor="w")
        self.doc_tree.column("größe", width=100, anchor="e")
        self.doc_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))

        # Tag für gelesene Einträge: grün + fett
        self.doc_tree.tag_configure("read", foreground="#08660c", font=("TkDefaultFont", 9, "bold"))
        # Optionaler Hintergrund statt Textfarbe:
        # self.doc_tree.tag_configure("read", background="#d9f7d9")

        # Binds
        self.doc_tree.bind("<Button-3>", self.on_doc_right_click)
        self.doc_tree.bind("<Double-1>", self.on_doc_double_click)
        self.doc_tree.bind("<<TreeviewSelect>>", self.on_doc_select)

        # Drag & Drop
        addbar = ttk.Frame(center)
        addbar.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(addbar, text="Hinzufügen", command=self.add_files_dialog).pack(side=tk.RIGHT)
        if TKDND_AVAILABLE:
            self.doc_tree.drop_target_register('*')  # type: ignore
            self.doc_tree.dnd_bind('<<Drop>>', self.on_drop)  # type: ignore

        # Rechte Spalte: Vorschau + Export
        right = ttk.Frame(paned)
        paned.add(right, weight=2)
        ttk.Label(right, text="Vorschau").pack(anchor="w", padx=8, pady=(8, 2))
        self.preview = tk.Canvas(right, bg="#fafafa", height=320)
        self.preview.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 6))
        self.preview_text = tk.Text(right, height=10, wrap="word")
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        export_frame = ttk.LabelFrame(right, text="Sammel-PDF")
        export_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.filter_var = tk.StringVar(value="alle")
        ttk.Radiobutton(export_frame, text="Alle", variable=self.filter_var, value="alle").pack(side=tk.LEFT, padx=6, pady=6)
        ttk.Radiobutton(export_frame, text="Gelesene", variable=self.filter_var, value="gelesene").pack(side=tk.LEFT, padx=6, pady=6)
        ttk.Radiobutton(export_frame, text="Ungelesene", variable=self.filter_var, value="ungelesene").pack(side=tk.LEFT, padx=6, pady=6)
        ttk.Button(export_frame, text="Sammel-PDF erzeugen", command=self.create_collection_pdf).pack(side=tk.RIGHT, padx=6, pady=6)

        # Kontextmenü
        self.doc_menu = tk.Menu(self, tearoff=0)
        self.doc_menu.add_command(label="Als gelesen markieren", command=lambda: self.set_selected_read(True))
        self.doc_menu.add_command(label="Gelesen-Markierung entfernen", command=lambda: self.set_selected_read(False))
        self.doc_menu.add_separator()
        self.doc_menu.add_command(label="Aus Bibliothek entfernen", command=self.remove_selected_doc)

    # Themen-Handling
    def _reload_topics(self):
        self.topic_list.delete(0, tk.END)
        for t in sorted(self.state_model.topics.keys(), key=str.lower):
            self.topic_list.insert(tk.END, t)

    def _select_topic(self, topic: str):
        self.state_model.current_topic = topic
        self._reload_docs()

    def on_topic_select(self, _=None):
        sel = self.topic_list.curselection()
        if not sel:
            return
        topic = self.topic_list.get(sel[0])
        self._select_topic(topic)
        self.state_model.save()

    def add_topic(self):
        name = simpledialog.askstring("Neues Thema", "Name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.state_model.topics:
            messagebox.showwarning("Hinweis", "Thema existiert bereits.")
            return
        self.state_model.ensure_topic(name)
        self.state_model.save()
        self._reload_topics()
        self._select_topic(name)

    def rename_topic(self):
        sel = self.topic_list.curselection()
        if not sel:
            return
        old = self.topic_list.get(sel[0])
        new = simpledialog.askstring("Thema umbenennen", "Neuer Name:", initialvalue=old)
        if not new:
            return
        new = new.strip()
        if not new or new == old:
            return
        if new in self.state_model.topics:
            messagebox.showwarning("Hinweis", "Ein Thema mit diesem Namen existiert bereits.")
            return
        self.state_model.topics[new] = self.state_model.topics.pop(old)
        if self.state_model.current_topic == old:
            self.state_model.current_topic = new
        self.state_model.save()
        self._reload_topics()
        self._select_topic(new)

    def delete_topic(self):
        sel = self.topic_list.curselection()
        if not sel:
            return
        topic = self.topic_list.get(sel[0])
        if messagebox.askyesno("Bestätigen", f"Thema '{topic}' entfernen? (Dateien bleiben am Originalort)"):
            self.state_model.topics.pop(topic, None)
            if self.state_model.current_topic == topic:
                self.state_model.current_topic = None
            self.state_model.save()
            self._reload_topics()
            self.doc_tree.delete(*self.doc_tree.get_children())
            self.clear_preview()

    # Dokumente-Handling
    def _reload_docs(self):
        self.doc_tree.delete(*self.doc_tree.get_children())
        topic = self.state_model.current_topic
        if not topic:
            return
        for d in self.state_model.list_docs(topic):
            path = d["path"]
            name = os.path.basename(path)
            ext = Path(path).suffix.lower()
            try:
                size = human_size(os.path.getsize(path))
            except OSError:
                size = "?"
            tags = ()
            if d.get("read"):
                tags = ("read",)
                name = "✓ " + name
            self.doc_tree.insert("", "end", iid=path, text=name,
                                 values=(ext[1:].upper(), size), tags=tags)
        self.clear_preview()

    def add_files_dialog(self):
        topic = self.state_model.current_topic
        if not topic:
            messagebox.showinfo("Hinweis", "Bitte zuerst ein Thema auswählen.")
            return
        paths = filedialog.askopenfilenames(
            title="Dateien hinzufügen",
            filetypes=[("Unterstützte Dateien", "*.txt;*.doc;*.docx;*.pdf;*.odt;*.rtf;*.jpg;*.jpeg;*.gif;*.png")]
        )
        if not paths:
            return
        added = self.state_model.add_docs(topic, paths)
        self.state_model.save()
        self._reload_docs()
        if added == 0:
            messagebox.showinfo("Hinweis", "Keine neuen unterstützten Dateien hinzugefügt.")

    def on_drop(self, event):
        topic = self.state_model.current_topic
        if not topic:
            return
        paths = self._split_dnd_paths(event.data)
        added = self.state_model.add_docs(topic, paths)
        self.state_model.save()
        self._reload_docs()
        if added == 0:
            self.status_info("Keine neuen unterstützten Dateien per Drag&Drop hinzugefügt.")

    @staticmethod
    def _split_dnd_paths(data: str):
        # Formate wie {C:\Pfad mit Leerzeichen\file.pdf} /home/user/x.pdf ...
        res = []
        cur = []
        in_brace = False
        for ch in data:
            if ch == "{":
                in_brace = True
                cur = []
            elif ch == "}":
                in_brace = False
                res.append("".join(cur))
                cur = []
            elif ch == " " and not in_brace:
                if cur:
                    res.append("".join(cur))
                    cur = []
            else:
                cur.append(ch)
        if cur:
            res.append("".join(cur))
        return [p.strip() for p in res if p.strip()]

    def on_doc_right_click(self, event):
        iid = self.doc_tree.identify_row(event.y)
        if iid:
            self.doc_tree.selection_set(iid)
            self.doc_menu.tk_popup(event.x_root, event.y_root)

    def on_doc_double_click(self, _=None):
        sel = self.doc_tree.selection()
        if not sel:
            return
        path = sel[0]
        try:
            if platform.system() == "Windows":
                os.startfile(path)  # type: ignore
            elif platform.system() == "Darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Datei nicht öffnen:\n{e}")

    def set_selected_read(self, is_read: bool):
        topic = self.state_model.current_topic
        sel = self.doc_tree.selection()
        if not topic or not sel:
            return
        path = sel[0]
        self.state_model.set_read(topic, path, is_read)
        self.state_model.save()
        self._reload_docs()

    def remove_selected_doc(self):
        topic = self.state_model.current_topic
        sel = self.doc_tree.selection()
        if not topic or not sel:
            return
        path = sel[0]
        if messagebox.askyesno("Entfernen", "Dokument aus der Bibliothek entfernen?\n(Originaldatei bleibt erhalten)"):
            self.state_model.remove_doc(topic, path)
            self.state_model.save()
            self._reload_docs()

    def on_doc_select(self, _=None):
        sel = self.doc_tree.selection()
        if not sel:
            self.clear_preview()
            return
        self.show_preview(sel[0])

    # Vorschau
    def clear_preview(self):
        self.preview.delete("all")
        self.preview_text.delete("1.0", tk.END)
        self.preview.create_text(10, 10, anchor="nw", text="Keine Vorschau", fill="#666")

    def show_preview(self, path: str):
        self.preview.delete("all")
        self.preview_text.delete("1.0", tk.END)
        ext = Path(path).suffix.lower()
        try:
            if ext in IMAGE_EXTS and PIL_AVAILABLE:
                img = Image.open(path)
                cw = self.preview.winfo_width() or 600
                ch = self.preview.winfo_height() or 320
                img.thumbnail((cw - 20, ch - 20))
                self._preview_img = ImageTk.PhotoImage(img)
                self.preview.create_image(10, 10, anchor="nw", image=self._preview_img)
                self.preview_text.insert("1.0", f"Bild: {img.width}x{img.height}px\n{path}")
            elif ext in TXT_EXTS:
                content = None
                for enc in ["utf-8", "latin-1"]:
                    try:
                        with open(path, "r", encoding=enc, errors="replace") as f:
                            content = f.read(TXT_PREVIEW_CHARS)
                        break
                    except (OSError, UnicodeDecodeError):
                        continue
                if content is None:
                    with open(path, "rb") as f:
                        content = f.read(256).hex(" ")
                self.preview.create_text(10, 10, anchor="nw", text="Textdatei", fill="#666")
                self.preview_text.insert("1.0", content if content else "(Leer)")
            elif ext in PDF_EXTS and (PDF2IMG_AVAILABLE or PYMUPDF_AVAILABLE) and PIL_AVAILABLE:
                img = None
                if PDF2IMG_AVAILABLE:
                    try:
                        pages = convert_from_path(path, first_page=1, last_page=1)
                        if pages:
                            img = pages[0]
                    except Exception:
                        img = None
                if img is None and PYMUPDF_AVAILABLE:
                    try:
                        doc = fitz.open(path)
                        if len(doc) > 0:
                            page = doc[0]
                            pix = page.get_pixmap()
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    except Exception:
                        img = None
                if img is not None:
                    cw = self.preview.winfo_width() or 600
                    ch = self.preview.winfo_height() or 320
                    img.thumbnail((cw - 20, ch - 20))
                    self._preview_img = ImageTk.PhotoImage(img)
                    self.preview.create_image(10, 10, anchor="nw", image=self._preview_img)
                self.preview_text.insert("1.0", f"PDF: {os.path.basename(path)}\n{path}")
            elif ext == ".docx" and DOCPREVIEW_AVAILABLE:
                try:
                    doc = docx.Document(path)
                    text = "\n".join(p.text for p in doc.paragraphs[:OFFICE_PREVIEW_PARAGRAPHS])
                    self.preview.create_text(10, 10, anchor="nw", text="DOCX-Vorschau", fill="#666")
                    self.preview_text.insert("1.0", text if text.strip() else "(Kein Textinhalt erkannt)")
                except Exception as e:
                    self.preview_text.insert("1.0", f"(Keine DOCX-Vorschau möglich)\n{e}")
            elif ext == ".odt" and ODFPREVIEW_AVAILABLE:
                try:
                    odt_doc = odf_load(path)
                    paras = odt_doc.getElementsByType(odf_text.P)  # type: ignore
                    text_content = "\n".join(teletype.extractText(p) for p in paras[:OFFICE_PREVIEW_PARAGRAPHS])
                    self.preview.create_text(10, 10, anchor="nw", text="ODT-Vorschau", fill="#666")
                    self.preview_text.insert("1.0", text_content if text_content.strip() else "(Kein Textinhalt erkannt)")
                except Exception as e:
                    self.preview_text.insert("1.0", f"(Keine ODT-Vorschau möglich)\n{e}")
            else:
                # Generische Metadaten
                try:
                    size = human_size(os.path.getsize(path))
                except Exception:
                    size = "?"
                self.preview.create_text(10, 10, anchor="nw", text="Keine Vorschau verfügbar", fill="#666")
                self.preview_text.insert("1.0", f"Datei: {os.path.basename(path)}\nTyp: {ext}\nGröße: {size}\nPfad: {path}")
        except Exception as e:
            self.preview.create_text(10, 10, anchor="nw", text="Vorschau-Fehler", fill="#666")
            self.preview_text.insert("1.0", f"Fehler: {e}")

    # Export Sammel-PDF
    def create_collection_pdf(self):
        topic = self.state_model.current_topic
        if not topic:
            messagebox.showinfo("Hinweis", "Bitte zuerst ein Thema auswählen.")
            return
        filter_mode = self.filter_var.get()
        threading.Thread(target=self._create_collection_pdf_worker, args=(topic, filter_mode), daemon=True).start()

    def _create_collection_pdf_worker(self, topic: str, filter_mode: str):
        self._set_busy(True)
        try:
            docs = self.state_model.list_docs(topic)
            if filter_mode == "gelesene":
                docs = [d for d in docs if d.get("read")]
            elif filter_mode == "ungelesene":
                docs = [d for d in docs if not d.get("read")]

            if not docs:
                self.status_info("Keine passenden Dokumente für das Sammel-PDF.")
                return

            out_path = desktop_path() / f"{topic}_{filter_mode}.pdf"

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)
                pdf_parts: list[str] = []
                log_lines: list[str] = []

                for d in docs:
                    src = d["path"]
                    ext = Path(src).suffix.lower()
                    try:
                        if ext in PDF_EXTS:
                            pdf_parts.append(src)
                            log_lines.append(f"OK: PDF übernommen: {src}")
                        elif ext in TXT_EXTS:
                            pdfp = self._txt_to_pdf(src, tmpdir)
                            if pdfp:
                                pdf_parts.append(pdfp)
                                log_lines.append(f"OK: TXT -> PDF: {src}")
                            else:
                                log_lines.append(f"Übersprungen (TXT ohne ReportLab): {src}")
                        elif ext in IMAGE_EXTS:
                            pdfp = self._image_to_pdf(src, tmpdir)
                            if pdfp:
                                pdf_parts.append(pdfp)
                                log_lines.append(f"OK: Bild -> PDF: {src}")
                            else:
                                log_lines.append(f"Übersprungen (Bild ohne ReportLab/Pillow): {src}")
                        elif ext in WORD_EXTS:
                            pdfp = self._office_to_pdf(src, tmpdir)
                            if pdfp:
                                pdf_parts.append(pdfp)
                                log_lines.append(f"OK: Office -> PDF: {src}")
                            else:
                                log_lines.append(f"Übersprungen (LibreOffice/Word nicht verfügbar): {src}")
                        else:
                            log_lines.append(f"Übersprungen (nicht unterstützt): {src}")
                    except Exception as e:
                        log_lines.append(f"Fehler bei {src}: {e}")

                if not pdf_parts:
                    self.status_info("Keine Dateien konnten in PDF überführt werden.")
                    return

                if not self._merge_pdfs(pdf_parts, out_path):
                    self.status_info("Konnte Sammel-PDF nicht erstellen (PDF-Merge-Bibliothek fehlt?).")
                    return

                summary = "Sammel-PDF erstellt:\n" + str(out_path)
                self.after(0, lambda: messagebox.showinfo("Erfolg", summary))
        finally:
            self._set_busy(False)

    # Busy/Status
    def _set_busy(self, busy: bool):
        def apply():
            self.config(cursor="watch" if busy else "")
            self.update_idletasks()
        self.after(0, apply)

    def status_info(self, msg: str):
        self.after(0, lambda: messagebox.showinfo("Info", msg))

    # Konvertierungen
    def _txt_to_pdf(self, path: str, tmpdir: Path) -> str | None:
        if not REPORTLAB_AVAILABLE:
            return None
        out = tmpdir / (Path(path).stem + "_txt.pdf")
        try:
            c = rl_canvas.Canvas(str(out), pagesize=A4)
            width, height = A4
            margin = 2 * cm
            y = height - margin
            line_height = 12
            c.setFont("Helvetica", 11)
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    line = line.rstrip("\n")
                    while line:
                        max_chars = int((width - 2 * margin) / 6)  # Näherung
                        part = line[:max_chars]
                        c.drawString(margin, y, part)
                        y -= line_height
                        line = line[len(part):]
                        if y < margin:
                            c.showPage()
                            c.setFont("Helvetica", 11)
                            y = height - margin
            c.showPage()
            c.save()
            return str(out)
        except Exception:
            try:
                if out.exists():
                    out.unlink()
            except Exception:
                pass
            return None

    def _image_to_pdf(self, path: str, tmpdir: Path) -> str | None:
        out = tmpdir / (Path(path).stem + "_img.pdf")
        # Bevorzugt ReportLab (saubere Skalierung)
        if REPORTLAB_AVAILABLE:
            try:
                c = rl_canvas.Canvas(str(out), pagesize=A4)
                width, height = A4
                img = ImageReader(path)
                iw, ih = img.getSize()
                max_w = width - 2 * cm
                max_h = height - 2 * cm
                scale = min(max_w / iw, max_h / ih)
                w = iw * scale
                h = ih * scale
                x = (width - w) / 2
                y = (height - h) / 2
                c.drawImage(img, x, y, w, h, preserveAspectRatio=True)
                c.showPage()
                c.save()
                return str(out)
            except Exception:
                pass
        # Fallback: Pillow direkt nach PDF
        if PIL_AVAILABLE:
            try:
                img = Image.open(path).convert("RGB")
                img.save(out, "PDF", resolution=150.0)
                return str(out)
            except Exception:
                pass
        return None

    def _office_to_pdf(self, path: str, tmpdir: Path) -> str | None:
        # 1) LibreOffice headless (soffice/libreoffice)
        for cand in ["soffice", "libreoffice"]:
            if shutil.which(cand):
                try:
                    subprocess.run(
                        [cand, "--headless", "--convert-to", "pdf", "--outdir", str(tmpdir), path],
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180, check=False
                    )
                    out = tmpdir / (Path(path).stem + ".pdf")
                    if out.exists():
                        return str(out)
                except Exception:
                    pass
        # 2) Microsoft Word COM (nur Windows; öffnet DOC/DOCX/RTF; ODT oft nicht)
        if platform.system() == "Windows":
            try:
                import win32com.client  # pywin32
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(path)
                out_path = str(tmpdir / (Path(path).stem + ".pdf"))
                wdFormatPDF = 17
                doc.SaveAs(out_path, FileFormat=wdFormatPDF)
                doc.Close(False)
                word.Quit()
                if os.path.exists(out_path):
                    return out_path
            except Exception:
                pass
        return None

    def _merge_pdfs(self, pdf_paths: list[str], out_path: Path) -> bool:
        if not PdfMerger:
            return False
        try:
            merger = PdfMerger()
            for p in pdf_paths:
                try:
                    merger.append(p)
                except Exception:
                    # Ignoriere defekte Einzel-PDFs
                    continue
            with open(out_path, "wb") as f:
                merger.write(f)
            try:
                merger.close()
            except Exception:
                pass
            return True
        except Exception:
            return False

    def on_close(self):
        self.state_model.save()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
