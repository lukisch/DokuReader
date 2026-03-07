# DokuReader - Dokumentenbibliothek

Eine einfache Desktop-Anwendung zur Verwaltung und Organisation von Dokumenten nach Themen mit Vorschau-Funktion und PDF-Export.

## Features

- **Themen-Organisation**: Themen fuer Dokumente erstellen, umbenennen und loeschen
- **Gelesen/Ungelesen**: Dokumente als gelesen markieren (gruen mit Haekchen)
- **Vorschau**:
  - Bilder (JPG, PNG, GIF)
  - PDF-Dokumente (erste Seite)
  - Textdateien (TXT)
  - Office-Dokumente (DOCX, ODT)
- **Drag & Drop**: Dateien einfach per Drag & Drop hinzufuegen (optional)
- **Doppelklick-Oeffnen**: Dokumente in der Standardanwendung oeffnen
- **Batch-PDF-Export**: Alle, gelesene oder ungelesene Dokumente als einzelnes PDF exportieren
  - Unterstuetzt: PDF, TXT, Bilder, DOC, DOCX, ODT, RTF
  - Automatische Konvertierung zu PDF (LibreOffice oder MS Word)
- **Plattformuebergreifend**: Windows, macOS, Linux

## Technische Details

- Python 3.10+ mit Tkinter
- Einzeldatei-Anwendung (756 Zeilen)
- JSON-basierte Speicherung im Home-Verzeichnis
- Nur Referenzen: Originaldateien bleiben unberuehrt

## Screenshots

![Hauptfenster](screenshots/main.png)

## Installation

### Erforderliche Abhaengigkeiten

```bash
pip install -r requirements.txt
```

### Optionale Abhaengigkeiten

Fuer volle Funktionalitaet:
- `pdf2image` oder `PyMuPDF` - PDF-Vorschau
- `tkinterdnd2` - Drag & Drop
- `python-docx` - DOCX-Vorschau
- `odfpy` - ODT-Vorschau
- `reportlab` - TXT/Bild zu PDF-Konvertierung
- `pypdf` oder `PyPDF2` - PDF-Zusammenfuehrung
- `pywin32` (Windows) - Word COM fuer Office-Konvertierung

### LibreOffice (fuer Office → PDF Konvertierung)

Fuer beste Unterstuetzung von DOC/DOCX/ODT/RTF → PDF:
- **Linux**: `sudo apt-get install libreoffice`
- **macOS**: `brew install --cask libreoffice`
- **Windows**: Download von https://www.libreoffice.org/

## Verwendung

```bash
python DokuReader.py
```

Oder via START.bat (Windows):
```bash
START.bat
```

## Datenspeicherung

Status wird gespeichert in: `~/.dokubibliothek_state.json`

## Unterstuetzte Dateiformate

- Dokumente: `.txt`, `.doc`, `.docx`, `.pdf`, `.odt`, `.rtf`
- Bilder: `.jpg`, `.jpeg`, `.gif`, `.png`

## Lizenz

GPL v3 - Siehe [LICENSE](LICENSE)

---

## English

# DokuReader - Document Library

A simple desktop application for managing and organizing documents by topic with preview functionality and PDF export.

### Features

- **Topic Organization**: Create, rename, and delete topics for your documents
- **Read/Unread**: Mark documents as read (green with checkmark)
- **Preview**:
  - Images (JPG, PNG, GIF)
  - PDF documents (first page)
  - Text files (TXT)
  - Office documents (DOCX, ODT)
- **Drag & Drop**: Easily add files via drag and drop (optional)
- **Double-Click Open**: Open documents in the default application
- **Batch PDF Export**: Export all, read, or unread documents as a single PDF
  - Supports: PDF, TXT, images, DOC, DOCX, ODT, RTF
  - Automatic conversion to PDF (LibreOffice or MS Word)
- **Cross-Platform**: Windows, macOS, Linux

### Technical Details

- Python 3.10+ with Tkinter
- Single-file application (756 lines)
- JSON-based persistence in the home directory
- Reference-only: Original files remain untouched

### Screenshots

![Main Window](screenshots/main.png)

### Installation

#### Required Dependencies

```bash
pip install -r requirements.txt
```

#### Optional Dependencies

For full functionality:
- `pdf2image` or `PyMuPDF` - PDF preview
- `tkinterdnd2` - Drag & Drop
- `python-docx` - DOCX preview
- `odfpy` - ODT preview
- `reportlab` - TXT/image to PDF conversion
- `pypdf` or `PyPDF2` - PDF merge
- `pywin32` (Windows) - Word COM for Office conversion

#### LibreOffice (for Office → PDF conversion)

For best support of DOC/DOCX/ODT/RTF → PDF:
- **Linux**: `sudo apt-get install libreoffice`
- **macOS**: `brew install --cask libreoffice`
- **Windows**: Download from https://www.libreoffice.org/

### Usage

```bash
python DokuReader.py
```

Or via START.bat (Windows):
```bash
START.bat
```

### Data Storage

State is saved in: `~/.dokubibliothek_state.json`

### Supported File Formats

- Documents: `.txt`, `.doc`, `.docx`, `.pdf`, `.odt`, `.rtf`
- Images: `.jpg`, `.jpeg`, `.gif`, `.png`

### License

GPL v3 - See [LICENSE](LICENSE)
