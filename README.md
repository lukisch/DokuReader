# DokuReader - Dokumentenbibliothek

Eine einfache Desktop-Anwendung zur Verwaltung und Organisation von Dokumenten nach Themen mit Vorschau-Funktion und PDF-Export.

## Features

- **Themen-Organisation**: Erstellen, umbenennen und löschen Sie Themen für Ihre Dokumente
- **Gelesen/Ungelesen**: Markieren Sie Dokumente als gelesen (grün mit ✓)
- **Vorschau-Funktion**:
  - Bilder (JPG, PNG, GIF)
  - PDF-Dokumente (erste Seite)
  - Textdateien (TXT)
  - Office-Dokumente (DOCX, ODT)
- **Drag & Drop**: Dateien einfach per Drag & Drop hinzufügen (optional)
- **Doppelklick-Öffnen**: Dokumente im Standard-Programm öffnen
- **Sammel-PDF-Export**: Exportieren Sie alle, gelesene oder ungelesene Dokumente als ein PDF
  - Unterstützt: PDF, TXT, Bilder, DOC, DOCX, ODT, RTF
  - Automatische Konvertierung zu PDF (LibreOffice oder MS Word)
- **Plattform-übergreifend**: Windows, macOS, Linux

## Technische Details

- Python 3.10+ mit Tkinter
- Single-File-Anwendung (756 Zeilen)
- JSON-basierte Persistenz im Home-Verzeichnis
- Nur Verweise: Originaldateien bleiben unberührt

## Installation

### Benötigte Dependencies

```bash
pip install -r requirements.txt
```

### Optionale Dependencies

Für volle Funktionalität:
- `pdf2image` oder `PyMuPDF` - PDF-Vorschau
- `tkinterdnd2` - Drag & Drop
- `python-docx` - DOCX-Vorschau
- `odfpy` - ODT-Vorschau
- `reportlab` - TXT/Bild zu PDF Konvertierung
- `pypdf` oder `PyPDF2` - PDF-Merge
- `pywin32` (Windows) - Word COM für Office-Konvertierung

### LibreOffice (für Office → PDF Konvertierung)

Für beste Unterstützung von DOC/DOCX/ODT/RTF → PDF:
- **Linux**: `sudo apt-get install libreoffice`
- **macOS**: `brew install --cask libreoffice`
- **Windows**: Download von https://www.libreoffice.org/

## Verwendung

```bash
python DokuReader.py
```

Oder per START.bat (Windows):
```bash
START.bat
```

## Dateiformat

State wird gespeichert in: `~/.dokubibliothek_state.json`

## Unterstützte Dateiformate

- Dokumente: `.txt`, `.doc`, `.docx`, `.pdf`, `.odt`, `.rtf`
- Bilder: `.jpg`, `.jpeg`, `.gif`, `.png`

## Lizenz

Proprietär - Lukas Geiger
