#!/usr/bin/env python
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PyPDF2 import PdfReader
import openpyxl
from openpyxl import Workbook

# Konfiguration
EXCEL_FILE = "Formulardaten.xlsx"
PDF_DIR = "."  # Aktuelles Verzeichnis


class PDFHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.ensure_excel_file()

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".pdf"):
            time.sleep(0.5)  # Warte bis die Datei vollständig geschrieben ist
            self.process_pdf(event.src_path)

    def ensure_excel_file(self):
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            wb.save(EXCEL_FILE)

    def process_pdf(self, pdf_path):
        try:
            # Extrahiere Formularfelder
            fields = self.extract_form_fields(pdf_path)
            if not fields:
                return

            # Lese Excel-Datei
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active

            # Erstelle Header beim ersten Durchgang
            if ws.max_row == 1 and not any(ws.iter_rows(max_row=1)):
                ws.append(list(fields.keys()))

            # Füge Daten in der richtigen Reihenfolge ein
            ordered_data = [fields[header.value] for header in ws[1]]
            ws.append(ordered_data)

            wb.save(EXCEL_FILE)
            os.remove(pdf_path)
            print(f"Verarbeitet: {os.path.basename(pdf_path)}")

        except Exception as e:
            print(f"Fehler bei {pdf_path}: {str(e)}")

    def extract_form_fields(self, pdf_path):
        try:
            reader = PdfReader(pdf_path)
            fields = reader.get_fields()
            return {
                name: field.value if field.value else ""
                for name, field in fields.items()
            }
        except Exception as e:
            print(f"PDF-Fehler: {str(e)}")
            return None


if __name__ == "__main__":
    event_handler = PDFHandler()
    observer = Observer()
    observer.schedule(event_handler, path=PDF_DIR, recursive=False)
    observer.start()

    print("Überwache Ordner... (Strg+C zum Beenden)")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
