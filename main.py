import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PyPDF2 import PdfReader
import openpyxl
from openpyxl import Workbook

EXCEL_FILE = "Formulardaten.xlsx"
PDF_DIR = "."

class PDFHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.headers_created = False
        self.init_excel()

    def init_excel(self):
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            ws = wb.active
            ws.title = "Formulardaten"
            wb.save(EXCEL_FILE)
            print("Neue Excel-Datei erstellt")
        else:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active
            if ws.max_row > 0:
                self.headers_created = True

    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return

        time.sleep(0.5)
        self.process_pdf(event.src_path)

    def process_pdf(self, pdf_path):
        try:
            print(f"\nVerarbeite Datei: {os.path.basename(pdf_path)}")
            
            # Formularfelder extrahieren
            fields = self.extract_form_fields(pdf_path)
            if not fields:
                print("Keine Formularfelder gefunden")
                return

            print("Gefundene Felder:", fields)

            # Excel-Datei bearbeiten
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active

            # Header erstellen (nur beim ersten Mal)
            if not self.headers_created:
                ws.append(list(fields.keys()))
                self.headers_created = True
                print("Header geschrieben")

            # Daten in korrekter Reihenfolge einfügen
            try:
                ordered_data = [str(fields[header.value]) for header in ws[1]]
                ws.append(ordered_data)
                print("Datenzeile hinzugefügt:", ordered_data)
            except KeyError as e:
                print(f"Fehler: Feld {e} nicht in Excel-Header gefunden")

            wb.save(EXCEL_FILE)
            print("Excel-Datei gespeichert")
            
            os.remove(pdf_path)
            print("PDF erfolgreich gelöscht")

        except PermissionError:
            print("Excel-Datei ist gesperrt - bitte schließen Sie die Datei")
        except Exception as e:
            print(f"Kritischer Fehler: {str(e)}")

    def extract_form_fields(self, pdf_path):
        try:
            with open(pdf_path, "rb") as f:
                reader = PdfReader(f)
                fields = reader.get_form_text_fields()
                if not fields:
                    return None
                
                return {
                    name: fields[name] or "" 
                    for name in fields
                }
        except Exception as e:
            print(f"PDF-Lesefehler: {str(e)}")
            return None

if __name__ == "__main__":
    event_handler = PDFHandler()
    observer = Observer()
    observer.schedule(event_handler, PDF_DIR)
    
    print(f"Überwache Ordner '{os.path.abspath(PDF_DIR)}'...")
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    
    observer.join()
