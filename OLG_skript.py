import pdfplumber
import re
from pathlib import Path
import pandas as pd
import os

# HIER DIE DATEIPFADE ZU DEN ORDNERN ERSETZEN
pfad_protokolle = Path("/mein/ordner/Protokolle")
pfad_rechnungen = Path("/mein/ordner/Rechnungen")
pfad_tabellen = Path("/mein/ordner/Tabellen")

# HIER ANZAHL DER ZIFFERN DER RECHNUNGSNUMMER EINTRAGEN
ziffern_rechnungsnummer = 3


# Schlüsselwörter für die Datenpunkte
keywords = {
    "Auftragsnummer": "Startort",
    "Auftrags-Nummer": "Startort",
    "Startort": "Zielort",
    "Amtliches Kennzeichen": "Hersteller",
    "Datum/Uhrzeit": "Kilometerstand",
    "Datum und Uhrzeit der Übernahme ": "Kilometerstand",
    "Kilometerstand abgelesen bei": "Übernahme"
}

def extract_data_from_pdf(pdf_path, keywords):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        
        # Kombiniere alle Zeilen zu einem einzigen Textblock
        combined_text = " ".join(text.split('\n'))
        
        extracted_data = {}
        
        for key, end_key in keywords.items():
            if key in combined_text:
                start_index = combined_text.find(key) + len(key)
                
                if end_key:
                    end_index = combined_text.find(end_key, start_index)
                    if end_index == -1:
                        end_index = None  # Falls das End-Schlüsselwort nicht gefunden wird, bis zum Ende des Textes
                else:
                    end_index = None  # Keine End-Beschränkung, wenn kein End-Schlüsselwort vorhanden
                
                # Extrahiere den Text zwischen Schlüsselwörtern
                value = combined_text[start_index:end_index].strip()
                
                # Bearbeite den Wert für "Auftragsnummer"
                if key == "Auftrags-Nummer":
                    value = re.sub(r'\D', '', value)[:7]

                # Bearbeite den Wert für "Auftragsnummer"
                if key == "Auftragsnummer":
                    value = re.sub(r'\D', '', value)[:7]
                
                extracted_data[key] = value

    return extracted_data

def process_pdfs_in_directory(base_directory, keywords):
    base_path = Path(base_directory)
    data_list = []
    
    # Suche nach allen "Übernahme-Protokoll.pdf" Dateien in den Unterordnern
    for pdf_file in base_path.rglob('Übernahme-Protokoll.pdf'):
        print(f"Verarbeite Datei: {pdf_file}")
        extracted_data = extract_data_from_pdf(pdf_file, keywords)

        # Keys umbenennen
        rename_keys = {
            "Auftragsnummer": "Fahrt ID-#",
            "Auftrags-Nummer": "Fahrt ID-#",
            "Startort": "Start/Ziel",
            "Amtliches Kennzeichen": "Kennzeichen",
            "Datum/Uhrzeit": "Übernahme",
            "Datum und Uhrzeit der Übernahme ": "Übernahme",
            "Kilometerstand abgelesen bei": "Kilometerstand Übernahme"
        }

        # Umbenennen der Keys im Dictionary
        renamed_data = {rename_keys.get(k, k): v for k, v in extracted_data.items()}

        # Füge die Daten der Liste hinzu
        data_list.append(renamed_data)

    # Konvertiere die Liste in ein DataFrame
    df = pd.DataFrame(data_list)
    
    # Speichere das DataFrame als Excel-Datei
    # Pfad ersetzen durch den Pfad zu einem Ordner mit dem Namen "Tabellen"
    output_path = pfad_tabellen / "Übernahme_Protokolle_Tabelle.xlsx"
    df.to_excel(output_path, index=False)

    print(f"Übernahme-Protokolle gescannt.")
    print(f"Die Daten wurden in {output_path} gespeichert.")


# Setze den Pfad zu deinem Hauptordner, der die Ordner mit den Protkollen für die Aufträge enthält, hier ein.
base_directory = pfad_protokolle
process_pdfs_in_directory(base_directory, keywords)




# Schlüsselwörter für die Datenpunkte
keywords = {
    "Auftragsnummer": "Startort",
    "Auftrags-Nummer": "Startort",
    #"Startort": "Zielort",
    #"Amtliches Kennzeichen": "Hersteller",
    "Datum/Uhrzeit": "Kilometerstand",
    "Datum und Uhrzeit der Abgabe ": "Kilometerstand",
    "Übernahme Kilometerstand abgelesen bei": "km Abgabe",
    ":00 Kilometerstand abgelesen bei": "Abgabe Tankanzeige"
}

def extract_data_from_pdf(pdf_path, keywords):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        
        # Kombiniere alle Zeilen zu einem einzigen Textblock
        combined_text = " ".join(text.split('\n'))
        
        extracted_data = {}
        
        for key, end_key in keywords.items():
            if key in combined_text:
                start_index = combined_text.find(key) + len(key)
                
                if end_key:
                    end_index = combined_text.find(end_key, start_index)
                    if end_index == -1:
                        end_index = None  # Falls das End-Schlüsselwort nicht gefunden wird, bis zum Ende des Textes
                else:
                    end_index = None  # Keine End-Beschränkung, wenn kein End-Schlüsselwort vorhanden
                
                # Extrahiere den Text zwischen Schlüsselwörtern
                value = combined_text[start_index:end_index].strip()
                
                # Bearbeite den Wert für "Auftragsnummer"
                if key == "Auftrags-Nummer":
                    value = re.sub(r'\D', '', value)[:7]

                # Bearbeite den Wert für "Auftragsnummer"
                if key == "Auftragsnummer":
                    value = re.sub(r'\D', '', value)[:7]

                if key == ":00 Kilometerstand abgelesen bei":
                # Suche nach allen Zahlen im Wert und nimm die letzte gefundene Zahl
                    numbers = re.findall(r'\d+', value)
                    if numbers:
                        value = numbers[-1]

                
                extracted_data[key] = value

    return extracted_data

def process_pdfs_in_directory(base_directory, keywords):
    base_path = Path(base_directory)
    data_list = []
    
    # Suche nach allen "Übernahme-Protokoll.pdf" Dateien in den Unterordnern
    for pdf_file in base_path.rglob('Abgabe-Protokoll.pdf'):
        print(f"Verarbeite Datei: {pdf_file}")
        extracted_data = extract_data_from_pdf(pdf_file, keywords)

        # Keys umbenennen
        rename_keys = {
            "Auftragsnummer": "Fahrt ID-#",
            "Auftrags-Nummer": "Fahrt ID-#",
            #"Startort": "Start/Ziel",
            #"Amtliches Kennzeichen": "Kennzeichen",
            "Datum/Uhrzeit": "Abgabe",
            "Datum und Uhrzeit der Abgabe ": "Abgabe",
            "Übernahme Kilometerstand abgelesen bei": "Kilometerstand Abgabe",
            ":00 Kilometerstand abgelesen bei": "Kilometerstand Abgabe"
        }

        # Umbenennen der Keys im Dictionary
        renamed_data = {rename_keys.get(k, k): v for k, v in extracted_data.items()}

        # Füge die Daten der Liste hinzu
        data_list.append(renamed_data)

    # Konvertiere die Liste in ein DataFrame
    df = pd.DataFrame(data_list)
    
   # Pfad ersetzen durch den Pfad zum gleichen Ordner mit dem Namen "Tabellen"
    output_path = pfad_tabellen / "Abgabe_Protokolle_Tabelle.xlsx"
    df.to_excel(output_path, index=False)

    print(f"Abgabe-Protokolle gescannt.")
    print(f"Die Daten wurden in {output_path} gespeichert.")

# Setze den Pfad zu deinem Hauptordner, der die Ordner mit den Protkollen für die Aufträge enthält, hier noch einmal ein.
base_directory = pfad_protokolle
process_pdfs_in_directory(base_directory, keywords)





# Pfad zu dem Ordner, in dem sich die Rechnungs-.xls-Dateien befinden einfügen
folder_path = pfad_rechnungen

# Leere Liste zum Speichern der einzelnen DataFrames
xls_list = []

# Durchlaufe alle Dateien im Ordner
for filename in os.listdir(folder_path):
    if filename.endswith(".xls"):  # Überprüfen, ob es sich um eine .xls-Datei handelt
        file_path = os.path.join(folder_path, filename)

        # .xls-Datei in einen DataFrame laden
        df = pd.read_excel(file_path, engine='xlrd')

        # Extrahiere die letzten drei Ziffern als Rechnungsnummer aus dem Dateinamen
        match = re.search(r'(\d{' + str(ziffern_rechnungsnummer) + r'})\.xls$', filename)
        if match:
            rechnungsnummer = match.group(1)  # Extrahiere die gefundenen 3 Ziffern
        else:
            rechnungsnummer = 'Unbekannt'  # Falls keine Rechnungsnummer gefunden wird

        # Füge eine neue Spalte "Rechnungsnummer" zum DataFrame hinzu
        df['Rechnungsnummer'] = rechnungsnummer

        # DataFrame zur Liste hinzufügen
        xls_list.append(df)

# Alle DataFrames in der Liste zu einem großen DataFrame zusammenführen
combined_df = pd.concat(xls_list, ignore_index=True)

# Pfad ersetzen durch den Pfad zum gleichen Ordner mit dem Namen "Tabellen"
output_path = pfad_tabellen / "AlleRechnungen_Tabelle.xlsx"
# Das zusammengeführte DataFrame als neue Excel-Datei oder CSV-Datei speichern
combined_df.to_excel(output_path, index=False)
# Alternativ als CSV-Datei speichern:
# combined_df.to_csv('Pfad/zur/zusammengeführten_datei.csv', index=False)

print(f"Rechnungs-Dateien kombiniert.")






# Pfad zu dem Ordner, der die Excel-Dateien enthält
ordner_pfad = pfad_tabellen  # Ersetze durch den tatsächlichen Ordnerpfad

# Die Basisdatei festlegen (z.B. die erste Datei)
basis_datei = 'AlleRechnungen_Tabelle.xlsx'  # Ersetze durch den tatsächlichen Dateinamen der Basistabelle

# Alle Excel-Dateien im Ordner auflisten
excel_dateien = [file for file in os.listdir(ordner_pfad) if file.endswith('.xlsx') and file != basis_datei]

# Basis DataFrame einlesen
basis_df = pd.read_excel(os.path.join(ordner_pfad, basis_datei))

# Weitere Dateien einlesen und zur Basis hinzufügen
for datei in excel_dateien:
    datei_pfad = os.path.join(ordner_pfad, datei)
    df = pd.read_excel(datei_pfad)
    # Outer Join, um sicherzustellen, dass keine Spalten verloren gehen
    basis_df = pd.merge(basis_df, df, on='Fahrt ID-#', how='outer')



# Abrufen der vorhandenen Spaltennamen
print(basis_df.columns)
# Beispiel für die neue Reihenfolge, die nur die wichtigsten Spalten ändert
neue_reihenfolge = ["Rechnungsnummer", "Fahrt ID-#", "Unit-ID", "Kostenstelle", "Interne Nr.", "Start/Ziel", "Transportart", "Kennzeichen", "Fahrzeugtyp", "zgh. Hin-/Rückfahrt ID-#", "Übernahme", "Kilometerstand Übernahme", "Abgabe", "Kilometerstand Abgabe", "Nettobetrag", "Mwst", "Auslagen", "Titel", "Kostenart", "Bruttobetrag"]

# Alle anderen Spalten unverändert hinten anfügen
restliche_spalten = [spalte for spalte in basis_df.columns if spalte not in neue_reihenfolge]

# Kombinierte Reihenfolge
endgueltige_reihenfolge = neue_reihenfolge + restliche_spalten

# DataFrame in der neuen Reihenfolge anordnen
basis_df_neu = basis_df[endgueltige_reihenfolge]


# Das Ergebnis in eine neue Excel-Datei speichern
basis_df_neu.to_excel(os.path.join(ordner_pfad, 'Mega_Tabelle.xlsx'), index=False)

print("Mega-Tabelle erstellt!")