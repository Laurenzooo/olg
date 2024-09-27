# README
Ein mit ChatGPT zusammengebasteltes Python Script für die Onlogist-Buchhaltung. Achtung: Ich verstehe so viel vom Programmieren wie ein 4. Klässler. Ein echter Programmierer würde beim Anblick des Codes einen Kollaps bekommen.
Wenn man nichts mit Python am Hut hat:
1. VS-Code installieren
2. Python-Extension in VS-Code installieren


## Tutorial

### Step 0:
Erstelle einen Ordner in dem du das Script "OLG_skript.py" speicherst. Erstelle in diesem Ordner außerdem die drei Ordner: "Protokolle", "Rechnungen", "Tabellen".

### Step 1:
Für jeden Auftrag muss es einen Ordner geben, in dem eine Datei "Übernahme-Protokoll.pdf" und eine Datei "Abgabe-Protokoll.pdf" enthalten ist. Diese 'Auftrags-Ordner' müssen alle im Ordner "Protokolle" gespeichert sein. Die Ordner können auch andere Dateien enthalten - es werden nur die beiden Protokoll-Dateien beachtet.

### Step 2:
Die excel-Dateien für jede Rechnung von Onlogist runterladen und in einem Ordner speichern. Das geht über die Website "https://portal.onlogist.com/invoices", in der App ist das nicht möglich. Man muss das Download-Icon rechts neben dem Drucken-Icon klicken. Diese Dateien alle im Ordner "Rechnungen" speichern.

### Step 3:
Die Pfade zu den Ordnern im Skript anpassen. Dazu das Skript öffnen und und die Pfade in den Zeilen 8, 9 und 10 mit den tatsächlichen Pfaden zu den Ordnern "Protokolle", "Rechnungen" und "Tabellen" ersetzen. Wenn nicht gewusst wie: -> Google

### Step 3.1:
Falls deine Rechnungsnummern nicht aus 3 Ziffern bestehen, musst du die richtige Anzahl der Ziffern in Zeile 13 anpassen.

### Step 4:
Im Terminal
```
pip install pdfplumber pandas
```
ausführen

### Step 5:
Run the script. Wenn alles richtig gemacht wurde, sollten im Ordner "Tabellen" vier Tabellen sein. In der Datei "Mega-Tabelle" sind jetzt zu jeden Auftrag die wichtigsten Daten zusammengetragen.

Die Tabelle kann man dann mit Excel oder Google-Sheets noch hübsch machen und einfach Formeln für Provision und Versicherung (mithilfe der Distanz, also der Differenz zwischen km[Übernahme] und km[Abgabe]) hinzufügen. Oder man trägt die Daten manuell ein, wenn einem das sonst zu ungenau ist. Oder man erstellt noch ein Skript, dass einem die genauen Summen aus den Versicherungs- und Provisions- Rechnungen extrahiert. Dazu war ich allerdings bisher zu faul. Dann kann man sich relativ einfach den Netto-Gewinn ausspucken lassen und auch alle möglichen statistischen Spielchen machen.
Es sind auch ein paar unnötige Spalten dabei, die man ggfs. löschen kann. Die Spalte Start/Ziel funktioniert auch nicht so wirklich, weil die Protokolle bescheuert formatiert sind und ich nicht programmieren kann.
