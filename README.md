# Mini Spreadsheet (Java)

Ein einfaches Tabellenkalkulations-Projekt in Java.  
Das Projekt implementiert eine **Spreadsheet-Klasse**, in der Zellen Werte oder Formeln enthalten können.  
Formeln unterstützen sowohl Grundrechenarten als auch einfache Funktionen (SUMME, MIN, MAX, MITTELWERT).

---

## Features

- **Zellen (Cell)**
  - Speichern Werte (`"123"`) oder Formeln (`"=A1+5"`).
  - Automatische Berechnung von Formeln.

- **Formeln**
  - Grundrechenarten: `+ - * / ^`
  - Zellreferenzen: `=A2+41`
  - Bereichsfunktionen:
    - `SUMME(A1:B3)`
    - `MIN(A1:A5)`
    - `MAX(A1:A5)`
    - `MITTELWERT(A1:A5)`

- **CSV I/O**
  - Speichern des gesamten Spreadsheets als CSV (`saveCsv`).
  - Laden aus CSV (`readCsv`), ab beliebiger Startzelle.

- **Limits**
  - Maximal 99 Zeilen (1–99).
  - Maximal 26 Spalten (A–Z).

---

## Projektstruktur


