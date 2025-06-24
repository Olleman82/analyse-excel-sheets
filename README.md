# Analyse Excel Sheets

Detta projekt analyserar alla Excel-filer (.xlsx) i den aktuella mappen och genererar en rapport i Markdown-format (`Excel_Analysis_Report.md`). Rapporten dokumenterar varje kolumns namn, format och genererar representativ, anonymiserad fejkdata.

## Funktioner
- Automatisk analys av alla `.xlsx`-filer i mappen
- Identifierar kolumnnamn, datatyper och format
- Genererar fejkad, anonymiserad exempeldata för känsliga fält
- Skapar en överskådlig rapport i Markdown-format

## Installation
1. Klona detta repo:
   ```bash
   git clone https://github.com/<ditt-användarnamn>/analyse-excel-sheets.git
   cd analyse-excel-sheets
   ```
2. Installera beroenden:
   ```bash
   pip install -r requirements.txt
   ```

## Användning
Lägg dina `.xlsx`-filer i projektmappen och kör:
```bash
python analyze_excel.py
```
En rapport (`Excel_Analysis_Report.md`) skapas i samma mapp.

## Beroenden
- pandas
- openpyxl
- Faker
- numpy

## Licens
MIT
