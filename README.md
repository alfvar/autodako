# Autodako

Automatiseringsverktyg för brevutskicken i den dagliga dako-körningen på kundservice. Körs genom programmet [Autohotkey v2](https://github.com/AutoHotkey/AutoHotkey). Programmet behöver installeras på varje maskin som ska köra autodako-skript.
## extract-adress.ahk

Som det låter, tar textfilen `kundnummer.txt` och skriver ut kundprofilens detaljer i en `kunddata.csv`. Bring your own `kundnummer.txt`.

## kunddata_till_brev.py
Genererar en `output.docx` där varje sida är ett förlängningsbrev för reskort nivå 4. Använder `kunddata.csv` som input.
