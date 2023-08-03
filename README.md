# Autodako

Automatiseringsverktyg för den dagliga dako-körningen på kundservice. Körs genom programmet [Autohotkey v2](https://github.com/AutoHotkey/AutoHotkey). Programmet behöver installeras på varje maskin som ska köra autodako-skript.
## Kundnr till brev.ahk

Som det låter, tar textfilen `kundnummer.txt` och skriver ut ett Reskort Nivå 4-välkomstbrev för varje kundnummer.

## Hotkey-kundnr till brev.ahk
När man kör scriptet väntar den på kortkommandot `Ctrl + Shfit + R`. När kommandot körs förväntar sig programmet ett öppet konto (sidan där man redigerar postadress för privatanvändare under registerunderhåll) och skriver sedan ut ett Reskort Nivå 4-välkomstbrev för den aktuella profilen.

## Namn till kort.ahk
Körs på kortskrivardatorn. Skriver ut ett reskort nivå 4 för varje rad i textfilen `namnlista.txt`. Förväntar sig formatet `2012123456,Test Testsson,2023-07-01` i den ordningen, separerat med kommatecken. Ny rad blir ett nytt kort.
