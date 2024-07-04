End::ExitApp
SetTitleMatchMode 2
csvFilePath := "kunddata.csv"
AllaKundnummer := FileRead("kundnummer.txt") ; 

WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
FileObject := FileOpen(csvFilePath, "w") ; Öppna filen i skrivläge, vilket rensar befintligt innehåll
FileObject.Close() ; Stäng filen

Loop parse, AllaKundnummer, "`n", "`r"  ; Loopa igenom kundnumren och kör detta en gång för varje kundnummer
{
Sleep 50

Kundnummer := FileRead("kundnummer.txt") ; läs innehållet i kundnummer.txt
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
Sleep 50

send "{Esc}" ; ifall man är inne i nånting redan
Sleep 50

send "{Esc}" ; ifall man är inne i nånting redan
Sleep 50

send "{Esc}" ; ifall man är inne i nånting redan
Sleep 50

send "{Esc}" ; ifall man är inne i nånting redan
MouseMove 100, 220 ; flytta musen till "kunder"
MouseClick
Sleep 250

WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
MouseMove 250, 90 ; flytta musen till kundnr
Sleep 50
MouseClick
Send A_LoopField
send "{Home}"
Sleep 50
send "^{c}"


; Gå ur kontot på Bookit
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Esc}"
send "{Esc}"

FileAppend(A_Clipboard . "`n", csvFilePath)


}