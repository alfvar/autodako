End::ExitApp
SetTitleMatchMode 2
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
AllaKundnummer := FileRead("kundnummer.txt") ; 

; Starta Word
Run "Förlängningsbrev Reskort Nivå 4.lnk"

WinWait "Förlängningsbrev Reskort Nivå 4"
WinWait "Förlängningsbrev Reskort Nivå 4"
WinActivate "Förlängningsbrev Reskort Nivå 4"
; Normalisera storlek för Word
WinMove 50, 50, 768, 1024
Sleep 200

send "^{a}"
Sleep 200
send "^{c}"
Sleep 200
Loop parse, AllaKundnummer, "`n", "`r"  
{
send "{Down}"
Sleep 100
send "^{Enter}"
Sleep 100
send "^{v}"
Sleep 100
}
Sleep 1000
send "^{Home}" ; Gå till början av dokumentet
Sleep 100
Loop parse, AllaKundnummer, "`n", "`r"  ; Loopa igenom kundnumren och kör detta en gång för varje kundnummer
{
Sleep 50

Kundnummer := FileRead("kundnummer.txt") ; läs innehållet i kundnummer.txt
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
Sleep 50

send "{Esc}"
Sleep 50

send "{Esc}"
Sleep 50

send "{Esc}"
Sleep 50

send "{Esc}" ; ifall man är inne i nånting redan
MouseMove 100, 220 ; flytta musen till "kunder"
MouseClick
Sleep 1500

MouseMove 250, 90 ; flytta musen till kundnr
Sleep 50
MouseClick
Send A_LoopField
send "{Home}"

MouseMove 300, 240 ; Öppna profilen
MouseClick
MouseClick
SetKeyDelay 25

; Hitta förnamn
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 250
MouseMove 250, 290
MouseClick
send "^{a}"
sleep 50

; Kopiera förnamn
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 50

send "^c"
Sleep 250
A_Clipboard := StrTitle(A_Clipboard)
Sleep 50

if (InStr(A_Clipboard, "-")) { ; Gör så att det är Stor Bokstav i början av dubbelnamn
    parts := StrSplit(A_Clipboard, "-")
    modifiedClipboard := ""
    for index, part in parts {
        modifiedPart := StrTitle(part)
        modifiedClipboard .= modifiedPart . "-"
    }
    modifiedClipboard := RTrim(modifiedClipboard, "-")
    A_Clipboard := modifiedClipboard
}


; Lägg in förnamn i Word
sleep 50

WinWait "Förlängningsbrev Reskort Nivå 4"
WinActivate "Förlängningsbrev Reskort Nivå 4"
; Normalisera storlek för Word
WinMove 50, 50, 768, 1024

sleep 10
send "{F11}"
sleep 10
send "^{v}"
sleep 10
send " "
sleep 10


; Hitta efternamn
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 50

send "{Tab}"
sleep 50

send "^{a}"
sleep 50

; Kopiera efternamn 
send "^{c}"
sleep 50
A_Clipboard := StrTitle(A_Clipboard)

if (InStr(A_Clipboard, "-")) { ; Gör så att det är Stor Bokstav i början av dubbelefternamn
    parts := StrSplit(A_Clipboard, "-")
    modifiedClipboard := ""
    for index, part in parts {
        modifiedPart := StrTitle(part)
        modifiedClipboard .= modifiedPart . "-"
    }
    modifiedClipboard := RTrim(modifiedClipboard, "-")
    A_Clipboard := modifiedClipboard
}
sleep 50

; Lägg in efternamn i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
sleep 50

send "^{v}"
sleep 50
send "{F11}"
sleep 50


; Hitta adress 1
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 50

send "{Tab}"
sleep 50

send "^{a}"
; Kopiera adress 1
send "^{c}"
sleep 50

A_Clipboard := StrTitle(A_Clipboard) ; gör så att "lgh" alltid har små bokstäver
sleep 50
if (InStr(A_Clipboard, "lgh", false, 1)) {
    A_Clipboard := StrReplace(A_Clipboard, "lgh", "lgh")
}

sleep 50


; Lägg in adress 1 i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
send "^{v}"
send "{F11}"
sleep 50

; Hitta postnummer
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Tab}"
send "^{a}"

; Kopiera postnummer
send "^{c}"
sleep 50

; Lägg in postnummer i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
send "^{v}"
send " " ; mellanrum efter postnummer
sleep 50

; Hitta postort
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Tab}"
send "^{c}" ; Kopiera postort
sleep 50

A_Clipboard := StrTitle(A_Clipboard)
sleep 50


; Lägg in postort i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
send "^{v}"


; Lägg in datum i Word
sleep 50
send "{F11}"
sleep 50
send "{F11}"
sleep 50

send "{Alt}{n}"
send "{d}{a}"
send "{Enter}"
sleep 200

send "{F5}"
sleep 200
send "{Raw}+1"
sleep 200
send "{Enter}"
sleep 200
send "{Esc}"



; Gå ur kontot på Bookit
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Esc}"
send "{Esc}"

}