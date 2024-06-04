End::ExitApp
HotIfWinActive "Registerunderhåll - http://prod.gonet.se/bookit/"
CoordMode "Mouse", "Window"
Hotkey "^+r", printRK
;Hotkey "^+p", printPK


SetTitleMatchMode 2

PrintRK(ThisHotkey)
{
SetKeyDelay 25

; Starta Word
Run "Förlängningsbrev Reskort Nivå 4.lnk"

WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
; Normalisera storlek för Bookit
WinMove 800, 50, 981, 775
MouseMove 200, 270 ; dubbelkolla att dessa pixlar är kundnr (så att man inte sitter i nån annan ruta)
MouseClick


; Hitta förnamn
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 1000
send "{Tab}"
send "{Tab}"
send "^{a}"
sleep 100

; Kopiera förnamn
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 1000

send "^c"
Sleep 10
A_Clipboard := StrTitle(A_Clipboard)
Sleep 50

if (InStr(A_Clipboard, "-")) { 
    parts := StrSplit(A_Clipboard, "-")
    modifiedClipboard := ""
    for index, part in parts {
        modifiedPart := StrTitle(part)
        modifiedClipboard .= modifiedPart . "-"
    }
    modifiedClipboard := RTrim(modifiedClipboard, "-")
    A_Clipboard := modifiedClipboard
}
; Gör så att det är Stor Bokstav i början av dubbelnamn

; Lägg in förnamn i Word
WinWait "Förlängningsbrev Reskort Nivå 4"
WinActivate "Förlängningsbrev Reskort Nivå 4"
; Normalisera storlek för Word
WinMove 50, 50, 768, 1024

send "{PgUp}"
sleep 10
send "{F11}"
sleep 10
send "^{v}"
sleep 10
send " "
sleep 10


; Hitta efternamn
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 10

send "{Tab}"
sleep 10

send "^{a}"
sleep 10

; Kopiera efternamn 
send "^{c}"
sleep 10
A_Clipboard := StrTitle(A_Clipboard)
sleep 10

; Lägg in efternamn i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
sleep 10

send "^{v}"
sleep 10
send "{F11}"
sleep 10


; Hitta adress 1
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
sleep 10

send "{Tab}"
sleep 10

send "^{a}"
; Kopiera adress 1
send "^{c}"
sleep 10

if (InStr(A_Clipboard, "-")) { 
    parts := StrSplit(A_Clipboard, "-")
    modifiedClipboard := ""
    for index, part in parts {
        modifiedPart := StrTitle(part)
        modifiedClipboard .= modifiedPart . "-"
    }
    modifiedClipboard := RTrim(modifiedClipboard, "-")
    A_Clipboard := modifiedClipboard
}
; Gör så att det är Stor Bokstav i början av dubbla postadresser

A_Clipboard := StrTitle(A_Clipboard) ; gör så att "lgh" alltid har små bokstäver
sleep 10
if (InStr(A_Clipboard, "lgh", false, 1)) {
    A_Clipboard := StrReplace(A_Clipboard, "lgh", "lgh")
}

sleep 50

; Lägg in adress 1 i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
send "^{v}"
send "{F11}"
sleep 10

; Hitta postnummer
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Tab}"
send "^{a}"

; Kopiera postnummer
send "^{c}"
sleep 10

; Lägg in postnummer i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
send "^{v}"
send " " ; mellanrum efter postnummer
sleep 10

; Hitta postort
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Tab}"
sleep 50
send "^{a}"
send "^{c}" ; Kopiera postort
sleep 10
A_Clipboard := StrTitle(A_Clipboard) 
if (InStr(A_Clipboard, "-")) { 
    parts := StrSplit(A_Clipboard, "-")
    modifiedClipboard := ""
    for index, part in parts {
        modifiedPart := StrTitle(part)
        modifiedClipboard .= modifiedPart . "-"
    }
    modifiedClipboard := RTrim(modifiedClipboard, "-")
    A_Clipboard := modifiedClipboard
}
; Gör så att det är Stor Bokstav i början av dubbla postorter

; Lägg in postort i Word
WinActivate "Förlängningsbrev Reskort Nivå 4"
send "^{v}"
sleep 200
send "{F11}"
send "{F11}"

; Lägg in datum i Word
send "{Alt}{n}"
send "{d}{a}"
send "{Enter}"
sleep 200

; Printa
send "^{p}"
sleep 200
send "{Enter}"
sleep 2000


; Gå ur kontot på Bookit
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Esc}"
send "{Esc}"
sleep 200

}
