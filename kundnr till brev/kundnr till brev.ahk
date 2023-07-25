End::ExitApp
SetTitleMatchMode 2
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
AllaKundnummer := FileRead("kundnummer.txt") ; 
Loop parse, AllaKundnummer, "`n", "`r"  
{
Sleep 50

; Starta Word
Run "Förlängningsbrev Reskort Nivå 4.lnk"
Sleep 500

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

SendInput "^c"
Sleep 250
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
WinMaximize "Förlängningsbrev Reskort Nivå 4"
sleep 50
send "{PgUp}"
sleep 50
send "{F11}"
sleep 50
send "^{v}"
sleep 50
send " "
sleep 50


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
send "{F11}"
send "{F11}"

MouseMove 150, 50 ; Tryck infoga
MouseClick
sleep 150

MouseMove 1660, 100 ; Tryck datum och tid
MouseClick
sleep 500
send "^{Enter}"
sleep 1000

; Printa
send "^{p}"
sleep 4000
MouseMove 220, 140 ; Tryck printa
MouseClick
sleep 6000

; Stäng word 
WinActivate "Förlängningsbrev Reskort Nivå 4"
MouseMove 1900, 10 ; Tryck kryss
MouseClick
sleep 6000

; WinActivate "Microsoft Word" Kommenterade eftersom vissa datorer inte hittar titel på fönstret. Den verkar autofokusa ändå
sleep 500
send "{Tab}" ; Tryck spara inte
send "{Enter}"
sleep 2000

; Gå ur kontot på Bookit
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
send "{Esc}"
send "{Esc}"

}