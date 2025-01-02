End:: ExitApp
SetTitleMatchMode 2
WinMinimizeAll
Sleep 200
WinRestore "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
Sleep 100
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
Sleep 100
WinMove 800, 50 ; Normalisera plats för Bookit
AllaKundnummer := FileRead("kundnummer.txt") ;
SetKeyDelay 100

; Kolla om man kan minimera alla fönster utom de man är i??

; Starta Word
Run "Förlängningsbrev Reskort Nivå 4.lnk"

WinWait "Förlängningsbrev Reskort Nivå 4"
WinWait "Förlängningsbrev Reskort Nivå 4"
WinActivate "Förlängningsbrev Reskort Nivå 4"
WinMove 50, 50, 768, 1024 ; Normalisera storlek för Word
Sleep 200

WinActivate "Förlängningsbrev Reskort Nivå 4" ; Fokusera på Word
A_Clipboard := "" ; Empty the clipboard
send "^{a}"
Sleep 200
send "^{c}"
ClipWait 2
Sleep 200
loop parse, AllaKundnummer, "`n", "`r" {
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
loop parse, AllaKundnummer, "`n", "`r"  ; Loopa igenom kundnumren och kör detta en gång för varje kundnummer
{

    Sleep 50

    Kundnummer := FileRead("kundnummer.txt") ; läs innehållet i kundnummer.txt
    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
    Sleep 50

    send "{Esc}" ; ifall man är inne i nånting redan

    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
    send "{Esc}" ; ifall man är inne i nånting redan

    Send("+{Tab}") ; flytta musen till "filter"
    MouseClick
    Sleep 100

    send "^{a}"
    Sleep 100

    send "{Del}"
    Sleep 100

    Send "kunder"
    Sleep 100

    MouseMove 70, 140 ; flytta musen till "kunder"
    MouseClick
    Sleep 100

    MouseMove 250, 90 ; flytta musen till kundnr
    Sleep 100

    MouseClick
    Sleep 100
    Send A_LoopField
    send "{Home}"
    Sleep 100
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
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 250
    ClipWait 2
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
    } else {
        Sleep 50
        A_Clipboard := StrTitle(A_Clipboard)
    }

    ; Lägg in förnamn i Word
    sleep 50

    WinWait "Förlängningsbrev Reskort Nivå 4"
    WinActivate "Förlängningsbrev Reskort Nivå 4"

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
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 250
    ClipWait 2
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
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 250
    ClipWait 2
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
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 250
    ClipWait 2
    sleep 50

    ; Lägg in postnummer i Word
    WinActivate "Förlängningsbrev Reskort Nivå 4"
    send "^{v}"
    send " " ; mellanrum efter postnummer
    sleep 50

    ; Hitta postort
    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
    send "{Tab}"
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 250
    ClipWait 2
    sleep 50

    A_Clipboard := StrTitle(A_Clipboard)
    sleep 50

    ; Lägg in postort i Word
    WinActivate "Förlängningsbrev Reskort Nivå 4"
    send "^{v}"
    sleep 200

    ; Lägg in datum i Word
    sleep 200
    send "{F11}"
    sleep 200
    send "{F11}"
    sleep 300

    send "{Alt}{n}"
    sleep 200
    send "{d}{a}"
    sleep 200
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
