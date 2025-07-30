End:: ExitApp
SetTitleMatchMode 2
; WinMinimizeAll
Sleep 400
WinRestore "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
Sleep 400
WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
WinWait "Registerunderhåll" ; 
Sleep 100
WinMove 800, 50 ; Normalisera plats för Bookit
AllaKundnummer := FileRead("kundnummer.txt") ;
kundadresser := A_WorkingDir "\kundadresser.xlsx"
SetKeyDelay 100

; Kolla om man kan minimera alla fönster utom de man är i??


; Create and save an Excel file "kundadresser.xlsx"
FileDelete kundadresser  ; Delete the file if it exists

xl := ComObject("Excel.Application")
xl.Visible := true
wb := xl.Workbooks.Add()
ws := wb.Worksheets(1)

; Set column headers
ws.Range("A1").Value := "name"
ws.Range("B1").Value := "address 1"
ws.Range("C1").Value := "address 2"

; Save the workbook in the script's directory
wb.SaveAs(A_WorkingDir "\kundadresser.xlsx")

; Close the workbook after creating it
wb.Close()

; Reopen the workbook for editing (keep open during loop)
wb := xl.Workbooks.Open(A_WorkingDir "\kundadresser.xlsx")
ws := wb.Worksheets(1)

Sleep 1000
send "^{Home}" ; Gå till början av dokumentet
Sleep 100
loop parse, AllaKundnummer, "`n", "`r"  ; Loopa igenom kundnumren och kör detta en gång för varje kundnummer
{

    Sleep 50

    Kundnummer := FileRead("kundnummer.txt") ; läs innehållet i kundnummer.txt
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    Sleep 50

    send "{Esc}" ; ifall man är inne i nånting redan

    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    send "{Esc}" ; ifall man är inne i nånting redan

    Send("+{Tab}") ; flytta musen till "filter"
    Sleep 50

    send "^{a}"
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    Sleep 50

    Send "kunder"
    Sleep 50

    Send("+{Tab}") ; flytta musen till kundnr
    Sleep 50
    Send "{Down}"    
    Sleep 50

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    Send "{Enter}"
    Sleep 50
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    MouseClick
    Sleep 50
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret

    Send A_LoopField
    send "{Home}"
    Sleep 50
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret

    MouseMove 300, 240 ; Öppna profilen
    MouseClick
    MouseClick
    SetKeyDelay 25

    ; Hitta förnamn
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 250
    MouseMove 250, 290
    MouseClick
    send "^{a}"
    sleep 50

    ; Kopiera förnamn
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
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

    if (InStr(A_Clipboard, " von ", false, 1)) {
    A_Clipboard := StrReplace(A_Clipboard, " von ", " von ")
}

    if (InStr(A_Clipboard, " af ", false, 1)) {
    A_Clipboard := StrReplace(A_Clipboard, " af ", " af ")
}

    if (InStr(A_Clipboard, " van ", false, 1)) {
    A_Clipboard := StrReplace(A_Clipboard, " van ", " van ")
}

    ; Find the first empty row in column A
    lastRow := ws.Cells(ws.Rows.Count, 1).End(-4162).Row ; -4162 = xlUp
    nextRow := lastRow + 1
    if (ws.Range("A1").Value = "")  ; If the sheet is empty
        nextRow := 1
    cellAddress := "A" nextRow  ; Store the cell address in a variable

    ; Write clipboard contents to the first empty cell in column A
    ws.Range(cellAddress).Value := A_Clipboard
    wb.Save()

 ; Hitta efternamn
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    Send("{Tab}")
    send "^{a}"
    sleep 50

    ; Kopiera efternamn
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 50
    ClipWait 2
    A_Clipboard := StrTitle(A_Clipboard)
    Sleep 50

    if (InStr(A_Clipboard, "-")) { ; Gör så att det är Stor Bokstav i början av dubbelefternamn
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


    ; Lägg till efternamn i namn
    currentValue := ws.Range(cellAddress).Value ; Lägg till efternamn i namn
    if (currentValue != "")
        ws.Range(cellAddress).Value := currentValue " " A_Clipboard
    else
        ws.Range(cellAddress).Value := A_Clipboard

    wb.Save()

 ; Hitta adress 1
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    Send("{Tab}")
    send "^{a}"
    sleep 50

    ; Kopiera adress 1
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 50
    ClipWait 2
    A_Clipboard := StrTitle(A_Clipboard)
    Sleep 50

    A_Clipboard := StrTitle(A_Clipboard) ; gör så att "lgh" alltid har små bokstäver
    sleep 50
    if (InStr(A_Clipboard, "lgh", false, 1)) {
        A_Clipboard := StrReplace(A_Clipboard, "lgh", "lgh")
    }
    
    cellAddress := "B" nextRow  ; Store the cell address in a variable

    ; Write clipboard contents to the first empty cell in column B
    ws.Range(cellAddress).Value := A_Clipboard
    wb.Save()

 ; Hitta postnummer
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    Send("{Tab}")
    send "^{a}"
    sleep 50

    ; Kopiera postnummer
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 50
    ClipWait 2
    
    ; Remove all spaces from clipboard content
    A_Clipboard := StrReplace(A_Clipboard, " ", "")
    
    ; Insert space after third character if length is greater than 3
    if (StrLen(A_Clipboard) > 3) {
        A_Clipboard := SubStr(A_Clipboard, 1, 3) " " SubStr(A_Clipboard, 4)
    }
    
    Sleep 50

    cellAddress := "C" nextRow  ; Store the cell address in a variable

    ; Write clipboard contents to the first empty cell in column C
    ws.Range(cellAddress).Value := A_Clipboard
    wb.Save()

 ; Hitta efternamn
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    Send("{Tab}")
    send "^{a}"
    sleep 50

    ; Kopiera postort
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    sleep 50
    A_Clipboard := "" ; Empty the clipboard
    sleep 50
    send "^{c}"
    Sleep 50
    ClipWait 2
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
        Sleep 50


    ; Lägg till postort i adress 2
    currentValue := ws.Range(cellAddress).Value ; Lägg till postort i namn
    if (currentValue != "")
        ws.Range(cellAddress).Value := currentValue " " A_Clipboard
    else
        ws.Range(cellAddress).Value := A_Clipboard

    wb.Save()

}
; Exit the profile
WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    Sleep 50

    send "{Esc}" ; ifall man är inne i nånting redan

    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100

    if WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
    loop {
        if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
    }

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    send "{Esc}" ; ifall man är inne i nånting redan

    Send("+{Tab}") ; flytta musen till "filter"
    Sleep 50

    send "^{a}"
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    Sleep 50

    Send "kunder"
    Sleep 50

    Send("+{Tab}") ; flytta musen till kundnr
    Sleep 50
    Send "{Down}"    
    Sleep 50

    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    Send "{Enter}"
    Sleep 50
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret
    MouseClick
    Sleep 50
    WinActivate "Registerunderhåll" ; Fokusera på Registerunderhåll-fönstret


xl := ""        ; Release the COM object