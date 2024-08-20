#Requires AutoHotkey v2
;@Ahk2Exe-SetMainIcon "edvin.ico"

End::ExitApp
SetTitleMatchMode 2
GroupAdd "Checkin", "Check-in - http://prod.gonet.se/bookit/"
GroupAdd "Checkin", "Bokning - http://prod.gonet.se/bookit/"
#HotIf WinActive("ahk_group Checkin")
CoordMode "Mouse", "Window"
SetKeyDelay 100

F7::dakob() ; Flytta musen till rutan för DAKOB
dakob() {
    WinActivate "ahk_group Checkin"    
    MouseMove 650, 160
    send "{F7}"
    Sleep 100
    MouseClick
    
}
; Disable the F6 key ; Så man slipper tappa sin sökning genom att göra en ny bokning av misstag
F6::return

^p::findCard() ; Tar kundnumret från aktuell checkinruta och tar upp kortsidan på dennes profil
findCard() {

    WinActivate "ahk_group Checkin"


    MouseMove 620, 450
    Sleep 100

    MouseGetPos , , &id, &control
    Kundnr := ControlGetText(control, "ahk_group Checkin")
    if !RegExMatch(Kundnr, "^\d+$")
    {
        Throw ValueError("Scriptet kördes inte inuti en bokning")
    }
    A_Clipboard := Kundnr
    Sleep 100

    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
    Sleep 10

    send "{Esc}" ; ifall man är inne i nånting redan

    Sleep 100


    if  WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
        Loop  {
            if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
        }

    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100


    if  WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
        Loop  {
            if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
        }

    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret

    send "{Esc}"
    Sleep 100


    if  WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
        WinActivate "Ändringar har gjorts"
        Loop  {
            if not WinExist("Ändringar har gjorts") ; Kolla efter popupfönster
            break
        }

    WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
    send "{Esc}" ; ifall man är inne i nånting redan

    MouseMove 50, 125 ; flytta musen till "filter"
    MouseClick


    send "^{a}"

    send "{Del}"

    Send "kunder"
    Sleep 100

    MouseMove 70, 175 ; flytta musen till "kunder"
    MouseClick
    Sleep 100

    MouseMove 250, 128 ; flytta musen till kundnr
    Sleep 100

    MouseClick
    send "^{v}"
    Sleep 100
    send "{Home}"
    Sleep 100

    MouseMove 925, 500
    MouseClick
    Sleep 100
    MouseMove 200, 500

    Send "{Wheeldown}"
    Sleep 10
    Send "{Wheeldown}"
    Sleep 10
    Send "{Wheeldown}"
    Sleep 10
    
    MouseMove 450, 550
    Send "{WheelLeft}"
    Sleep 10
    Send "{WheelLeft}"
    Sleep 10
    Send "{WheelLeft}"
    Sleep 10


}