#Requires AutoHotkey v2
End::ExitApp
SetTitleMatchMode 2

; Disable the F6 key
F6::return

HotIfWinActive "Bokning - http://prod.gonet.se/bookit/" || "Check-in - http://prod.gonet.se/bookit/"
CoordMode "Mouse", "Window"
SetKeyDelay 100

^p::findCard()
SetTitleMatchMode 2

findCard() {

WinActivate "Bokning - http://prod.gonet.se/bookit/" || "Check-in - http://prod.gonet.se/bookit/" ;


MouseMove 620, 450
Sleep 100

MouseGetPos , , &id, &control
Kundnr := ControlGetText(control, "Bokning - http://prod.gonet.se/bookit/" || "Check-in - http://prod.gonet.se/bookit/")
A_Clipboard := Kundnr
Sleep 100

WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret
Sleep 100

send "{Esc}"
Sleep 100
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret

send "{Esc}"
Sleep 100
WinActivate "Registerunderhåll - http://prod.gonet.se/bookit/" ; Fokusera på Registerunderhåll-fönstret

send "{Esc}"
Sleep 100
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

}