Esc::ExitApp
SetTitleMatchMode 2
WinActivate "CardStudio" ; Fokusera på CardStudio-fönstret
AllaRader := FileRead("namnlista.txt") ; 
lines := StrSplit(AllaRader, "`n")
lineCount := lines.Length
send "^{p}"
Sleep 500
MouseMove 40, 400
Sleep 500
MouseClick
send "{Tab}"
Sleep 500

send "^{a}"
Sleep 500

send lineCount
Sleep 500

send "^{Enter}"
Sleep 500



Loop parse, AllaRader, "`n", "`r"  
{

    entries := StrSplit(A_LoopField, ",") ; Splitta nuvarande line till en array
	kundnr := entries[1]  ; Den första entryn är kundnr
	namn := entries[2]  ; Den andra är namn
	datum := entries[3]  ; Den tredje är datum
	Sleep 500
	
	; Magnetremsa
	
	MouseMove 164, 135 ; Markera textrutan
	Sleep 500

	MouseClick
		Sleep 500

	send "{End}" ; Ställ den efter 1209
		Sleep 500

	send kundnr
			Sleep 500

	mouseMove 200, 200
	MouseClick
		Sleep 500
		
	; Namn

	MouseMove 164, 135 ; Markera textrutan
	Sleep 500

	MouseClick
		Sleep 500
	send "^{a}"

	send namn
			Sleep 500

	mouseMove 200, 200
	MouseClick
	
		; Giltighetstid

	MouseMove 164, 135 ; Markera textrutan
	Sleep 500

	MouseClick
		Sleep 500
	send "^{a}"
	send datum
			Sleep 500

	mouseMove 200, 200
	MouseClick
	
		
		; Kundnummer

	MouseMove 164, 135 ; Markera textrutan
	Sleep 500

	MouseClick
		Sleep 500
	send "^{a}"

	send kundnr
			Sleep 500

	mouseMove 200, 200
	MouseClick
}
