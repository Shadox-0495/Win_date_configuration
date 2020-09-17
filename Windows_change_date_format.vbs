Set oShell = CreateObject("WScript.Shell")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
Winv = ""
For Each os in oss
if InStr(os.Caption,"10")>0 then Winv = "10"
if InStr(os.Caption,"7")>0 then Winv = "7"
Next

Wscript.Sleep 1000

oShell.SendKeys "^{ESC}"
	
Wscript.Sleep 750

oShell.SendKeys "Change the way dates and lists are displayed"

Wscript.Sleep 1500

oShell.SendKeys "{Enter}"

Wscript.Sleep 1500

oShell.SendKeys "%{Down}"

Wscript.Sleep 500

oShell.SendKeys "A"

Wscript.Sleep 3000

oShell.SendKeys "E"

Dim x

if Winv="10" then
	for x=1 to 100
		oShell.SendKeys "{Down}"
	next
End if

if Winv="7" then
	for x=1 to 14
		oShell.SendKeys "{Down}"
	next
End if

Wscript.Sleep 500

oShell.SendKeys "{Enter}"

Wscript.Sleep 500

if Winv="10" then
	for x=1 to 10
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if

if Winv="7" then
	for x=1 to 11
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if

oShell.SendKeys "{Enter}"

for x=1 to 2
oShell.SendKeys "{TAB}"
Wscript.Sleep 100
next

oShell.SendKeys "d"

if Winv="10" then
	for x=1 to 8
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if

if Winv="7" then
	for x=1 to 9
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if

oShell.SendKeys "{Enter}"


if Winv="10" then
	for x=1 to 10
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if

if Winv="7" then
	for x=1 to 11
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if


oShell.SendKeys "{RIGHT}"

Wscript.Sleep 500

oShell.SendKeys "{TAB}"

Wscript.Sleep 500

oShell.SendKeys "%{Down}"

Wscript.Sleep 500

oShell.SendKeys "a"

Wscript.Sleep 3000

oShell.SendKeys "u"


if Winv="10" then
	for x=1 to 6
		oShell.SendKeys "{DOWN}"
	next
End if

if Winv="7" then
	for x=1 to 4
		oShell.SendKeys "{DOWN}"
	next
End if

oShell.SendKeys "{Enter}"


if Winv="10" then
	for x=1 to 3
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if

if Winv="7" then
	for x=1 to 4
		oShell.SendKeys "{TAB}"
		Wscript.Sleep 100
	next
End if

oShell.SendKeys "{Enter}"
Wscript.Sleep 500
oShell.SendKeys "{Enter}"

set oShell=nothing
set dtmConvertedDate=nothing
set objWMIService=nothing
set oss=nothing
