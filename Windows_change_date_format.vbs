Set oShell = CreateObject("WScript.Shell")

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

Wscript.Sleep 500

oShell.SendKeys "E"

Dim x

for x=1 to 100
oShell.SendKeys "{Down}"
next

oShell.SendKeys "{Enter}"

for x=1 to 10
oShell.SendKeys "{TAB}"
Wscript.Sleep 100
next

oShell.SendKeys "{Enter}"

for x=1 to 2
oShell.SendKeys "{TAB}"
Wscript.Sleep 100
next

oShell.SendKeys "d"

for x=1 to 8
oShell.SendKeys "{TAB}"
Wscript.Sleep 100
next

oShell.SendKeys "{Enter}"

for x=1 to 10
oShell.SendKeys "{TAB}"
Wscript.Sleep 100
next

oShell.SendKeys "{RIGHT}"

Wscript.Sleep 500

oShell.SendKeys "{TAB}"

Wscript.Sleep 500

oShell.SendKeys "%{Down}"

Wscript.Sleep 500

oShell.SendKeys "a"

Wscript.Sleep 2500

oShell.SendKeys "u"

for x=1 to 6
oShell.SendKeys "{DOWN}"
next

oShell.SendKeys "{Enter}"

for x=1 to 3
oShell.SendKeys "{TAB}"
next

oShell.SendKeys "{Enter}"
Wscript.Sleep 500
oShell.SendKeys "{Enter}"

set oShell=nothing