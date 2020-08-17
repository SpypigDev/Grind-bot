Set WshShell = WScript.CreateObject("WScript.Shell")

Dim sInput
sInput = InputBox("Enter what you grinding for")
sTime = InputBox("Whats the cooldown time? (In Min)")

sTime = sTime * "60000"
val = 0
var = "5000"
WScript.Sleep 5000
 WshShell.SendKeys "Grind bot online.."
 WshShell.SendKeys "{ENTER}"
WScript.Sleep 5000

Do
 WshShell.SendKeys "rpg " & sInput
 WshShell.SendKeys "{ENTER}"
 WScript.Sleep sTime
 WshShell.SendKeys "rpg " & sInput
 WshShell.SendKeys "{ENTER}"
 WScript.Sleep sTime
 WshShell.SendKeys "rpg " & sInput
 WshShell.SendKeys "{ENTER}"
 WScript.Sleep sTime
WshShell.SendKeys "Imitating human"
 WshShell.SendKeys "{ENTER}"
 WScript.Sleep 10000
WshShell.SendKeys "buffering"
 WshShell.SendKeys "{ENTER}"
 WScript.Sleep 10000
Loop
