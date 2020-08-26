Set WshShell = WScript.CreateObject("WScript.Shell")
sTime = "60"
sTime = sTime * "60000"
sTime = sTime + "1000"
val = 0
var = "5000"
WScript.Sleep 5000
 WshShell.SendKeys "Bump bot online.."
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
