Set WshShell = WScript.CreateObject("WScript.Shell")

Dim sInput
sInput = InputBox("Enter what you grinding for")
sTime = InputBox("Whats the cooldown time? (In Min)")
sTime = sTime * "60000"
sTime = sTime + "1000"
val = 0
var = "5000"
WScript.Sleep 5000
 WshShell.SendKeys "Grind bot online.."
 WshShell.SendKeys "{ENTER}"
WScript.Sleep 5000

intAnswer = _
    Msgbox("Are you running grind bot over night?", _
        vbYesNo, "Yes")
        WScript.Sleep 5000
If intAnswer = vbYes Then
  if clock >= "3,600,000" then
  WshShell.SendKeys "Imitating asleep human (buffering)"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 300000
  WshShell.SendKeys "Imitating asleep human (buffering)"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 300000
  WshShell.SendKeys "Imitating asleep human (buffering)"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 300000
  WshShell.SendKeys "Imitating asleep human (buffering)"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 300000
  WshShell.SendKeys "Imitating asleep human (buffering)"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 300000
  WshShell.SendKeys "Imitating asleep human (buffering)"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 300000

  Else
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
  WshShell.SendKeys "Imitating asleep human"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 10000
  WshShell.SendKeys "buffering"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 10000
  clock = clock + sTime + sTime + sTime
  Loop
  End If
Else
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
End If
