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

Do
  if clock >= "3,600,000" then
  Do Until myNum = 6
  With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
  WshShell.SendKeys "Imitating asleep human (buffering)"
  WshShell.SendKeys "{ENTER}"
  With CreateObject("WScript.Shell")
.Run "nircmd setcursor 320 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys "Imitating asleep human (buffering)"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 300000
myNum = myNum + 1
Loop

  Else


  With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
  WshShell.SendKeys  "rpg " & sInput
  WshShell.SendKeys "{ENTER}"
  With CreateObject("WScript.Shell")
.Run "nircmd setcursor 320 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
WScript.Sleep sTime


With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
With CreateObject("WScript.Shell")
.Run "nircmd setcursor 320 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
WScript.Sleep sTime


With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
With CreateObject("WScript.Shell")
.Run "nircmd setcursor 320 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
WScript.Sleep sTime





With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "Imitating human"
WshShell.SendKeys "{ENTER}"
With CreateObject("WScript.Shell")
.Run "nircmd setcursor 320 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "Imitating human"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 10000


With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "Buffering"
WshShell.SendKeys "{ENTER}"
With CreateObject("WScript.Shell")
.Run "nircmd setcursor 320 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WshShell.SendKeys  "Buffering"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 10000

  clock = clock + sTime + sTime + sTime
  End If
  Loop
