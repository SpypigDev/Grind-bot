Set WshShell = WScript.CreateObject("WScript.Shell")

Dim sInput
sInput = InputBox("Enter what you grinding for")
sTime = InputBox("Whats the cooldown time? (In Min)")
sTime = sTime * "60000"
sTime = sTime + "1000"
val = 0
var = "1000"
WScript.Sleep 1000
 WshShell.SendKeys "Grind bot online.."
 WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000
clock = 0
addtime = 3600000
quarter = clock / "4"
clock = clock + sTime + sTime + sTime

Do
if myNum = 6 then
  clock = 0
  End If

if clock >= addtime then
  Do Until myNum >= 6
  With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
  WshShell.SendKeys clock
  WshShell.SendKeys "{ENTER}"
  With CreateObject("WScript.Shell")
  WScript.Sleep 10000
.Run "nircmd setcursor 400 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
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
WScript.Sleep 1000
  WshShell.SendKeys  "rpg " & sInput
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 1000
  With CreateObject("WScript.Shell")
.Run "nircmd setcursor 400 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
WScript.Sleep sTime


With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
With CreateObject("WScript.Shell")
WScript.Sleep 1000
.Run "nircmd setcursor 400 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
WScript.Sleep sTime


With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000
With CreateObject("WScript.Shell")
.Run "nircmd setcursor 400 10", 0, True
.Run "nircmd sendmouse left click", 0, True
WScript.Sleep 1000
End With
WshShell.SendKeys  "rpg " & sInput
WshShell.SendKeys "{ENTER}"
WScript.Sleep sTime





With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "Imitating human"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000
With CreateObject("WScript.Shell")
.Run "nircmd setcursor 400 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "Imitating human"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 10000


With CreateObject("WScript.Shell")
.Run "nircmd setcursor 129 14", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "Buffering"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000
With CreateObject("WScript.Shell")
.Run "nircmd setcursor 400 10", 0, True
.Run "nircmd sendmouse left click", 0, True
End With
WScript.Sleep 1000
WshShell.SendKeys  "Buffering"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 10000

clock = clock + sTime + sTime + sTime
  End If
  Loop
