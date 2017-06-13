
Set wshShell =wscript.CreateObject("WScript.Shell")
do
wscript.sleep 100
'change the number above (100) to vary the flickering speeds
wshshell.sendkeys "{CAPSLOCK}"
wshshell.sendkeys "{NUMLOCK}"
wshshell.sendkeys "{SCROLLLOCK}"
loop
