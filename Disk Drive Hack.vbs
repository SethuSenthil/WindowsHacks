Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
do
if colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count -1
colCDROMs.Item(i).Eject
Next
For i = 0 to colCDROMs.Count -1
colCDROMs.Item(i).Eject
Next
End If
wscript.sleep 5000
'Change the speed of the Optical Disk Drive opening and closing by modifying the number above (5000)
loop
