
Set wmp = CreateObject("WMPlayer.ocx.7")
wmp.URL = "..\easteregg\wav\keygen.wav"
'wmp.play ' can be omitted
Do While wmp.playState <> 1
	WScript.Sleep 100
Loop
wmp.Close
