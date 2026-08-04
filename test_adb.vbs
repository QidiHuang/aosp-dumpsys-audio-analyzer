Set WshShell = CreateObject("WScript.Shell")
Set retExec = WshShell.Exec("adb shell ""dumpsys media.audio_flinger; echo ---PS_START---; ps -A || ps""")
WScript.Echo "Started"
out = retExec.StdOut.ReadAll()
WScript.Echo Len(out)
