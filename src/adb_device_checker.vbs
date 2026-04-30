'=============================================================================
'  This script is used for checking if device is connected properly.
'
'=============================================================================

Function checkAdbConnection()
	retCode = 0
	Set WshShell = CreateObject("WScript.Shell")
	On Error Resume Next
	Set retExec = WshShell.Exec("adb shell date") ' Exec() returns a WshScriptExec object
	If Err.number <> 0 Then
		' adb command not found. Could search for "android platform tools" over Internet,
		' or directly download it from https://developer.android.google.cn/studio/releases/platform-tools?hl=zh-cn
		retBtn = MsgBox("adb command is not installed." & vbCrLf & "Download it now?", vbOKCancel+vbQuestion, "Essential Component Missing")
		If retBtn = vbOK Then
			WshShell.Run "https://developer.android.google.cn/studio/releases/platform-tools?hl=zh-cn", 3  ' open in maximum window
		End If
		retCode = -1
		checkAdbConnection = retCode
		Exit Function
	End If
	' Status code is 0 means command is being executed.
	' Wait until command get completely executed (Status code becomes 1)
	Do While retExec.Status = 0
		 WScript.Sleep 100
	Loop
	' check if any output from StdErr
	strFromStderr = retExec.StdErr.ReadAll
	'MsgBox "stderr: " & strFromStderr
	retFound = InStr(strFromStderr, "no devices")
	If retFound <> 0 Then
		MsgBox "Android device is not connected, or ADB is not enabled. Please check.", vbOKOnly+vbExclamation, "Error"
		retCode = -1
	End If
	checkAdbConnection = retCode
End Function