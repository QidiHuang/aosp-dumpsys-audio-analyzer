'=============================================================================
'  This is the Model component, responsible for collecting dump info
'  periodically and notify to Control component.
'
'=============================================================================





'-------------------------------------------------------
Class DumpsysModel
	Private mWso
	Private mListeners
	Private mFlagClsInited ' = false
	Private mTimer

	Public Function init(ByRef wso)
		Set mWso = wso
		Set mListeners = CreateObject("System.Collections.ArrayList")
		' set a timer
		Set mTimer = mWso.CreateTimer()
		mTimer.Interval = 10000    ' 10000ms as default value
		mTimer.OnExecute = GetRef("OnTimer")
		mTimer.Active = false

		bFlagDumpFilePrepared  = false
		mFlagClsInited = true
		init = 0
	End Function

	Public Function startCollectingInfo(ByVal timerIntervalMs)
		mTimer.Interval = timerIntervalMs
		mTimer.Active = true
	End Function

	Public Function finalize()
		Set mListeners = Nothing
		' unregister listener
		mFlagClsInited = false
		finalize = 0
	End Function

	Public Function registerListener(ByRef controlListener)
		If mFlagClsInited Then
			mListeners.Add controlListener
		Else
			MsgBox "register listener controlListener failed!", vbOKOnly+vbExclamation, "Error"
			registerListener = 1
			Exit Function
		End If
		registerListener = 0
	End Function

	Public Function updateInfo(ByVal strDumpInfo)
		cnt = mListeners.Count
		iModel = 1
		For Each listenerElement in mListeners
			'MsgBox "Updating changes Model-->Control listener..."
			listenerElement.updateInfo iModel & ": " & strDumpInfo
			iModel = iModel + 1
		Next
		updateInfo = 0
	End Function

End Class


'----------------------------------------------------
' Below are function definitions
'----------------------------------------------------
Sub OnTimer(Sender)
	'Text1.Text = Text1.Text + 10
	'Set fso = CreateObject("Scripting.FileSystemObject")
	'If fso.FileExists(".\freshDump.txt") Then
	'	t.Active = false
	'	Exit Sub
	'End If

	bFlagDumpFilePrepared = false

	Set WshShell = CreateObject("WScript.Shell")
	dumpText = ""
	On Error Resume Next
	Set retExec = WshShell.Exec("adb shell dumpsys media.audio_flinger") ' Exec() returns a WshScriptExec object
	If Err.number <> 0 Then
		' ideally, adb tool is gurantted as installed at this stage
	End If
	Do While Not retExec.StdOut.AtEndOfStream
		dumpText = retExec.StdOut.ReadAll()
		WScript.Sleep 100
	Loop
	' Status code is 0 means command is being executed.
	' Wait until command get completely executed (Status code becomes 1)
	Do While retExec.Status = 0
		 WScript.Sleep 100
	Loop
	'msgbox dumpText

	'Set tmpFile = fso.OpenTextFile(".\freshDump.txt", 2, true)
	'tmpFile.Write dumpText
	'tmpFile.Close

	bFlagDumpFilePrepared = true

	'model.updateInfo "DumpInfo is prepared"
	model.updateInfo dumpText
End Sub
