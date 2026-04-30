'=============================================================================
'  This is entry of whole project.
'
'=============================================================================

include ".\dll_status_checker.vbs"
include ".\adb_device_checker.vbs"
'include ".\ui_painter.vbs"
'include ".\info_parser.vbs"
'include ".\info_gatherer.vbs"

' check if dll environment is prepared
retStatus = checkAndInstallDllEnvironment()
If retStatus <> 0 Then
	WScript.Quit
End If

' check if an Android device is connected
retStatus = checkAdbConnection()
If retStatus <> 0 Then
	WScript.Quit
End If

' wso.dll is guaranteed to be OK at this stage,
' now let's start worker thread to collect and analyze dumpsys info.
'Set shell = CreateObject("WScript.Shell")
'shell.Run ".\dumpsys_worker_thread.vbs"

Set baseWso = WScript.CreateObject("Scripting.WindowSystemObject")

Set view = New DumpsysView
view.init(baseWso)
view.paintUI

Set controller = New DumpsysControl
controller.init(baseWso)
controller.registerListener(view)

Set model = New DumpsysModel
model.init(baseWso)
model.registerListener(controller)
model.startCollectingInfo(10000)

baseWso.Run()


'----------------------------------------------------
' Below are function definitions
'----------------------------------------------------
Sub include(vbsFilePath)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set scriptFile = fso.OpenTextFile(vbsFilePath)
	scriptFileContent = scriptFile.ReadAll
	scriptFile.Close
	ExecuteGlobal scriptFileContent
End Sub




'=============================================================================
'  This is the worker thread for collecting and analyzing dumpsys information.
'
'  Ideally, this worker thread shalln't own a form.
'
'=============================================================================
'----------------------------------------------------
' register listener to dumpsys_worker_thread for
' receiving parsed dump info.
'----------------------------------------------------
Class DumpsysView
	Private mWso
	Private mIdNum
	Private mFlagClsInited ' = false
	Private mForm, mText1

	Public Property Get id()
		id = mIdNum
	End Property

	Public Function init(ByRef wso)
		'Set mWso = WScript.CreateObject("Scripting.WindowSystemObject")
		Set mWso = wso
		mFlagClsInited = true
		init = 0
	End Function

	Public Function paintUI()
		Set mForm = mWso.CreateForm(200,200,650,400)
		mForm.Text = "Worker Thread Demo"
		Set mText1 = mForm.TextOut(10,60,"0")

		'----------------------------------------------------
		' Show UI and start running script
		If mFlagClsInited Then
			mForm.Show()
			'mWso.Run()
		End If
		paintUI = 0
	End Function

	Public Function setIdNum(ByVal idNum)
		mIdNum = idNum
		setIdNum = mIdNum
	End Function

	Public Function updateInfo(ByVal strInfo)
		MsgBox "View receives feedback from Control for " & strInfo
		mText1.Text = strInfo
		updateInfo = 0
	End Function
End Class





'=============================================================================
'  This is the worker thread for collecting and analyzing dumpsys information.
'
'  Ideally, this worker thread shalln't own a form.
'
'=============================================================================

'----------------------------------------------------
' Business logic about parsing dump info as below
'----------------------------------------------------
'dumpFilePath = "..\tests\test_dumpsys.txt"
dumpFilePath = ".\freshDump.txt"


Class DumpsysControl
	Private mWso
	Private mUiListeners
	Private mFlagClsInited ' = false

	Public Function init(ByRef wso)
		Set mUiListeners = CreateObject("System.Collections.ArrayList")
		Set mWso = wso
		mFlagClsInited = true
		init = 0
	End Function

	Public Function finalize()
		Set mUiListeners = Nothing
		mFlagClsInited = false
		finalize = 0
	End Function

	Public Function registerListener(ByRef uiListener)
		If mFlagClsInited Then
			mUiListeners.Add uiListener
		Else
			MsgBox "register listener uiListener failed!", vbOKOnly+vbExclamation, "Error"
			registerListener = 1
			Exit Function
		End If
		registerListener = 0
	End Function

	Public Function updateInfo(ByVal strModelInfo)
		cnt = mUiListeners.Count
		i = 1
		'strTrackInfo = parseDumpFile() ' TODO: we can save the process for writing dump file, instead we parse strings in memory directly after executing adb command
		For Each listenerElement in mUiListeners
			MsgBox "Updating changes Control-->View listener..."
			listenerElement.updateInfo strModelInfo
			i = i + 1
		Next
		updateInfo = 0
	End Function

	Public Function parseDumpFile()
		' wait until dump info is collected
		Do Until bFlagDumpFilePrepared
			WScript.Sleep 100
		Loop

		Set fso = CreateObject("Scripting.FileSystemObject")
		' read complete dump info
		Const ForReading = 1  ' TODO: move this constant variable to a separated vbs file and use it as included
		Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
		dumpFileContent = dumpFile.readAll
		dumpFile.Close
		Set dumpFile = Nothing  ' release memory

		' parse dump info to get output thread string location
		Set threadInfoArray = CreateObject("System.Collections.ArrayList")
		Set outputThreadArray = CreateObject("System.Collections.ArrayList")
		ret = findOutputThreadPos(dumpFileContent, outputThreadArray)
		For Each threadElement in outputThreadArray
			threadElementArray = Split(threadElement, "+")
			threadInfoArray.Add threadElementArray
		Next
		'MsgBox "thread info size: " & threadInfoArray.Count

		' process thread info one by one
		iMax = threadInfoArray.Count - 1
		For i = 0 to iMax
			' read track info within a thread info block
			Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
			fileOffset = CLng(threadInfoArray(i)(1))
			'MsgBox threadInfoArray(i)(0) & ", pos: " & threadInfoArray(i)(1)
			trackInfoInThread = dumpFile.read(fileOffset)  ' read until desired output thread shows up
			If i <> iMax Then
				'MsgBox "previous: " & chr(34) & threadInfoArray(i)(1) & chr(34) & ", latter: " & chr(34) & threadInfoArray(i+1)(1) & chr(34)
				threadInfoSegmentLen = CLng(threadInfoArray(i+1)(1)) - CLng(threadInfoArray(i)(1))
			Else
				threadInfoSegmentLen = Len(dumpFileContent) - CLng(threadInfoArray(i)(1))
			End If
			trackInfoInThread = dumpFile.read(threadInfoSegmentLen)
			'WScript.Echo trackInfoInThread
			'WScript.Echo "=========================================================================================="
			dumpFile.Close
			Set dumpFile = Nothing

			' parse and get active track info for each thread
			Set trackArray = CreateObject("System.Collections.ArrayList")
			ret = findTrackPos(trackInfoInThread, trackArray) ' return value will be 0, if no active track is found in this thread
			For Each trackElement in trackArray
				trackElementArray = Split(trackElement, "+")
				Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
				dumpFile.read(fileOffset)
				dumpFile.read(CLng(trackElementArray(1)))
				trackLine = dumpFile.readLine
				processedTrackLine = Trim(replaceStr("\s+", " ", trackLine)) ' replace consecutive spaces with just one space
				trackStatusArray = Split(processedTrackLine)  ' convert line to array.
				' track info format as "Index  Active  Full  Partial  Empty  Recent  Ready  Written"
				MsgBox "Track " & trackStatusArray(0) & " in " & threadInfoArray(i)(0) & " has written " & trackStatusArray(7) & " bytes data.", vbOkOnly+vbInformation
				dumpFile.Close
				Set dumpFile = Nothing
			Next
			Set trackArray = Nothing



			' search for "Id Active Client" to locate more detailed track info, and search for "Effect Chains" for end of detailed track list
			trackInfoStartLine = 0
			trackInfoEndLine = 0

			Set detailedTrackArray = CreateObject("System.Collections.ArrayList")
			ret = findTargetStrPos("Id Active Client", trackInfoInThread, detailedTrackArray)
			If ret <> 0 Then
				' Action for searching trackInfoEndLine shall depend on search result of "Id Active Client", as
				' "Effect Chains" shows up even there's no track exists in a thread.
				Set effectChainArray = CreateObject("System.Collections.ArrayList")
				ret = findTargetStrPos("[0-9] Effect Chains", trackInfoInThread, effectChainArray)
				' string "Effect Chains" always show up, no matter there is active track or effect chain
				' Read the first element only
				effectChainElement = effectChainArray.Item(0)
				effectChainElementArray = Split(effectChainElement, "+")
				Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
				dumpFile.read(fileOffset)
				dumpFile.read(CLng(effectChainElementArray(1)))
				trackInfoEndLine = dumpFile.Line
				dumpFile.Close
				Set dumpFile = Nothing

				'MsgBox "found " & ret & " targets, for thread " & i
				For Each detailedTrackElement in detailedTrackArray
					detailedTrackElementArray = Split(detailedTrackElement, "+")
					Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
					dumpFile.read(fileOffset)
					dumpFile.read(CLng(detailedTrackElementArray(1)))
					trackInfoStartLine = dumpFile.Line
					dumpFile.skipLine
					trackLineCnt = trackInfoEndLine - trackInfoStartLine - 1
					'MsgBox "startLine: " & trackInfoStartLine & ", endLine: " & trackInfoEndLine & ", TrackCnt: " & trackLineCnt
					For j = 1 To trackLineCnt
						trackLine = dumpFile.readLine
						MsgBox "Line " & dumpFile.Line & ": " & trackLine
					Next
					dumpFile.Close
					Set dumpFile = Nothing
				Next
			End If
			Set detailedTrackArray = Nothing
			Set effectChainArray = Nothing



			' add other parsing logic below
			' ------------------------------

		Next
		parseDumpFile = 0
	End Function
End Class

'----------------------------------------------------
' Below are function definitions
'----------------------------------------------------
Function findTargetStrPos(ByVal strPattern, ByRef completeDumpInfo, ByRef saveToArray)
	Set re = New RegExp
	re.Pattern = strPattern
	re.IgnoreCase = False
	re.Global = True
	Set foundStrs = re.Execute(completeDumpInfo)
	idxStr = 0
	For Each str in foundStrs
		' format is "value+position"
		saveToArray.Add str.value & "+" & str.FirstIndex
		' or use ReDim Preserve array(idxStr+1) to rescale array size?
		idxStr = idxStr + 1
	Next
	findTargetStrPos = idxStr
End Function

Function findOutputThreadPos(ByRef completeDumpInfo, ByRef saveToArray)
	strPattern = "Output thread 0x[0-9a-z]*"
	findOutputThreadPos = findTargetStrPos(strPattern, completeDumpInfo, saveToArray)
End Function

Function findTrackPos(ByRef completeDumpInfo, ByRef saveToArray)
	strPattern = "[^1-9][0-9]\s*yes"
	findTrackPos = findTargetStrPos(strPattern, completeDumpInfo, saveToArray)
End Function

Function replaceStr(ByVal fromStrPattern, ByVal toStr, ByRef targetStr)
	Set re = New RegExp
	re.Pattern = fromStrPattern
	re.Global = True
	replaceStr = re.Replace(targetStr, toStr)
End Function






'=============================================================================
'  This is the worker thread for collecting and analyzing dumpsys information.
'
'  Ideally, this worker thread shalln't own a form.
'
'=============================================================================
'-------------------------------------------------------

bFlagDumpFilePrepared = false

Class DumpsysModel
	Private mWso
	Private mListeners
	Private mFlagClsInited ' = false
	Private mTimer

	Public Function init(ByRef wso)
		'Set mWso = WScript.CreateObject("Scripting.WindowSystemObject")
		Set mWso = wso
		Set mListeners = CreateObject("System.Collections.ArrayList")
		' set a timer
		Set mTimer = mWso.CreateTimer()
		mTimer.Interval = 10000    ' 10000ms as default value
		mTimer.OnExecute = GetRef("OnTimer")
		mTimer.Active = false

		'mWso.Run()
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
		i = 1
		For Each listenerElement in mListeners
			MsgBox "Updating changes Model-->Control listener..."
			listenerElement.updateInfo i & ": " & strDumpInfo
			i = i + 1
		Next
		updateInfo = 0
	End Function

End Class

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

	Set tmpFile = fso.OpenTextFile(".\freshDump.txt", 2, true)
	tmpFile.Write dumpText
	tmpFile.Close

	bFlagDumpFilePrepared = true
	'parseDumpFile
	model.updateInfo "fake dump info..."
End Sub
