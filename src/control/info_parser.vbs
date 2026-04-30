'=============================================================================
'  This is the Control component, responsible for parsing desired info in
'  dump file and notify to View component.
'
'=============================================================================



'----------------------------------------------------
' Business logic about parsing dump info as below
'----------------------------------------------------
'dumpFilePath = "..\tests\test_dumpsys.txt"
dumpFilePath = ".\freshDump.txt"
psFilePath = ".\freshPS.txt"
assembledUpdateInfo = ""

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
'		i = 1
		'strTrackInfo = parseDumpFile() ' TODO: we can save the process for writing dump file, instead we parse strings in memory directly after executing adb command
		strParsedInfo = parseDumpFile(strModelInfo)
'		For Each listenerElement in mUiListeners
'			'MsgBox "Updating changes Control-->View listener..."
'			listenerElement.updateInfo strParsedInfo
'			'listenerElement.updateInfo "TAG_APP|pid 2333, track 0=TAG_THREAD|id 0x5678=TAG_TRACK|0, 0x5678, 1944=TAG_BUFFER|1024 bytes=TAG_DEVICE|pcmC0D0p, running" '"TAG_APP|pid 2333, track 0=TAG_THREAD|id 0x5678=TAG_BUFFER|1024 bytes=TAG_DEVICE|pcmC0D0p, running" '
'			i = i + 1
'		Next
		updateInfo = 0
	End Function

	Public Function parseDumpFile(ByVal strDumpInfo)
		' wait until dump info is collected
		'Do Until bFlagDumpFilePrepared
		'	WScript.Sleep 100
		'Loop

		'MsgBox "start parsing..."
		Set fso = CreateObject("Scripting.FileSystemObject")
		' read complete dump info
		Const ForReading = 1  ' TODO: move this constant variable to a separated vbs file and use it as included
		Const ForWriting = 2
		'Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
		Set dumpFile = fso.OpenTextFile(dumpFilePath, ForWriting, True)
		'dumpFileContent = dumpFile.readAll
		dumpFile.Write strDumpInfo
		dumpFileContent = strDumpInfo
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

			' parse and get active fast track info for each thread
'			Set trackArray = CreateObject("System.Collections.ArrayList")
'			ret = findFastTrackPos(trackInfoInThread, trackArray) ' return value will be 0, if no active track is found in this thread
'			For Each trackElement in trackArray
'				trackElementArray = Split(trackElement, "+")
'				Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
'				dumpFile.read(fileOffset)
'				dumpFile.read(CLng(trackElementArray(1)))
'				trackLine = dumpFile.readLine
'				processedTrackLine = Trim(replaceStr("\s+", " ", trackLine)) ' replace consecutive spaces with just one space
'				trackStatusArray = Split(processedTrackLine)  ' convert line to array.
'				' track info format as "Index  Active  Full  Partial  Empty  Recent  Ready  Written"
'				MsgBox "Track " & trackStatusArray(0) & " in " & threadInfoArray(i)(0) & " has written " & trackStatusArray(7) & " bytes data.", vbOkOnly+vbInformation
'				dumpFile.Close
'				Set dumpFile = Nothing
'			Next
'			Set trackArray = Nothing


			' parse and get thread standy state
			Set threadStateArray = CreateObject("System.Collections.ArrayList")
			ret = findTargetStrPos("Standby:", trackInfoInThread, threadStateArray)
			threadStateElement = threadStateArray.Item(0)
			threadStateElementArray = Split(threadStateElement, "+")
			Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
			dumpFile.read(fileOffset)
			dumpFile.read(CLng(threadStateElementArray(1)))
			threadStateContent = dumpFile.readLine
			threadStateContentArray = Split(threadStateContent, ":")
			threadStateString = Trim(threadStateContentArray(1))
			If threadStateString = "yes" Then
				threadStateString = "Idle"
			Else
				threadStateString = "Running"
			End If
			dumpFile.Close
			Set dumpFile = Nothing
			Set threadStateArray = Nothing

			' parse and get AHAL buffer address per thread
			Set ahalAddressArray = CreateObject("System.Collections.ArrayList")
			ret = findTargetStrPos("Sink buffer :", trackInfoInThread, ahalAddressArray)
			ahalAddressElement = ahalAddressArray.Item(0)
			ahalAddressElementArray = Split(ahalAddressElement, "+")
			Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
			dumpFile.read(fileOffset)
			dumpFile.read(CLng(ahalAddressElementArray(1)))
			ahalAddressContent = dumpFile.readLine
			ahalAddressContentArray = Split(ahalAddressContent, ":")
			ahalAddressString = Trim(ahalAddressContentArray(1))
			'MsgBox threadInfoArray(i)(0) & " is using AHAL buffer at addr " & ahalAddressString
			dumpFile.Close
			Set dumpFile = Nothing
			Set ahalAddressArray = Nothing

			' parse and get frames written per thread
			Set framesWrittenInThreadArray = CreateObject("System.Collections.ArrayList")
			ret = findTargetStrPos("Total writes:", trackInfoInThread, framesWrittenInThreadArray)
			framesWrittenElement = framesWrittenInThreadArray.Item(0)
			framesWrittenElementArray = Split(framesWrittenElement, "+")
			Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
			dumpFile.read(fileOffset)
			dumpFile.read(CLng(framesWrittenElementArray(1)))
			framesWrittenContent = dumpFile.readLine
			framesWrittenContentArray = Split(framesWrittenContent, ":")
			framesWrittenString = Trim(framesWrittenContentArray(1))
			'MsgBox threadInfoArray(i)(0) & " is using AHAL buffer at addr " & ahalAddressString
			dumpFile.Close
			Set dumpFile = Nothing
			Set framesWrittenInThreadArray = Nothing

			' parse and get AHAL frame count per thread
			Set framesCountInThreadArray = CreateObject("System.Collections.ArrayList")
			ret = findTargetStrPos("HAL frame count:", trackInfoInThread, framesCountInThreadArray)
			framesCountElement = framesCountInThreadArray.Item(0)
			framesCountElementArray = Split(framesCountElement, "+")
			Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
			dumpFile.read(fileOffset)
			dumpFile.read(CLng(framesCountElementArray(1)))
			framesCountContent = dumpFile.readLine
			framesCountContentArray = Split(framesCountContent, ":")
			framesCountString = Trim(framesCountContentArray(1))
			dumpFile.Close
			Set dumpFile = Nothing
			Set framesCountInThreadArray = Nothing

			' parse and get AHAL frame size per thread
			Set framesSizeInThreadArray = CreateObject("System.Collections.ArrayList")
			ret = findTargetStrPos("Processing frame size:", trackInfoInThread, framesSizeInThreadArray)
			framesSizeElement = framesSizeInThreadArray.Item(0)
			framesSizeElementArray = Split(framesSizeElement, "+")
			Set dumpFile = fso.OpenTextFile(dumpFilePath, ForReading)
			dumpFile.read(fileOffset)
			dumpFile.read(CLng(framesSizeElementArray(1)))
			framesSizeContent = dumpFile.readLine
			framesSizeContentArray = Split(framesSizeContent, ":")
			framesSizeStringTmp = Trim(framesSizeContentArray(1))
			framesSizeStringArray = Split(framesSizeStringTmp)
			framesSizeString = framesSizeStringArray(0)
			'framesSizeString = ""
			dumpFile.Close
			Set dumpFile = Nothing
			Set framesSizeInThreadArray = Nothing


			'If threadStateString = "Running" Then

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

						' If any track is found, get ps info for later search work
						If trackLineCnt > 0 Then
							Set WshShell = CreateObject("WScript.Shell")
							psText = ""
							On Error Resume Next
							Set retExec = WshShell.Exec("adb shell ps -A") ' Exec() returns a WshScriptExec object
							If Err.number <> 0 Then
								' ideally, adb tool is gurantted as installed at this stage
							End If
							Do While Not retExec.StdOut.AtEndOfStream
								psText = retExec.StdOut.ReadAll()
								WScript.Sleep 100
							Loop
							' Status code is 0 means command is being executed.
							' Wait until command get completely executed (Status code becomes 1)
							Do While retExec.Status = 0
								 WScript.Sleep 100
							Loop
							'msgbox psText
							'WScript.Echo psText
							Set psFile = fso.OpenTextFile(psFilePath, ForWriting, True)
							psFile.Write psText
							psFile.Close
							Set psFile = Nothing  ' release memory
						End If
						For j = 1 To trackLineCnt
							trackLine = dumpFile.readLine
							'MsgBox "Line " & dumpFile.Line & ": " & trackLine
							processedTrackLine = Trim(replaceStr("\s+", " ", trackLine)) ' replace consecutive spaces with just one space
							trackDetailedInfoArray = Split(processedTrackLine)  ' convert line to array.
							'MsgBox "AppId " & trackDetailedInfoArray(3) & " is playing audio."
							If LCase(trackDetailedInfoArray(0)) = trackDetailedInfoArray(0) And UCase(trackDetailedInfoArray(0)) = trackDetailedInfoArray(0) Then
								trackLine = "CTrack" & trackLine
								processedTrackLine = Trim(replaceStr("\s+", " ", trackLine)) ' replace consecutive spaces with just one space
								trackDetailedInfoArray = Split(processedTrackLine)  ' convert line to array.
							End If

							' parse and get APP name from ps info
							Set psInfoArray = CreateObject("System.Collections.ArrayList")
							ret = findTargetStrPos(trackDetailedInfoArray(3), psText, psInfoArray)
							psInfoElement = psInfoArray.Item(0)
							psInfoElementArray = Split(psInfoElement, "+")
							Set psFile = fso.OpenTextFile(psFilePath, ForReading)
							psFile.read(CLng(psInfoElementArray(1)))
							psInfoContent = psFile.readLine
							processedPsInfoContent = replaceStr("\s+", " ", psInfoContent)
							processedPsInfoContentArray = Split(processedPsInfoContent)
							'MsgBox "UBound: " & UBound(processedPsInfoContentArray)
							processedPsInfoString = processedPsInfoContentArray(UBound(processedPsInfoContentArray))
							'MsgBox "App " & chr(34) & processedPsInfoString & chr(34) & " is playing audio."
							psFile.Close
							Set psFile = Nothing
							Set psInfoArray = Nothing

							'assembledUpdateInfo = "TAG_APP|" & trackDetailedInfoArray(3) & "," & processedPsInfoString & "=TAG_THREAD|" & threadInfoArray(i)(0) & "," & "Running" & "=TAG_TRACK|" & j & "=TAG_BUFFER|" & ahalAddressString & "=TAG_DEVICE|pcmC0D0p, running"
							assembledUpdateInfo = "TAG_APP|" & trackDetailedInfoArray(3) & "," & processedPsInfoString & "," & trackDetailedInfoArray(2) & "=TAG_THREAD|" & threadInfoArray(i)(0) & "," & threadStateString & "=TAG_TRACK|" & threadInfoArray(i)(0) & "," & trackDetailedInfoArray(1) & "," & trackDetailedInfoArray(0) & "=TAG_BUFFER|" & ahalAddressString & "," & framesSizeString & "*" & framesCountString & "bytes" & "=TAG_DEVICE|pcmC0D0p, running"
							'WScript.Echo assembledUpdateInfo
							For Each listenerElement in mUiListeners
								'MsgBox "Updating changes Control-->View listener..."
								listenerElement.updateInfo assembledUpdateInfo
							Next
							assembledUpdateInfo = ""
						Next
						dumpFile.Close
						Set dumpFile = Nothing
					Next
				Else
					' no detailed track found in a single thread, this represents thread is in clean standby state
					assembledUpdateInfo = "TAG_THREAD|" & threadInfoArray(i)(0) & "," & threadStateString & "Completely"
					'WScript.Echo assembledUpdateInfo
					For Each listenerElement in mUiListeners
						'MsgBox "Updating changes Control-->View listener..."
						listenerElement.updateInfo assembledUpdateInfo
					Next
					assembledUpdateInfo = ""
				End If
				Set detailedTrackArray = Nothing
				Set effectChainArray = Nothing
			'End If


			' add other parsing logic below
			' ------------------------------

			'-------------------------------
			'MsgBox i
		Next
		' TAG_APP|pid 2333, track 0=TAG_THREAD|id 0x5678=TAG_TRACK|0, 0x5678, 1944=TAG_BUFFER|1024 bytes=TAG_DEVICE|pcmC0D0p, running
		'parseDumpFile = "TAG_APP|" & trackDetailedInfoArray(3) & "," & processedPsInfoString & "=TAG_THREAD|" & threadInfoArray(i)(0) & "=TAG_TRACK|" & j & "=TAG_BUFFER|" & ahalAddressString
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

Function findFastTrackPos(ByRef completeDumpInfo, ByRef saveToArray)
	strPattern = "[^1-9][0-9]\s*yes"
	findFastTrackPos = findTargetStrPos(strPattern, completeDumpInfo, saveToArray)
End Function

Function replaceStr(ByVal fromStrPattern, ByVal toStr, ByRef targetStr)
	Set re = New RegExp
	re.Pattern = fromStrPattern
	re.Global = True
	replaceStr = re.Replace(targetStr, toStr)
End Function