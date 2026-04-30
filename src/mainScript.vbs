'=============================================================================
'  This is entry of whole project.
'
'=============================================================================

include ".\dll_status_checker.vbs"
include ".\adb_device_checker.vbs"
include ".\view\ui_painter.vbs"
include ".\control\info_parser.vbs"
include ".\model\info_gatherer.vbs"

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
model.startCollectingInfo(3000)

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