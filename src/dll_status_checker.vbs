'=============================================================================
'  Using this script to check if wso.dll was registered into this system.
'
'  If dll was not registered, script will act as below:
'  a) check if dll library exist;
'  b) register dll library if file is found;
'  c) report error if dll library not found;
'
'=============================================================================

Function checkAndInstallDllEnvironment()
	retCode = 0
	Set shell = CreateObject("WScript.Shell")
	regKeyRootPath = "HKEY_CURRENT_USER\SOFTWARE\DumpsysAnalyzer\"
	regKeyStatusPath = regKeyRootPath & "dllStatus"

	On Error Resume Next
	dllStatus = shell.RegRead(regKeyStatusPath)
	If Err.number <> 0 Then
		' Reading key failed, means wso.dll was not registered or key is destroyed by mistake.
		'MsgBox "RegKey doesn't exist. Error code: " & err, vbOkOnly+vbInformation, "Read Reg Failed"
		Err.clear

		Set fso = CreateObject("Scripting.FileSystemObject")
		Set UAC = CreateObject("Shell.Application")

		currentPath = fso.GetFolder(".").Path
		'MsgBox "current script path: " & currentPath, vbOkOnly, ""

		' Try to clean environment
		'batFile_unreg = ".\unregisterDLL.bat"
		'If Not fso.FileExists(batFile_unreg) Then
		'	MsgBox "Try to clean env", vbOkOnly+vbInformation, ""
		'	Set batFileNew = fso.CreateTextFile(batFile_unreg, True)
		'	batFileNew.WriteLine("@echo off")
		'	batFileNew.WriteLine(":: unregister wso.dll libraries by regsvr32")
		'	batFileNew.WriteLine("echo unregister dlls...")
		'	batFileNew.WriteLine("regsvr32.exe /u " & currentPath & "\..\prebuilt_libs\WSO.dll")
		'	batFileNew.WriteLine("regsvr32.exe /u " & currentPath & "\..\prebuilt_libs\x64\WSO.dll")
		'	batFileNew.WriteLine("echo unregister dlls successfully.")
		'	'batFileNew.WriteLine("pause")
		'	batFileNew.Close
		'End If
		' Run as administrator to register wso.dll into system
		'UAC.ShellExecute batFile_unreg, "", "", "runas", 0
		' Wait for 4000ms for executing unregistration
		'WScript.Sleep(4000)

		' And try to regsiter wso.dll to re-establish environment in below code.

		' Check if dll libraries exist
		dllFile_x32 = "..\prebuilt_libs\WSO.dll"
		dllFile_x64 = "..\prebuilt_libs\x64\WSO.dll"
		batFile_reg = ".\registerDLL.bat"

		If fso.FileExists(dllFile_x32) And fso.FileExists(dllFile_x64) Then
			if Not fso.FileExists(batFile_reg) Then
				' Missing bat script for registering dll libraries,
				' try to create the bat script
				Set batFileNew = fso.CreateTextFile(batFile_reg, True)
				batFileNew.WriteLine("@echo off")
				batFileNew.WriteLine(":: register wso.dll libraries by regsvr32")
				batFileNew.WriteLine("echo register dlls...")
				batFileNew.WriteLine("regsvr32.exe " & currentPath & "\..\prebuilt_libs\WSO.dll")
				batFileNew.WriteLine("regsvr32.exe " & currentPath & "\..\prebuilt_libs\x64\WSO.dll")
				batFileNew.WriteLine("echo register dlls successfully.")
				'batFileNew.WriteLine("pause")
				batFileNew.Close
				Set batFileNew = Nothing
			End If
			ret = UAC.ShellExecute(batFile_reg, "", "", "runas", 0)

			' Write dll registration status to RegKey
			shell.RegWrite regKeyRootPath
			shell.RegWrite regKeyStatusPath, "OK"

			' Wait for 4000ms for executing registration
			' TODO: try to get active window and use SendKeys "{ENTER}" for confirming automatically
			WScript.Sleep(4000)
		Else
			' dll libraries are not complete
			MsgBox "ERROR: wso.dll libraries not registerd and couldn't be found.", vbOkOnly+vbInformation, "Environment Not Ready"
			'WScript.Quit
			retCode = -1
		End If
	Else
		'MsgBox "dll registration status: " & dllStatus, vbOkOnly, ""
	End If

	'MsgBox "Steady on she goes!", vbOkOnly+vbInformation, ""
	checkAndInstallDllEnvironment = retCode
End Function
