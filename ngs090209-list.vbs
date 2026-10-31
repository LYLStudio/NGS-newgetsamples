On Error Resume Next

Const HKEY_CLASSES_ROOT	= &H80000000
Const HKEY_CURRENT_USER	= &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG	= &H80000005
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7

Dim intReturn
Dim specFileName
Dim fnInfo

wmiReturnDec = Array("Success"_
										,"Not Supported"_
										,"Access Denied"_
										,"Dependent Services Running"_
										,"Invalid Service Control"_
										,"Service Cannot Accept Control"_
										,"Service Not Active"_
										,"Service Request Timeout"_
										,"Unknown Failure"_
										,"Path Not Found"_
										,"Service Already Running"_
										,"Service Database Locked"_
										,"Service Dependency Deleted"_
										,"Service Dependency Failure"_
										,"Service Disabled"_
										,"Service Logon Failure"_
										,"Service Marked For Deletion"_
										,"Service No Thread"_
										,"Status Circular Dependency"_
										,"Status Duplicate Name"_
										,"Status Invalid Name"_
										,"Status Invalid Parameter"_
										,"Status Invalid Service Account"_
										,"Status Service Exists"_
										,"Service Already Paused")

ooxxinfo = MsgBox("注意" & vbCrLf & vbCrLf & _
									"為了保護您的重要資料，避免在操作過程中遺失或損毀" & vbCrLf & vbCrLf & _
									"請在使用此工具前先備份重要資料，並關閉其他使用中的應用程式!" & vbCrLf & vbCrLf & _
									"繼續執行，請按...[是]" & vbCrLf & vbCrLf & _
									"退出請按...............[否]",292,"Important!")

If ooxxinfo = 6 Then

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("PROCESS")
	Set objShell = CreateObject("Shell.Application")
	strDesktop = WshShell.SpecialFolders("Desktop")
	gvslog = strDesktop&"\gvsLog.txt"
		
	WriteTextFile gvsLog, ("**************Log File Start*************")
	WriteTextFile gvsLog, "NewGetSamples Tool by Liang Li-ying, 2007-2009"
	
	OS_INFO_KEYPATH = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
	vnProductName = "ProductName"
	vnCurrentVersion = "CurrentVersion"
	vnCurrentBuildNumber = "CurrentBuildNumber"
	vnCSDVersion = "CSDVersion"
	oReg.GetStringValue HKEY_LOCAL_MACHINE,OS_INFO_KEYPATH,vnProductName,vProductName
	oReg.GetStringValue HKEY_LOCAL_MACHINE,OS_INFO_KEYPATH,vnCurrentVersion,vCurrentVersion
	oReg.GetStringValue HKEY_LOCAL_MACHINE,OS_INFO_KEYPATH,vnCurrentBuildNumber,vCurrentBuildNumber
	oReg.GetStringValue HKEY_LOCAL_MACHINE,OS_INFO_KEYPATH,vnCSDVersion,vCSDVersion
	OS_Info = vProductName & " " & vCurrentVersion & "." & vCurrentBuildNumber & ", " & vCSDVersion
	
	WriteTextFile gvsLog, ( "OS_INFO: " & OS_Info)
	WriteTextFile gvsLog, ( "System Date Time: " & Now )

	Set colStartupCommands = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
	For Each objStartupCommand in colStartupCommands
	  BootupState = objStartupCommand.BootupState
	Next

	WriteTextFile gvsLog, ("Bootup State: " & BootupState)

	If BootupState = "Normal boot" Then
		bootmodecheck = MsgBox ("建議重新啟動電腦進入[安全模式]!!" & vbCrLf & vbCrLf & _
														"重新啟動電腦請選...[是]" & vbCrLf & vbCrLf & _
														"繼續執行請選...........[否]",292,"開機模式確認")
		If bootmodecheck = 6 Then
			WriteTextFile gvsLog, ("重新啟動電腦 = YES")
			WriteTextFile gvsLog, Now
			WriteTextFile gvsLog, ("**************Log File End***************")

			rebootPC

		Else
			WriteTextFile gvsLog, ("重新啟動電腦 = NO")
		End If
	End If

	ParentFolder = strDesktop&"\"
	
	DestinationDir= strDesktop&"\VirusSample"
	If objFSO.FolderExists(DestinationDir) = False Then
		objFSO.CreateFolder(DestinationDir)
	End If 
	DestinationDir=DestinationDir & "\"

	WriteTextFile gvsLog, ("Samples Save To: "&DestinationDir)
	WriteTextFile gvsLog, (" ")
	
	WriteTextFile gvsLog, ("==========[Check IFEO1]")
	delIFEO
	WriteTextFile gvsLog, ""
	
	WriteTextFile gvsLog, ("==========[Stop Process]")
	StopProcess
	WriteTextFile gvsLog, ""
	
	WriteTextFile gvsLog, ("==========[Check IFEO2]")
	delIFEO
	WriteTextFile gvsLog, ""

	WriteTextFile gvsLog, ("==========[List Processing]")
	If objFSO.FileExists(WshShell.ExpandEnvironmentStrings("%temp%\List.txt")) Then
		Set listTXT = objFSO.OpenTextFile (WshShell.ExpandEnvironmentStrings("%temp%\List.txt"), 1)
		Do Until listTXT.AtEndOfStream
			lineValue = listTXT.readline
			arrLineValue = Split(lineValue,"*")

			funcName = arrLineValue(0)

			Select Case arrLineValue(0)

				Case "[StopService]"
					serviceNAME = arrLineValue(1)
					StopService serviceNAME
			
				Case "[DelService]"
					specFileName = 	WshShell.ExpandEnvironmentStrings(arrLineValue(2))
					serviceNAME = arrLineValue(1)
					objFSO.DeleteFile specFileName & "*ngs" , True
					DelService serviceNAME, specFileName
	
				Case "[DelDrivers]"
					specFileName = WshShell.ExpandEnvironmentStrings(arrLineValue(2))
					driverNAME = arrLineValue(1)
					objFSO.DeleteFile specFileName & "*ngs" , True
					DelDrivers driverNAME, specFileName

				Case "[DelRegValue]"
					specFileName = WshShell.ExpandEnvironmentStrings(arrLineValue(3))
					objFSO.DeleteFile specFileName & "*ngs" , True
					DelRegValue arrLineValue(1), arrLineValue(2), specFileName

				Case "[ModRegValue]"
					ModRegValue

				Case "[DelSpecFiles]"
					specFileName = WshShell.ExpandEnvironmentStrings(arrLineValue(1))
					objFSO.DeleteFile specFileName & "*ngs" , True
					
					If objFSO.FileExists(specFileName) = True Then
						msg_funcStarup = "StartUp " & funcName & ": " '& vbtab
						WriteTextFile gvsLog, msg_funcStarup
						DelSpecFiles specFileName
					End If 
				
				Case Else
					WriteTextFile gvsLog, ("Unknown list format->" & lineValue)

			End Select
		loop

	Else
		WriteTextFile gvsLog, ("No List!!!!")
	End If
	WriteTextFile gvsLog, ""	
	
	WriteTextFile gvsLog, ("==========[Search ShellExecuteHooks & ShellServiceObjectDelayLoad]")
	delShellExecuteHooks
	delShellServiceObjectDelayLoad
	WriteTextFile gvsLog, ""

	'WriteTextFile gvsLog, ("[CleanTemp]")
	'CleanTemp
	WriteTextFile gvsLog, ("==========[SearchAutorun]")
	DelAutorun
	WriteTextFile gvsLog, ""
	
	showAll

	WriteTextFile gvsLog, ("==========[RestoreHOSTS]")
	
	resetHOSTS = MsgBox("是否要重設HOSTS檔案為預設值?",292,"重設確認")
	If resetHOSTS = 6 Then
		WriteTextFile gvsLog, ("重設HOSTS = YES")
		restoreHOSTS
	Else
		WriteTextFile gvsLog, ("重設HOSTS = NO")
	End If
	WriteTextFile gvsLog, ""

	VirusSampleRAR = Chr(34)+strDesktop+"\VirusSample.rar"+Chr(34)
	VirusSampleTarget = Chr(34)+DestinationDir+"*.*"+Chr(34)
	GVSlogLoc = Chr(34) & gvslog & Chr(34)
	AutoScanLogLoc = Chr(34)+ParentFolder+"kavAutoScan.log"+Chr(34)
	rar = WshShell.ExpandEnvironmentStrings("%temp%\rar.exe")
	rarCMD = Chr(34) + rar + Chr(34) + " a -pvirus -ep "+VirusSampleRAR+" "+VirusSampleTarget+" "+GVSlogLoc+" "+AutoScanLogLoc

	'WriteTextFile gvsLog, ("壓縮命令列: " & rarCMD)

	'autoscan = MsgBox("是否要呼叫卡巴斯基進行掃瞄?",292,"掃瞄確認")
	'If autoscan = 6 Then
		'WriteTextFile gvsLog, ("呼叫AVP進行掃瞄 = Yes")
		'**********
		'kavAutoSacn
		'**********
	'Else
		'WriteTextFile gvsLog, ("呼叫AVP進行掃瞄 = NO")
	'End If
	
	CreateObject("WScript.Shell").Run rarCMD, 1, True
	
	
	WriteTextFile gvsLog, "==========[Check Current Registry Value...]"
	ListValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run"
	ListValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run"
	ListValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"
	ListValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks"
	ListValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
	
	
	
	
	restart = MsgBox("在正常模式下，由於作業系統本身的限制" & vbCrLf & vbCrLf & _
									 "部份檔案需重新啟動電腦後才能完整刪除" & vbCrLf & vbCrLf & _
									 "請您立即重新啟動電腦，並再執行一次此工具！" & vbCrLf & vbCrLf & _
									 "安全模式僅需執行一次即可！" & vbCrLf & vbCrLf & _
									 "是否立即重新啟動電腦？",292,"重新啟動系統確認")

	If restart = 6 Then
		WriteTextFile gvsLog, ("重新啟動電腦 = YES")
		WriteTextFile gvsLog, Now
		WriteTextFile gvsLog, ("**************Log File End***************")
		rebootPC
	Else
		WriteTextFile gvsLog, ("重新啟動電腦 = NO")
		WriteTextFile gvsLog, Now
		WriteTextFile gvsLog, ("**************Log File End***************")
	End If

End If

'******************************************************************************

Sub StopProcess
	On Error Resume Next
	msg_funcStarup = "StartUp " & funcName & ": " '& vbtab

	strPQuery = "Select * from Win32_Process" ' Where Not Name = 'NGS1xList.exe' And Not Name = 'winlogon.exe'"
	Set colProcessList = objWMIService.ExecQuery(strPQuery)
	For Each objProcess in colProcessList
		If lcase(objProcess.Name) <> lcase("NGS1xList.exe") And lcase(objProcess.Name) <> lcase("winlogon.exe") then
			intReturn = objProcess.Terminate
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" " & objProcess.Name & wmi_errResult)
		End If
	Next
	
End Sub

'******************************************************************************

Sub StopService(svcName)

	On Error Resume Next
	msg_funcStarup = "StartUp " & funcName & ": " '& vbtab
		
	strSQuery = "Select * from Win32_Service" ' Where Name = '" & svcName & "'"
	Set colListOfServices = objWMIService.ExecQuery(strSQuery)
	For Each objService in colListOfServices
		If LCase(objService.Name) = LCase(svcName) Then
			WriteTextFile gvsLog, msg_funcStarup
			intReturn = objService.StopService
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Stop Service->" & svcName & wmi_errResult)

			intReturn = objService.ChangeStartMode("Disabled")
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Disable Service->" & svcName & wmi_errResult)
			
			exit for
			
		End If
	Next
	
End Sub

'******************************************************************************

Sub DelService(svcName , fileName)
	On Error Resume Next
	msg_funcStarup = "StartUp " & funcName & ": " '& vbtab

	strSQuery = "Select * from Win32_Service" 'Where Name = '" & svcName & "'"
	Set colListOfServices = objWMIService.ExecQuery(strSQuery)
	For Each objService in colListOfServices
		If LCase(objService.Name) = LCase(svcName) Then
			WriteTextFile gvsLog, msg_funcStarup
			
			intReturn = objService.StopService
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Stop Service->" & svcName & wmi_errResult)

			intReturn = objService.ChangeStartMode("Disabled")
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Disable Service->" & svcName & wmi_errResult)

			intReturn = objService.Delete
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Delete Service->" & svcName & wmi_errResult)
			
			If fileName <>"" Then
				DelSpecFiles fileName
			End If
			
			exit for
			
		End If
	Next

End Sub

'******************************************************************************

Sub DelDrivers(drvName, fileName)
	On Error Resume Next
	msg_funcStarup = "StartUp " & funcName & ": " '& vbtab

	strDQuery = "SELECT * FROM Win32_SystemDriver" ' Where Name ='"&drvName&"'"
	Set colInstances = objWMIService.ExecQuery(strDQuery)
	For Each objDriver in colInstances
		If LCase(objDriver.Name) = LCase(drvName) Then
			WriteTextFile gvsLog, msg_funcStarup
			intReturn = objDriver.StopService
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Stop Drive->" & drvName & wmi_errResult)

			intReturn = objDriver.ChangeStartMode("Disabled")
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Disable Drive->" & drvName & wmi_errResult)

			intReturn = objDriver.Delete
			wmi_errResult = "->" & intReturn & " " & wmiReturnDec(intReturn)
			WriteTextFile gvsLog, (" Delete Drive->" & drvName & wmi_errResult)
	
			If fileName <>"" Then
				DelSpecFiles fileName
			End If
			
			exit for
			
		End If
	Next
	
End Sub

'******************************************************************************
Sub DelRegValue(kp, vn, fileName)
	On Error Resume Next

	allLen = Len(kp)
	searchChar = "\"
	rootlen = InStr(kp,searchChar)
	keyRoot = Left(kp,rootlen-1)
	'MsgBox keyRoot
	pathlen = InStr(kp,searchChar)
	keypath = Right(kp,allLen-rootlen)
	'MsgBox "KEYROOT:: "&keyRoot&Chr(10)&"KEY PATH:: "&keypath&Chr(10)&"VALUE NAME:: "&valuename&Chr(10)&"FILE NAME:: "&filename
	Select Case keyRoot
		Case	"HKEY_CLASSES_ROOT"
			keyRoot = HKEY_CLASSES_ROOT
			keyRootView = "HKEY_CLASSES_ROOT"
		Case	"HKEY_CURRENT_USER"
			keyRoot = HKEY_CURRENT_USER
			keyRootView = "HKEY_CURRENT_USER"
		Case	"HKEY_LOCAL_MACHINE"
			keyRoot = HKEY_LOCAL_MACHINE
			keyRootView = "HKEY_LOCAL_MACHINE"
		Case	"HKEY_USERS"
			keyRoot = HKEY_USERS
			keyRootView = "HKEY_USERS"
		Case	"HKEY_CURRENT_CONFIG"
			keyRoot = HKEY_CURRENT_CONFIG
			keyRootView = "HKEY_CURRENT_CONFIG"
		'Case Else MsgBox "not yet"
	End Select
	
	
	oReg.EnumValues keyRoot, keypath, vns
	
	For each v in vns
		'MsgBox (vns(i)	 & vbCrLf & vn)
		If LCase(v) = LCase(vn) Then
			msg_funcStarup = "StartUp " & funcName & ": " '& vbtab
			WriteTextFile gvsLog, (msg_funcStarup)

			oReg.DeleteValue keyRoot, keypath, vn
			WriteTextFile gvsLog, ("DeleteRegValue: " & keyRootView & "\" & keypath & "\" & vn & "->" & Err.number & " " & Err.Description)
			Err.Clear
			specFileName = WshShell.ExpandEnvironmentStrings(fileName)
			If specFileName <>"" Then
				DelSpecFiles specFileName
			End If
			Exit For
		End If
	Next

End Sub

'******************************************************************************

Sub rebootPC
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate,(Shutdown)}!\\" & _
	        strComputer & "\root\cimv2")

	Set colOperatingSystems = objWMIService.ExecQuery _
	    ("Select * from Win32_OperatingSystem")

	For Each objOperatingSystem in colOperatingSystems
	    objOperatingSystem.Reboot()
	Next
End Sub

'******************************************************************************

Sub DelSpecFiles(fileName)
	On Error Resume Next
	'WScript.echo fileName
	fileName=WshShell.ExpandEnvironmentStrings(fileName)
	'WScript.echo fileName
	If objFSO.FileExists(fileName) = False Then
		WriteTextFile gvsLog, (" DeleteTarget:->" & fileName & " not exist!!")
	Else
		Set objFile = objFSO.GetFile(fileName)
		objFile.Attributes = 0
		objFSO.CopyFile fileName , DestinationDir, True
		eno = err.number
		ddfn=DestinationDir & fileName
		ddfn=WshShell.ExpandEnvironmentStrings(ddfn)
		If objFSO.FileExists(ddfn) = True Then
			Set objDDFN = objFSO.GetFile(ddfn)
			ddfnCtime=objDDFN.DateCreated
			WriteTextFile gvsLog, (" Copy->" & fileName & "->" & ddfn & "(" & ddfnCtime & ")->" & eno)
		Else
			WriteTextFile gvsLog, (" Copy->" & fileName & "->" & eno	)
		End If
		err.Clear
				
		objFSO.DeleteFile (fileName)
		eno = err.number
		WriteTextFile gvsLog, (" Delete->" & fileName & "->" & eno )
		err.Clear
		
		If objFSO.FileExists(fileName) Then
			ReNameFiles fileName
		End If
	End If

End Sub

'******************************************************************************

Sub ModRegValue
	On Error Resume Next
	msg_funcStarup = "StartUp " & funcName & ":" & vbtab
	WriteTextFile gvsLog, (msg_funcStarup)

	valuename = arrLineValue(2)
	newvalue = arrLineValue(3)
	allLen = Len(arrLineValue(1))
	searchChar = "\"
	rootlen = InStr(arrLineValue(1),searchChar)
	keyRootName = Left(arrLineValue(1),rootlen-1)
	'MsgBox keyRoot
	pathlen = InStr(arrLineValue(1),searchChar)
	keypath = Right(arrLineValue(1),allLen-rootlen)
	'MsgBox "KEYROOT:: "&keyRoot&Chr(10)&"KEY PATH:: "&keypath&Chr(10)&"VALUE NAME:: "&valuename&Chr(10)&"New Value:: "&newvalue

	Select Case keyRootName
		Case	"HKEY_CLASSES_ROOT"
			keyRoot = HKEY_CLASSES_ROOT
			keyRootView = "HKEY_CLASSES_ROOT"
		Case	"HKEY_CURRENT_USER"
			keyRoot = HKEY_CURRENT_USER
			keyRootView = "HKEY_CURRENT_USER"
		Case	"HKEY_LOCAL_MACHINE"
			keyRoot = HKEY_LOCAL_MACHINE
			keyRootView = "HKEY_LOCAL_MACHINE"
		Case	"HKEY_USERS"
			keyRoot = HKEY_USERS
			keyRootView = "HKEY_USERS"
		Case	"HKEY_CURRENT_CONFIG"
			keyRoot = HKEY_CURRENT_CONFIG
			keyRootView = "HKEY_CURRENT_CONFIG"
			'Case Else MsgBox "not yet"
	End Select

oReg.EnumValues keyRoot, keypath, arrValueNames, arrValueTypes

For i=0 To UBound(arrValueNames)
  'Wscript.Echo "Value Name: " & arrValueNames(i)
	If arrValueNames(i) = valuename Then
    Select Case arrValueTypes(i)
      Case REG_SZ
        Err.Raise oReg.SetStringValue (keyRoot,keypath,valuename,newvalue)
        WriteTextFile gvsLog, ("ModifyRegValue: "&keyRootView&"\"&keypath&"\"&valuename&"->"&newvalue &"->"&Err.number&" "&Err.Description)
				Err.Clear
      Case REG_EXPAND_SZ
        Err.Raise oReg.SetExpandedStringValue (keyRoot,keypath,valuename,newvalue)
				WriteTextFile gvsLog, ("ModifyRegValue: "&keyRootView&"\"&keypath&"\"&valuename&"->"&newvalue &"->"&Err.number&" "&Err.Description)
				Err.Clear
      'Case REG_BINARY
        'Wscript.Echo "Data Type: Binary"
        ''objLogFile.Writeline(Now&" "&keyRootName&"\"&keypath&"\VALUE NAME:: "&valuename&"->New Value:: "&newvalue)
        'Select Case keyRootName
        	'Case "HKEY_CLASSES_ROOT"
        		'rootname = "HKCR"
        	'Case "HKEY_CURRENT_USER"
        		'rootname = "HKCU"
        	'Case "HKEY_LOCAL_MACHINE"
        		'rootname = "HKLM"
        	'Case "HKEY_CURRENT_CONFIG"
        		'rootname = "HKEY_CURRENT_CONFIG"
        	'Case "HKEY_USERS"
        	 'rootname = "HKEY_USERS"
        'End Select
        'newvalue = CInt(newvalue)
        'newvalue = &1
        'MsgBox newvalue
        'regbinaryMod = rootname&"\"&keypath&"\"&valuename&"\"
        'Err.Raise WshShell.RegWrite (regbinaryMod, newvalue, "REG_BINARY")
				''objLogFile.Writeline(Now&" Reg ModifyValue #"&Err.number&" "&Err.Description)
				'Err.Clear

      Case REG_DWORD
        Err.Raise oReg.SetDWORDValue (keyRoot,keypath,valuename,newvalue)
        WriteTextFile gvsLog, ("ModifyRegValue: " & keyRootView & "\" & keypath & "\" & valuename & "->" & newvalue & "->"&Err.number & " " & Err.Description)
				Err.Clear

      Case REG_MULTI_SZ
        Err.Raise oReg.SetMultiStringValue (keyRoot,keypath,valuename,newvalue)
				WriteTextFile gvsLog, ("ModifyRegValue: " & keyRootView & "\" & keypath & "\" & valuename & "->" & newvalue & "->" & Err.number & " " & Err.Description)
				Err.Clear
	  End Select
  	Exit For
  End If
Next

End Sub

'******************************************************************************

Sub kavAutoSacn

	strKeyPath = "SYSTEM\CurrentControlSet\Services\AVP"
	strValueName = "ImagePath"
	oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	kav=strValue

	If IsNull(kav) Then
		MsgBox ("No Kaspersky!!")
		WriteTextFile gvsLog, ("NO Kaspersky!!")

	Else
		strKavLoc = kav
		searchCharEnd = "e"
		searchCharEndLen = InStrRev(strKavLoc, searchCharEnd)
		kavCMDtemp = Left(strKavLoc, searchCharEndLen-3)
		searchCharStart = Chr(34)
		searchCharStartLen =InStr(kavCMDtemp,searchCharStart)
		If searchCharStartLen = 0 Then
			kavCMD = Chr(34)&kavCMDtemp&"com"&Chr(34)
		Else
			If searchCharStartLen = 1 Then
				kavCMD = kavCMDtemp & "com" & Chr(34)
			End If
		End If

		logloc = Chr(34)+ParentFolder&"kavAutoScan.log"+Chr(34)
		scancmd = kavCMD & " scan /startup /ra:"&logloc&" /i3 /fa /ichecker=off /iswift=off"
		CreateObject("WScript.Shell").Run scancmd,1, True
		WriteTextFile gvsLog, ("Auto Scan Log: "&logloc)
	End If

End Sub

'******************************************************************************

Sub DelAutorun
	On Error Resume Next
	fname = "autorun.inf"
	Set colDisks = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk")
	For Each objDisk in colDisks

		If (objDisk.DriveType = 2) Or (objDisk.DriveType = 3) Or (objDisk.DriveType = 6) Then

			autorun = objDisk.deviceID & "\" & fname
			If objFSO.FileExists(autorun) Then
				Set objFile = objFSO.GetFile(autorun)
				WriteTextFile gvsLog, (" " & autorun & " Found!")
				objFile.Attributes = 0
				Set ooxx = objFSO.OpenTextFile (autorun, 1)
				Do Until ooxx.AtEndOfStream
					strExecute = ooxx.Readline
					strExecuteLen = Len(strExecute)
					SearchChar = "="
					strSearchCharLen = Instr(strExecute, SearchChar)
					targetFileName = Right(strExecute, (strExecuteLen-strSearchCharLen))

					virus= objDisk.deviceID & "\" & targetFileName

					If objFSO.FileExists(virus) Then
						WriteTextFile gvsLog, (" " & virus & " Found!")

						Set objFileInside = objFSO.GetFile(virus)
						objFileInside.Attributes = 0
						DelSpecFiles virus
					End If
				Loop
				ooxx.Close
				DelSpecFiles autorun
			End If
		End If
	Next
	
End Sub

'******************************************************************************

Sub delIFEO
	On Error Resume Next
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
	oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	For Each subkey In arrSubKeys
	    intErr = oReg.DeleteKey (HKEY_LOCAL_MACHINE, strKeyPath&"\"&subkey)
	    WriteTextFile gvsLog, (" Delete IFEO:" & subkey & "->" & intErr) 
	Next
End Sub

'******************************************************************************

Sub showAll
	On Error Resume Next
	dwValue = 1
	strKeyPath1 =	"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
	strKeyPath2 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\SHOWALL"
	
	strValueName1 = "Hidden"
	strValueName2 = "SuperHidden"
	strValueName3 = "CheckedValue"
	
	oReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath1,strValueName1,dwValue
	oReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath1,strValueName2,dwValue
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath2,strValueName3,dwValue
End Sub

'******************************************************************************

Sub RestoreHOSTS
	On Error Resume Next
	strKeyPath = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
	strValueName = "DataBasePath"
	oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	rootHOSTS = strValue
	fileHOSTS = rootHOSTS&"\hosts"
	msg = " HOSTS File: " & fileHOSTS
	WriteTextFile gvsLog, msg
	fileREADER fileHOSTS
	'MsgBox fileHOSTS
	If objFSO.FileExists(fileHOSTS) = True Then
		objFSO.DeleteFile(fileHOSTS)
	End If
	msg = "127.0.0.1		localhost"
	WriteTextFile fileHOSTS, msg
	msg = "::1		localhost"
	WriteTextFile fileHOSTS, msg

End Sub

'******************************************************************************

Sub CleanTemp
	On error resume next
	oReg.GetStringValue HKEY_CURRENT_USER,"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders","Cache",CachePath

	winTEMP = WshSysEnv("windir")&"\TEMP"
	userTEMP = WshSysEnv("temp")
	ieTEMP = CachePath

	temp=Array(winTEMP,userTEMP,ieTEMP)

	For t = 0 To UBound(temp)
		objStartFolder = temp(t)
		Set objFolder = objFSO.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			objFile.Attributes = 0
			filespec  = objFile.path
			objFSO.DeleteFile filespec, True '))
		Next
		ShowSubfolders objFSO.GetFolder(objStartFolder)
	Next
End sub

'******************************************************************************

Sub ShowSubFolders(Folder)
	On Error Resume Next
	For Each Subfolder in Folder.SubFolders
		Set objFolder = objFSO.GetFolder(Subfolder.Path)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			objFile.Attributes = 0
			filespec1 = objFile.path
			objFSO.DeleteFile filespec1, True '))
		Next
		ShowSubFolders Subfolder
	Next
End Sub

'******************************************************************************

Sub delShellExecuteHooks
	On Error Resume Next
	SEHsKP = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks"
	oReg.EnumValues HKEY_LOCAL_MACHINE, SEHsKP, SEHaVN

	For i=0 To UBound(SEHaVN)
	
		If (SEHaVN(i) <> "{AEB6717E-7E19-11d0-97EE-00C04FD91972}") And (SEHaVN(i) <> "{B5A7F190-DDA6-4420-B3BA-52453494E6CD}") And (SEHaVN(i) <> "{56F9679E-7826-4C84-81F3-532071A8BCC5}") Then
			intReturn = oReg.DeleteValue (HKEY_LOCAL_MACHINE, SEHsKP, SEHaVN(i))
			WriteTextFile gvsLog, (" DelShellExecuteHooks: " & SEHaVN(i) & "->" & intReturn)
			'MsgBox intreturn
			'MsgBox SEHaVN(i)

			fnInfoSEHsKP = "CLSID\" & SEHaVN(i) & "\InProcServer32"
			ListValue HKEY_CLASSES_ROOT, fnInfoSEHsKP
			
			fnLen = InStrRev(fnInfo,".")
			fnExec = Left(fnInfo,fnLen) & "exe"

			specFileName = WshShell.ExpandEnvironmentStrings(fnExec)
			objFSO.DeleteFile specFileName & "*ngs" , True
			DelSpecFiles specFileName

			specFileName = WshShell.ExpandEnvironmentStrings(fnInfo)
			objFSO.DeleteFile specFileName & "*ngs" , True
			DelSpecFiles specFileName
		End If
	Next
	
End Sub

Sub delShellServiceObjectDelayLoad
	SSODLPK = "SOFTWARE\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
	oReg.EnumValues HKEY_LOCAL_MACHINE, SSODLPK, SSODLVN
	For i=0 To UBound(SSODLVN)
		If (SSODLVN(i) <> "CDBurn") And (SSODLVN(i) <> "PostBootReminder") And (SSODLVN(i) <> "SysTray") And (SSODLVN(i) <> "WebCheck") And (SSODLVN(i) <> "WPDShServiceObj") And (SSODLVN(i) <> "UPnPMonitor") Then
			oReg.GetStringValue HKEY_LOCAL_MACHINE, SSODLPK, SSODLVN(i), strValue
			fnInfoSSODLPK = "CLSID\" & strValue & "\InProcServer32"
			ListValue HKEY_CLASSES_ROOT, fnInfoSSODLPK
			fnLen = InStrRev(fnInfo,".")
			fnExec = Left(fnInfo,fnLen) & "exe"
			intReturn = oReg.DeleteValue (HKEY_LOCAL_MACHINE, SSODLPK, SSODLVN(i))
			WriteTextFile gvsLog, (" DelShellServiceObjectDelayLoad: " & SSODLVN(i) & "->" & intReturn)
			
			specFileName = WshShell.ExpandEnvironmentStrings(fnExec)
			objFSO.DeleteFile specFileName & "*.ngs" , True
			DelSpecFiles specFileName
			
			specFileName = WshShell.ExpandEnvironmentStrings(fnInfo)
			objFSO.DeleteFile specFileName & "*.ngs" , True
			DelSpecFiles specFileName
			
		End If
	Next

End Sub


'******************************************************************************
	
Sub ListValue(RegRoot,keyInfo)
	On Error Resume Next
	oReg.EnumValues RegRoot, keyInfo, arrValueNames, arrValueTypes
	oReg.GetStringValue RegRoot, keyInfo, arrValueNames(0), strValue
	fnInfo = strValue
End Sub

'******************************************************************************

Sub RenameFiles(fileName)
	On Error Resume Next
	search = chr(92)
	replacechr = chr(92)+chr(92)
	reName = replace(fileName, search, replacechr)

	Set colFiles = objWMIService.ExecQuery ("Select * from Cim_Datafile where Name = '" & reName & "'")

	For Each objFile in colFiles
		errResult = objFile.Rename(Rename & ".ngs")
		reNamed = reName & ".ngs"
		
		WriteTextFile gvsLog, " " & "ReName->" & reNamed & "->" & errResult
	Next

End Sub

'******************************************************************************

Sub fileREADER(readTarget)
	On Error Resume Next

	If objFSO.FileExists(readTarget) = True Then
		Set objReadTarget = objFSO.OpenTextFile(readTarget, 1)
		Do Until objReadTarget.AtEndOfStream 
			lineValue = objReadTarget.Readline

			If Left(lineValue,1) <> "#" Then 
				WriteTextFile gvsLog, lineValue
			End If
		loop
	End If
End Sub 

'******************************************************************************

Sub ListValue(RegRoot,regKeyPath)
	On Error Resume Next
	Select Case RegRoot
		Case HKEY_CLASSES_ROOT
			RegRootView = "HKEY_CLASSES_ROOT"
		Case HKEY_CURRENT_USER
			RegRootView = "HKEY_CURRENT_USER"
		Case HKEY_LOCAL_MACHINE
			RegRootView = "HKEY_LOCAL_MACHINE"
		Case HKEY_USERS
			RegRootView = "HKEY_USERS"
		Case HKEY_CURRENT_CONFIG
			RegRootView = "HKEY_CURRENT_CONFIG"
	End Select
		
	oReg.EnumValues regRoot, regKeyPath, arrValueNames, arrValueTypes
	WriteTextFile gvsLog, RegRootView & "\" & regKeyPath
	
	If IsNull(arrValueNames) Then
		'WriteTextFile gvsLog, "no value" & vbCrLf
		Exit Sub
	End If

	For i=0 To UBound(arrValueNames)	
		Select Case arrValueTypes(i)
			Case REG_SZ
				'WriteTextFile gvsLog, "*REG_SZ*"
				oReg.GetStringValue regRoot, regKeyPath, arrValueNames(i), regValue
				WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" &  regValue
	
			Case REG_EXPAND_SZ
		    'WriteTextFile gvsLog, "*REG_EXPAND_SZ*"
				oReg.GetExpandedStringValue regRoot, regKeyPath, arrValueNames(i), regValue
				WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" &  regValue
	
			Case REG_BINARY
		    'WriteTextFile gvsLog, "*REG_BINARY*"
		    oReg.GetBinaryValue regRoot, regKeyPath, arrValueNames(i), regValue
				For h = lBound(regValue) to uBound(regValue)
					ValueProperty = ValueProperty & vbCrLf & Hex(regValue(h))
				Next
				WriteTextFile gvsLog, " " & arrValueNames(i) & vbCrLf & ValueProperty

			Case REG_DWORD
		    'WriteTextFile gvsLog, "*REG_DWORD*"
		    oReg.GetDWORDValue regRoot, regKeyPath, arrValueNames(i), regValue
				WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" & Hex(regValue) & "(H)"
			
			Case REG_MULTI_SZ
		    'WriteTextFile gvsLog, "*REG_MULTI_SZ*"
		    oReg.GetMultiStringValue  regRoot, regKeyPath, arrValueNames(i), regValue
				For Each strValue In regValue
				    WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" & strValue
				Next
				
	    End Select
	Next
	WriteTextFile gvsLog, ""
End Sub

'=============================================================================================================================

Function ListValueOne(RegRoot,regKeyPath,regValueName)
	On Error Resume Next
	Select Case RegRoot
		Case HKEY_CLASSES_ROOT
			RegRootView = "HKEY_CLASSES_ROOT"
		Case HKEY_CURRENT_USER
			RegRootView = "HKEY_CURRENT_USER"
		Case HKEY_LOCAL_MACHINE
			RegRootView = "HKEY_LOCAL_MACHINE"
		Case HKEY_USERS
			RegRootView = "HKEY_USERS"
		Case HKEY_CURRENT_CONFIG
			RegRootView = "HKEY_CURRENT_CONFIG"
	End Select
		
	objReg.EnumValues regRoot, regKeyPath, arrValueNames, arrValueTypes
	WriteTextFile gvsLog, RegRootView & "\" & regKeyPath
	
	If IsNull(arrValueNames) Then
		WriteTextFile gvsLog, "no value"
		Exit Function
	End If

	For i=0 To UBound(arrValueNames)
		If arrValueNames(i) = regValueName Then
				
			Select Case arrValueTypes(i)
				Case REG_SZ
					'WriteTextFile gvsLog, "*REG_SZ*"
					objReg.GetStringValue regRoot, regKeyPath, arrValueNames(i), regValue
					WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" &  regValue
					ListValueOne = regValue
				Case REG_EXPAND_SZ
			    'WriteTextFile gvsLog, "*REG_EXPAND_SZ*"
					objReg.GetExpandedStringValue regRoot, regKeyPath, arrValueNames(i), regValue
					WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" &  regValue
					ListValueOne = regValue
				Case REG_BINARY
			    'WriteTextFile gvsLog, "*REG_BINARY*"
			    objReg.GetBinaryValue regRoot, regKeyPath, arrValueNames(i), regValue
					For h = lBound(regValue) to uBound(regValue)
						ValueProperty = ValueProperty & vbCrLf & Hex(regValue(h))
					Next
					WriteTextFile gvsLog, " " & arrValueNames(i) & vbCrLf & ValueProperty
	
				Case REG_DWORD
			    'WriteTextFile gvsLog, "*REG_DWORD*"
			    objReg.GetDWORDValue regRoot, regKeyPath, arrValueNames(i), regValue
					WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" & Hex(regValue) & "(H)"
				
				Case REG_MULTI_SZ
			    'WriteTextFile gvsLog, "*REG_MULTI_SZ*"
			    objReg.GetMultiStringValue  regRoot, regKeyPath, arrValueNames(i), regValue
					For Each strValue In regValue
					    WriteTextFile gvsLog, " " & arrValueNames(i) & "|*|" & strValue
					Next
		  End Select
		End If
	Next
	WriteTextFile gvsLog, ""
End Function

'=============================================================================================================================

Sub WriteTextFile(strFileName, strOutput)
	On Error Resume Next
	Const FOR_APPENDING = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFileName) Then
	  Set objTextStream = objFSO.OpenTextFile(strFileName, FOR_APPENDING)
	Else
	  Set objTextStream = objFSO.CreateTextFile(strFileName)
	End If
	objTextStream.WriteLine strOutput
	objTextStream.Close
End Sub
