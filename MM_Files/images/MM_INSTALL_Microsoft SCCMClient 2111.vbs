'==========================================================================
' DO NOT CHANGE THIS SECTION
' NAME: MassMutual_Install_Template 1.0.0.vbs
' TEMPLATE AUTHOR: Goutham M
' DATE: 11/20/2019
'==========================================================================
' REVISION HISTORY
'==========================================================================

'==========================================================================



Option Explicit
On Error Resume Next
Dim Product,RunType,bErrorFound,ResultCode,SaveResultCodes,IgnoreError,WaitForFinish,ScriptPath,LogPath,LogFile,strOS,strProcArch,objNetwork,BuildLogFile,Start_date,strUserName,Publisher,AppName,AppVersion,RegaccessValue
Dim oShell,windir,ProgramData,ProgramFiles,ProgramFilesX86,ComSpec,SystemDrive,CommonProgramFiles,CommonProgramFilesX86,ProperGUID,BuildLogFileProg,oemPath,objRegistry,strCheck,strCheck1

Dim sScriptName,iErrorCode,WshNetwork,oDrives,sDriveLetter,i
Const HKCR=&H80000000
Const HKLM=&H80000002
Start_date = Date & " " & Time
Dim strComputer: strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
'========================
'Admin Rights Elevation
'========================
'Get Script Name
sScriptName = WScript.ScriptFullName

'check if admin rights
Set oShell = CreateObject("WScript.Shell")
iErrorCode = oShell.run("cmd.exe /C >nul 2>&1 ""%SYSTEMROOT%\system32\cacls.exe"" ""%SYSTEMROOT%\system32\config\system""",,True)
'Error Code is 5 if the script isn't elevated
If iErrorCode <> 0 Then
    'Check if there is a need to dereference mapped drive
    sDriveLetter = Left(sScriptName,2)
    If sDriveLetter <> "\\" or sDriveLetter <> "C:" Then
        Set WshNetwork = WScript.CreateObject("WScript.Network")
        Set oDrives = WshNetwork.EnumNetworkDrives
        For i = 0 to oDrives.Count - 1 Step 2
            If oDrives(i) = sDriveLetter Then
                sScriptName = Replace(sScriptName,sDriveLetter,oDrives(i+1))
            End if
        Next
    End if
    Set oShell = CreateObject("Shell.Application")
    'launch application prompting for admin rights
    oShell.ShellExecute WScript.FullName,chr(34) & sScriptName & chr(34),,"runas",1
    'quit the non-elevated script in favor of the elevated one
    WScript.Quit
End if


Set oShell = WScript.CreateObject("WScript.Shell")
windir = oShell.ExpandEnvironmentStrings("%windir%") 'C:\Windows
ProgramData = oShell.ExpandEnvironmentStrings("%ProgramData%") 'C:\ProgramData
ProgramFiles = oShell.ExpandEnvironmentStrings("%ProgramFiles%") 'C:\Program Files
ProgramFilesX86 = oShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") 'C:\Program Files (x86)
ComSpec = oShell.ExpandEnvironmentStrings("%ComSpec%") 'C:\Windows\system32\cmd.exe
SystemDrive = oShell.ExpandEnvironmentStrings("%systemdrive%") 'C:
CommonProgramFiles = oShell.ExpandEnvironmentStrings("%CommonProgramFiles%") 'C:\Program Files\Common Files
CommonProgramFilesX86 = oShell.ExpandEnvironmentStrings("%CommonProgramFiles(x86)%") 'C:\Program Files (x86)\Common Files
strUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
BuildLogFile = SystemDrive & "\Build.txt"
BuildLogFileProg = ProgramFilesX86 & "\sms\Build\Build.txt"

'--------------------------------------------------------------------------------------------------------------------------------

' -------- CHANGE THIS SECTION (1) --------------------

' AUTHOR: Maliha Azam' Packager Name
' DATE  : 02/01/2022

Product = "Microsoft SCCMClient 2111" ' Product name and version
RunType = "Install" ' Install/Uninstall/Configuration
Publisher = "Microsoft"
AppName = "SCCMClient"
AppVersion = "2111"

Permission Chr(34) & ProgramFilesX86 & "\sms" & Chr(34)
InitialRun ' Do not change this line
RegaccessValue = Regaccess

' Example: RunCommandLine "msiexec /i """ & oempath & "FIS_OmniPayDB_NextGen1.0.msi"" /qb /L*V """ & LogPath & "FIS_OmniPayDB_NextGen1.0_Inst.log"""
' Example: RunCommandLine "msiexec /i " & Chr(34) & oempath & "Test.msi" & Chr(34) & " TRANSFORMS=" & Chr(34) & oempath & "Test.mst" & Chr(34) & " /qn /L*V """ & LogPath & "Test_Inst.log"""
' Example: RunCommandLine "msiexec /x ""{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"" /qn /L*V """ & LogPath & "Test_Uninst.log"""
' Example: RunCommandLine Chr(34) & oemPath & "7z1805-x64.exe" & Chr(34) & " /S" 
' Example: RunCommandLine Chr(34) & ProgramFiles & "\7-Zip\Uninstall.exe" & Chr(34) & " /S"
' Example: Permission Chr(34) & ProgramFilesX86 & "\Java" & Chr(34) 'Apply full permission to Users
' Example: RemoveOrphanFolder ProgramFilesX86 & "\Java"
' Example: RemoveOrphanFile ProgramFilesX86 & "\Java\README.txt"
' Example: RemoveFolderOnlyIfEmpty ProgramFilesX86 & "\Test"
' Example: RegAddDWORD HKLM,"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{26A24AE4-039D-4CA4-87B4-2F32180144F0}","NoModify",00000001
' Example: RegAddSTRING HKLM,"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{26A24AE4-039D-4CA4-87B4-2F32180144F0}","DisplayName","Java 8 Update 144"
' Example: DeleteSubkeys HKLM,"SOFTWARE\JavaSoft"
' Example: DeleteOrphanRegistry HKLM,"SOFTWARE\JavaSoft"
' Example: DeleteOrphanRegistryOnlyIfEmpty HKLM,"SOFTWARE\JavaSoft"
' Example: DeleteService "SnowInventoryAgent5"
' Example: CopyFile ProgramFilesX86 & "\Java\Test.info",LogPath
' Example: CopyFolder ProgramFilesX86 & "\Test",LogPath
' Example: UninstallMSI "Snow"
' Example: KillProcess "notepad.exe"
' Example: popupforuserloggedin "Test.exe"
' Example: MM_Branding 32,"\\iscapmm1\Prod\VB,ISM,MSTTemplates\VBTemplateFinal18Dec" (or) MM_Branding 64,"\\iscapmm1\Prod\VB,ISM,MSTTemplates\VBTemplateFinal18Dec"   ' 32 for WoW6432node
' Example: Delete_MM_Branding 32 (or) Delete_MM_Branding 64   ' 32 for WoW6432node
'======================================================================================
'Provide the CommandLine below

Dim arch : arch = GetProcessorArchitecture()
RunCommandLine Chr(34) & oemPath & "wmi.bat" & Chr(34)
RunCommandLine Chr(34) & oemPath & "ccmsetup.exe" & Chr(34) & " /skipprereq:ndp452-kb2901907-x86-x64-allos-enu.exe /service SMSSITECODE=MMF SMSCACHESIZE=16000 CCMLOGLEVEL=0 CCMLOGMAXSIZE=51200000 DISABLECACHEOPT=TRUE RESETKEYINFORMATION=TRUE FSP=ISSCCMFBPR10.na.mmfg.net"
Wscript.sleep 600000
CheckInstalled

If arch="x86" Then
	MM_Branding 32,"\\iscapmm1\Prod\Microsoft\System Center Configuration Manager Client\2111"
	DeleteSubkeys HKLM, "SOFTWARE\WOW6432Node\Massmutual\Microsoft_SCCMClient_2002" 
	DeleteSubkeys HKLM, "SOFTWARE\WOW6432Node\Massmutual\Microsoft_SCCMClient_2006"
	DeleteSubkeys HKLM, "SOFTWARE\WOW6432Node\Massmutual\Microsoft_SCCMClient_2010"
Else
	MM_Branding 64,"\\iscapmm1\Prod\Microsoft\System Center Configuration Manager Client\2111"
	DeleteSubkeys HKLM, "SOFTWARE\Massmutual\Microsoft_SCCMClient_2002"
	DeleteSubkeys HKLM, "SOFTWARE\Massmutual\Microsoft_SCCMClient_2006"
	DeleteSubkeys HKLM, "SOFTWARE\Massmutual\Microsoft_SCCMClient_2010"
End If

' -------- END OF CHANGES (1) --------------------------


'msgbox RegaccessValue
'BuildLogWrite

CloseOut
'WScript.Quit ResultCode
'--------Main Functions starts here!!!------------------
Sub InitialRun
Dim TemplateVersion
Dim oFSO,oShell,windir
TemplateVersion = "1.0.0"
strProcArch = GetProcessorArchitecture()
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject("WScript.Shell")
windir = oShell.ExpandEnvironmentStrings("%windir%")
LogPath = ProgramFilesX86 & "\sms"
If Not oFSO.FolderExists (LogPath) Then oFSO.CreateFolder LogPath
LogPath = LogPath & "\Log\"
if Not oFSO.FolderExists (LogPath) Then oFSO.CreateFolder LogPath
LogPath = LogPath & "\" & Publisher & "\" 
if Not oFSO.FolderExists (LogPath) Then oFSO.CreateFolder LogPath
LogPath = LogPath & "\" & AppName & "\"
if Not oFSO.FolderExists (LogPath) Then oFSO.CreateFolder LogPath
LogPath = LogPath & "\" & AppVersion & "\"
if Not oFSO.FolderExists (LogPath) Then oFSO.CreateFolder LogPath
LogFile = LogPath & Product & " _Overview.log"



LogWrite "===================================================="
LogWrite "Template Version: " & TemplateVersion
LogWrite "Run Type: " & RunType
strOS = GetOSFamily()
LogWrite "Processor family: " & strProcArch
  GetMachineName()
LogWrite "Current Logged on User Name: " & strUserName
ScriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, "\"))
oempath = ScriptPath & "Source\oem\"
'msgbox oempath
CheckForScriptError "Looking up script location"
oShell.CurrentDirectory = ScriptPath
CheckForScriptError "Setting current directory"
LogWrite "Script path: " & ScriptPath
CopyFile oemPath & "loggedon.exe" , ProgramFilesX86 & "\sms\"
bErrorFound = False
ResultCode = 0
SaveResultCodes = ""
IgnoreError = False
WaitForFinish = True 

Set oShell = Nothing

End Sub


Sub RunCommandLine(CommandLine)
Dim bStepSuccess
Dim oShell

LogWrite "Looking for executable from command line: " & CommandLine
If ExecutableFound (CommandLine) = False Then 
bErrorFound = True
ResultCode = 9990
ProcessError ResultCode
End If

LogWrite "Launching command line (Wait for finish = " & WaitForFinish & ")"
LogWrite "Command Line: " & CommandLine
Set oShell = CreateObject("Wscript.Shell")
ResultCode = oShell.Run (CommandLine,0,WaitForFinish)

bStepSuccess = False
CheckforExitCode ResultCode, bStepSuccess
CheckStepResult bStepSuccess

Set oShell = Nothing
IgnoreError = False
WaitForFinish = True
End Sub

Sub CheckforExitCode(ResultCode, bStepSuccess)
if WaitForFinish then
Select Case ResultCode
Case 0
bStepSuccess = True
LogWrite "Exit Code:" & ResultCode
LogWrite "Action completed successfully"
Case 1601
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "The Windows Installer service could not be accessed"
Case 1602
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "User cancel installation"
Case 1603
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "Fatal error during installation"
Case 1604
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "Installation suspended, Incomplete"
Case 1608
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "Unknown property" 
Case 1612
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "The installation source for this product is not available. Verify that the source exists and that you can access it" 
Case 1618
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "Another installation is already in Progress!!!"
Case 1619
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "This installation package could not be opened. Verify that the package exists and is accessible"
Case 1620
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "This installation package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer package"
Case 1622
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "Error opening installation log file. Verify that the specified log file location exists and is writable"
Case 1624
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "Error applying transforms. Verify that the specified transform paths are valid"
Case 1638
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "Another version of this product is already installed. Installation of this version cannot continue. To configure or remove the existing version of this product, use Add/Remove Programs on the Control Panel"
Case 1633
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "This installation package is not supported on this platform"
Case 1316
bStepSuccess = False
LogWrite "Exit Code:" & ResultCode
LogWrite "The specified account already exists" 
Case 1605
If lcase(RunType) = "install" Then
bStepSuccess = True
LogWrite "Exit Code:" & ResultCode
LogWrite "Step failed because product is not installed; ignoring(This action is only valid for products that are currently installed)"
Else
bStepSuccess = False
bErrorFound = True
End If
Case 1641
bStepSuccess = True
LogWrite "Exit Code:" & ResultCode
LogWrite "Step was successful, and triggered a reboot"
Case 3010
bStepSuccess = True
LogWrite "Exit Code:" & ResultCode
LogWrite "Step was successful, and restart is required to complete the install"
ResultCode = 0
Case 19
bStepSuccess = True
ResultCode = 0
LogWrite "Exit Code:" & ResultCode
LogWrite "Step was successful"
Case Else
bStepSuccess = False
bErrorFound = True
End Select
Else 
Select Case ResultCode
Case 0
bStepSuccess = True
LogWrite "Step was successfully started"
Case Else
bStepSuccess = False
bErrorFound = True
End Select
End If
End Sub

Sub LogWrite(Message)
Dim oFSO, oLogFile
Const forAppending = 8

Set oFSO = CreateObject ("Scripting.FileSystemObject")
Set oLogFile = oFSO.OpenTextFile (LogFile,forAppending,True)
oLogFile.WriteLine Date & " " & Time & " " & message
oLogFile.Close

Set oFSO = Nothing
Set oLogFile = Nothing
End Sub


Function ExecutableFound(CommandLine)
Dim strExecutable, strExecutableAndExe
Dim iPos
Dim bContainsPath
Dim oFSO,oShell,windir
Set oShell = WScript.CreateObject("WScript.Shell")
windir = oShell.ExpandEnvironmentStrings("%windir%")

strExecutable = trim(CommandLine)

If Left (strExecutable,1) = """" Then
strExecutable = Mid(strExecutable,2,InStr(2,strExecutable,"""")-2)
Else
iPos = InStr(2,strExecutable," ")
if iPos > 0 Then strExecutable = left(strExecutable,iPos-1)
End If

If InStr(strExecutable,"\") > 0 Then 
bContainsPath = True
Else
bContainsPath = False
End If

If Instr(strExecutable,".") = 0 then strExecutableAndExe = strExecutable & ".exe"

Set oFSO = CreateObject ("Scripting.FileSystemObject")
ExecutableFound = False

If oFSO.FileExists (strExecutable) Then
ExecutableFound = True
ElseIf bContainsPath = False Then
If oFSO.FileExists (windir &"\system32\" & strExecutable) Then 
ExecutableFound = True
ElseIf oFSO.FileExists (windir & "\system32\" & strExecutableAndExe) Then 
ExecutableFound = True
End If
End If

If ExecutableFound = False Then LogWrite "ERROR: Executable not found: " & strExecutable

Set oFSO = Nothing
End Function


Sub CheckStepResult(bStepSuccess)

If bStepSuccess = False Then
SaveResultCodes = SaveResultCodes & ResultCode & " "
ProcessError ResultCode
End If
end sub


Sub ProcessError(ResultCode)

'LogWrite "ERROR: Step failed due to error " & ResultCode
If IgnoreError Then
ResultCode = 0 
LogWrite "Error ignored by request; continuing with next steps"
Else
'LogWrite "Processing cancelled due to error"
CloseOut 
End If
End Sub

Sub CheckErrAndLogResults(strMessage, ResultCodeIfFailed)

Dim bStepSuccess

If Err then
bStepSuccess = False
bErrorFound = True
LogWrite "ERROR: " & err.number & ": " & err.description
Err.Clear
LogWrite strMessage
ResultCode = ResultCodeIfFailed 
Else
bStepSuccess = True
LogWrite "Step was successful"
End If

CheckStepResult bStepSuccess
End Sub


Sub CheckForProcessingError(ResultCode)
If ResultCode <> 0 Then CloseOut
End Sub


Sub CheckForScriptError(Step)

If Err.Number <> 0 Then
LogWrite "ERROR: An error occurred during processing of step: " & Step & " - (" & CStr(Err.Number) & ") " & Err.Description
bErrorFound = True
ResultCode = 9991
Err.Clear
CloseOut
End If
End Sub


Sub CloseOut()

If bErrorFound and ResultCode <> 0 and IgnoreError = False Then
LogWrite Product & ": " & RunType & " failed (error " & ResultCode & ")"
ElseIf bErrorFound and IgnoreError Then
LogWrite Product & ": " & RunType & " was successful (ignored errors: " & trim(SaveResultCodes) & ")"
ElseIf ResultCode <> 0 Then
Select Case ResultCode
Case 1605
LogWrite Product & ": " & RunType & " was successful"
ResultCode = 0
Case 1641
LogWrite Product & ": " & RunType & " was successful"
ResultCode = 0
Case 3010
LogWrite Product & ": " & RunType & " was successful"
ResultCode = 0
Case Else
LogWrite Product & ": " & RunType & " failed (error " & ResultCode & ")" 
End Select
Else 
LogWrite Product & ": " & RunType & " was successful"
End If

LogWrite "Final result code: " & ResultCode
BuildLogWrite
WScript.Quit ResultCode
End Sub


Function WMIClassExists(strWMIPath, WMIClassName) 
Dim objWMIProvider, colClasses, objClass 

WMIClassExists = False 
Set objWMIProvider = GetObject("winmgmts:" & strWMIPath) 
If Err = False Then
Set colClasses = objWMIProvider.SubclassesOf() 
For Each objClass In colClasses 
if instr(objClass.Path_.Path,":" & WMIClassName) Then 
WMIClassExists = True 
End If 
Next
End If

Set objWMIProvider = Nothing 
Set colClasses = Nothing 
Set objClass = Nothing
CheckForScriptError "Checking if WMI class exists"
End Function

Function GetMachineName()
Dim oShell,strComputerName,strUserName
Set oShell = CreateObject("Wscript.Shell")
strComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
LogWrite "Computer Name: " & strComputerName


End Function

Function GetOSFamily()   
Dim strWMIPath, strCaption, strOSFamily 
Dim objWMIProvider, colOSInfo, oOSProperty 

LogWrite "Checking OS"
strWMIPath = "\\.\root\cimv2"

If WMIClassExists(strWMIPath, "Win32_OperatingSystem") Then
Set objWMIProvider = GetObject("winmgmts:{impersonationLevel=impersonate}!" & strWMIPath)
Set colOSInfo = objWMIProvider.ExecQuery("Select * from Win32_OperatingSystem") 
For Each oOSProperty in colOSInfo     
strCaption = oOSProperty.Caption   
If InStr(1,strCaption, "Windows 7", vbTextCompare) Then 
strOSFamily = "WIN7"
Elseif InStr(1,strCaption, "Microsoft Windows 10", vbTextCompare) Then 
strOSFamily = "WIN10"
Elseif InStr(1,strCaption, "Microsoft Windows Server 2016", vbTextCompare) Then 
strOSFamily = "Windows Server 2016"
Elseif InStr(1,strCaption, "Microsoft Windows Server 2008", vbTextCompare) Then 
strOSFamily = "Windows Server 2008"

End If
Next

LogWrite "OS family: " & strOSFamily
GetOSFamily = strOSFamily   
Set objWMIProvider = Nothing
Set colOSInfo = Nothing
Set oOSProperty = Nothing
End If

CheckForScriptError "Getting OS family"
CheckForProcessingError ResultCode
End Function

Function GetProcessorArchitecture()
Dim strProcessorFamily, strWMIPath
Dim objWMIProvider, colOSInfo, oProcessorProperty, strCaption

ResultCode = 9070
strWMIPath = "\\.\root\cimv2"
If WMIClassExists(strWMIPath, "Win32_Processor") Then
Set objWMIProvider = GetObject("winmgmts:{impersonationLevel=impersonate}!" & strWMIPath)
Set colOSInfo = objWMIProvider.ExecQuery("Select * from Win32_Processor") 
For Each oProcessorProperty in colOSInfo     
strCaption = oProcessorProperty.AddressWidth   
If InStr(1,strCaption, "64", vbTextCompare) Then 
strProcessorFamily = "x64"
ResultCode = 0
ElseIf InStr(1,strCaption, "32", vbTextCompare) Then 
strProcessorFamily = "x86" 
ResultCode = 0
Else
LogWrite "ERROR: Processor family not 32 or 64"
ResultCode = 9072
End If
Next

GetProcessorArchitecture = strProcessorFamily  
ResultCode = 0
Set objWMIProvider = Nothing  
Set colOSInfo = Nothing
Set oProcessorProperty = Nothing
End If

End Function

'-----------------Custom Functions starts here-----------

Sub Permission(strHomeFolder)
Dim WshShell,SysRoot
Dim intRunError,objFSO,win,exepath,file

Set WshShell = CreateObject("Wscript.Shell")
SysRoot = WshShell.ExpandEnvironmentStrings("%SystemDrive%")
win = WshShell.ExpandEnvironmentStrings("%windir%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'MsgBox strhomefolder 
exepath=win & "\system32\icacls.exe"
file= chr(34) & exepath & chr(34) & Chr(32) & strHomeFolder & " /Q /C /T /grant Users:(OI)(CI)F"
' file= chr(34) & exepath & chr(34) & Chr(32) & strHomeFolder & " /Q /C /T /grant Everyone:(OI)(CI)F" 
' N - no access
'                F - full access
'                M - modify access
'                RX - read and execute access
'                R - read-only access
'                W - write-only access
'                D - delete access
' For more insfo check ICACLS /?
intRunError = WshShell.Run(file,0,True)
'LogWrite "Exit code:" & intRunError 
'CheckForScriptError "Checking for Permission"
End Sub

Sub UninstallMSI(AppName)
On Error Resume Next
Const HKEY_CLASS_ROOT = &H80000000

Dim strComputer, strKeyPath, objShell, oReg, arrSubkeys, strkeyvalue
Dim position, searchleft, searchright, resultguid, strSubkey, strkeyvalue1 

strComputer = "."
strKeyPath = "Installer\Products"

Set objShell = CreateObject("WScript.Shell") 

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
strComputer & "\root\default:StdRegProv")

oReg.EnumKey HKEY_CLASS_ROOT, strKeyPath, arrSubkeys

If IsArray(arrSubkeys) Then 
For Each strSubkey In arrSubkeys
strkeyvalue1 = objShell.RegRead("HKCR\Installer\Products\" & strSubkey & "\ProductName")
If Err.Number<> 0 Then
Err.Clear 
Else 
position = InStr(1,strkeyvalue1, AppName, 1)
If position > 0 Then
strkeyvalue = objShell.RegRead("HKCR\Installer\Products\" & strSubkey & "\ProductIcon")
If Err.Number<> 0 Then
Err.Clear 
Else
searchleft="{"
searchright="}"
resultguid = Mid(strkeyvalue, InStr(strkeyvalue, searchleft), InStrRev(strkeyvalue, searchright)-InStr(strkeyvalue, searchleft)+1)
End If
LogWrite strkeyvalue1 & " Found. GUID is " & resultguid
RunCommandLine "msiexec /x " & resultguid & " REBOOT=REALLYSUPPRESS /qn /L*V """ & strkeyvalue1 & "_Uninst.log""" 
End If
End If
Next
End If

CheckForScriptError "Checking for installed products"
End Sub

Sub RemoveOrphanfolder(OFol)

Dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject")

if filesys.folderexists(OFol) Then
filesys.DeleteFolder OFol,True
LogWrite OFol & "-->Orphan Directory deleted successfully"
end if

Set filesys = Nothing

End Sub

Sub RemoveOrphanfile(OFile)

Dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject")

if filesys.fileexists(OFile) Then
'MsgBox OFile
filesys.DeleteFile OFile,True
LogWrite OFile & "-->Orphan File deleted successfully"
end if

Set filesys = Nothing

End Sub

Sub DeleteService(ServiceName)
Dim strComputer,objWMIService,colListOfServices,objService
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colListOfServices = objWMIService.ExecQuery _
("Select * from Win32_Service Where Name = '"&ServiceName&"'")

For Each objService in colListOfServices
objService.StopService()
objService.Delete()
LogWrite "Deleted Service--> " & ServiceName
Next
CheckForScriptError "Deleting Service..."
End Sub

Sub DeleteOrphanRegistry(HKey,strKeyPath)
Dim strComputer,objShell,oReg,arrSubKeys
Set objShell = CreateObject("WScript.Shell")
strComputer = "." 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

If oReg.EnumKey(HKey,strKeyPath,arrSubKeys)=0 Then
oReg.DeleteKey HKey,strKeyPath
LogWrite "Deleted the Orphan keys from registry-->" & strKeyPath
End If

Set objshell = Nothing
Set oReg = Nothing
Set arrSubkeys = Nothing
CheckForScriptError "Deleting Orphan Registry..."
End Sub

Sub DeleteOrphanRegistryOnlyIfEmpty(HKey,strKeyPath)
Dim strComputer,objShell,oReg,arrSubKeys
Set objShell = CreateObject("WScript.Shell")
strComputer = "." 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

oReg.EnumKey HKey, strKeyPath, arrSubKeys 
If IsNull(arrSubKeys) Then
oReg.DeleteKey HKey, strKeyPath
LogWrite "Deleted the Orphan keys from registry-->" & strKeyPath
End If
CheckForScriptError "Deleting Registry only Empty..."
End Sub

Sub RemoveFolderOnlyIfEmpty(sFolderName)
    Dim fso,objFolder
    Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(sFolderName) Then
Set objFolder = fso.GetFolder(sFolderName)
If objFolder.Size = 0 Then
fso.DeleteFolder(sFolderName)
LogWrite "Deleted folder-->" & sFolderName
End If
End If            
    Set fso = Nothing
CheckForScriptError "Deleting Folder only Empty..."
End Sub

Sub DeleteSubkeys(HKey, strKeyPath)
On Error Resume Next 
Dim strComputer,objRegistry,arrSubkeys,strSubkey
strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & _
    strComputer & "\root\default:StdRegProv") 
    
    objRegistry.EnumKey HKey, strKeyPath, arrSubkeys 

    If IsArray(arrSubkeys) Then 
        For Each strSubkey In arrSubkeys 
            DeleteSubkeys HKey, strKeyPath & "\" & strSubkey 
        Next 
    End If 

    objRegistry.DeleteKey HKey, strKeyPath 
    LogWrite "Deleted Subkeys-->" &strKeyPath
    CheckForScriptError "Deleting Subkeys..."
End Sub

Sub RegAddDWORD(HKey,EyeOneMatchkey,strValueName,dwValue)

Dim objRegistry,strComputer,arrSubKeys,wshShell
strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
set wshShell= Wscript.CreateObject("WScript.Shell")

If objRegistry.EnumKey(HKey,EyeOneMatchkey ,arrSubKeys)=0 Then
objRegistry.SetDWORDValue HKey,EyeOneMatchkey,strValueName,dwValue
LogWrite "Added registry --> " & strValueName & "=" & dwValue 
Else
objRegistry.CreateKey HKey,EyeOneMatchkey
objRegistry.SetDWORDValue HKey,EyeOneMatchkey,strValueName,dwValue
LogWrite "Added registry --> " & strValueName & "=" & dwValue
End if

Set objRegistry = Nothing
Set strComputer = Nothing

End Sub

Sub RegAddSTRING(HKey,EyeOneMatchkey,strValueName,dwValue)

Dim objRegistry,strComputer,arrSubKeys,wshShell
strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
set wshShell= Wscript.CreateObject("WScript.Shell")

If objRegistry.EnumKey(HKey,EyeOneMatchkey ,arrSubKeys)=0 Then
objRegistry.SetStringValue HKey,EyeOneMatchkey,strValueName,dwValue
LogWrite "Added registry --> " & strValueName & "=" & dwValue
Else
objRegistry.CreateKey HKey,EyeOneMatchkey
objRegistry.SetStringValue HKey,EyeOneMatchkey,strValueName,dwValue
LogWrite "Added registry --> " & strValueName & "=" & dwValue 
End if

Set objRegistry = Nothing
Set strComputer = Nothing

End Sub

Sub KillProcess(ProcessName)
    Dim objWMIService,colProcessList,WshShell,strComputer,objFSO,objProcess
    strComputer = "."  
    Set WshShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '"&ProcessName&"'")
    If colProcessList.Count > 0 Then
For Each objProcess in colProcessList
objProcess.Terminate()
Next 
LogWrite "Killed process-->" & ProcessName
    End If

    Set objWMIService = Nothing
    Set colProcessList = Nothing
    Set WshShell = Nothing
    Set strComputer = Nothing

    CheckForScriptError " Checking whether application related process is running "

End Sub


Sub CopyFile(Source,Destination)
Dim filesys, oShell
set filesys = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

if filesys.FileExists(Source) Then
  filesys.CopyFile Source, Destination, True 
End If
  
LogWrite Source & "--File Copied to--" & Destination
Set oShell = Nothing
set filesys = Nothing

CheckForScriptError "Checking for FileCopy" 
End Sub

Sub CopyFolder(Source,Destination)
Dim filesys, oShell
set filesys = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

if filesys.FolderExists(Source) Then
  filesys.CopyFolder Source, Destination , True
End If
  
LogWrite Source & "--Folder Copied to--" & Destination
Set oShell = Nothing
set filesys = Nothing

CheckForScriptError "Checking for FolderCopy" 
End Sub


Sub BuildLogWrite()
Dim oFSO, bLogFile, pLogFile
Const forAppending = 8

Set oFSO = CreateObject ("Scripting.FileSystemObject")
Set bLogFile = oFSO.OpenTextFile (BuildLogFile,forAppending,True)
bLogFile.WriteLine "Start: " & Start_date & " End: " & Date & " " & Time & " " & Publisher & "\" & AppName & "\" & AppVersion & " " & RunType & "ed by " & StrUsername & " RC " & ResultCode & " Reg Access " & RegaccessValue
bLogFile.Close

Set pLogFile = oFSO.OpenTextFile (BuildLogFileProg,forAppending,True)
pLogFile.WriteLine "Start: " & Start_date & " End: " & Date & " " & Time & " " & Publisher & "\" & AppName & "\" & AppVersion & " " & RunType & "ed by " & StrUsername & " RC " & ResultCode & " Reg Access " & RegaccessValue
pLogFile.Close

Set oFSO = Nothing
Set bLogFile = Nothing

End Sub

' -------- CHANGE THIS SECTION (2) --------------------
' Any custom functions needed, will be placed here


Function Regaccess
Regaccess=0
End Function

Function RegaccessNope
Dim oFSO,a,content,Reg,file
Set oFSO = CreateObject("Scripting.FileSystemObject")
If oFSO.FileExists (ProgramFilesX86 & "\sms\loggedon.exe") Then
    Set a = oFSO.CreateTextFile(ProgramFilesX86 & "\sms\test.txt", True)
Dim Text
Text = Chr(34) & "C:\Program Files (x86)\sms\loggedon.exe" & Chr(34) & " -l -x >" & Chr(34) & "C:\Program Files (x86)\sms\test.txt" & Chr(34) 
    a.WriteLine(text)
    a.Close
 CopyFile ProgramFilesX86 & "\sms\Test.txt",ProgramFilesX86 & "\sms\Test.bat"
RunCommandLine Chr(34) & ProgramFilesX86 & "\sms\Test.bat" & Chr(34)
Set file = oFSO.OpenTextFile(ProgramFilesX86 & "\sms\Test.txt", 1)
content = file.ReadAll
If InStr(1,content, "Massmutual", vbTextCompare) Then
Reg = 1
Else
Reg = 0
End If
Regaccess = Reg
Else
  CopyFile "\\iscapmm1\DDS\SDS Tools\AdminStudio\WiseEditor\DDS Template\Loggedon.exe",Windir & "\Temp\loggedon.exe"
        Set a = oFSO.CreateTextFile(Windir & "\temp\test.txt", True)
    a.WriteLine("C:\WINDOWS\temp\loggedon.exe -l -x >C:\WINDOWS\temp\test.txt")
    a.Close
 CopyFile Windir & "\temp\Test.txt",Windir & "\temp\Test.bat"
RunCommandLine Chr(34) & Windir & "\temp\Test.bat" & Chr(34)
Set file = oFSO.OpenTextFile(Windir & "\temp\Test.txt", 1)
content = file.ReadAll
If InStr(1,content, "Massmutual", vbTextCompare) Then
Reg = 1
Else
Reg = 0
End If
Regaccess = Reg
End If
End Function

sub popupforuserloggedin(ExeName)

RegaccessValue = Regaccess
If RegaccessValue = 1 then
RunCommandLine Chr(34) & oempath & ExeName & Chr(34)
Else 
RunCommandLine Chr(34) & "shutdown.exe" & Chr(34) & " /R /T 60" 
End If

End Sub


Sub MM_Branding(bitness,ProdLoc)

If bitness="32" Then

RegAddSTRING HKLM,"SOFTWARE\WOW6432Node\Massmutual\" & Product,"ApplicationName",Publisher & "_" & Appname
RegAddSTRING HKLM,"SOFTWARE\WOW6432Node\Massmutual\" & Product,"ApplicationVersion",AppVersion
RegAddSTRING HKLM,"SOFTWARE\WOW6432Node\Massmutual\" & Product,"SourceLocation",ProdLoc

Else

RegAddSTRING HKLM,"SOFTWARE\Massmutual\" & Product,"ApplicationName",Publisher & "_" & Appname
RegAddSTRING HKLM,"SOFTWARE\Massmutual\" & Product,"ApplicationVersion",AppVersion
RegAddSTRING HKLM,"SOFTWARE\Massmutual\" & Product,"SourceLocation",ProdLoc

End If

End Sub

Sub Delete_MM_Branding(bitness)

If bitness="32" Then

DeleteSubkeys HKLM,"SOFTWARE\WOW6432Node\Massmutual\" & Product

Else

DeleteSubkeys HKLM,"SOFTWARE\Massmutual\" & Product

End If

End Sub
