'by Vakulenko Nikolai 08.2015
option Explicit
'on error resume next
const HKEY_LOCAL_MACHINE = &H80000002
dim objWbemLocator, objWMIService, objWMIServiceReg, objReg, objFSO, objProc, colServices, objService
dim strComputer, strKeyPath, RemoteLocalServerPath, RemoteCopyPath, LocalProgramBase
dim BitWidth, intProcessId
dim dqt

if WScript.Arguments.Count < 1 then
	WScript.Echo "Machine name argument is absent."
	Lquit(-10)
end if
strComputer = WScript.Arguments(0)
'connecting to remote side
Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objwbemLocator.ConnectServer(strComputer, "root\CIMV2")
if Err.Number <> 0 then
	WScript.Echo "Authentication failed: " & " " & Err.Number & ": " & Err.Description
	Lquit(Err.Number)
end if
'on XP machines StdRegProv available only on default object
Set objWMIServiceReg = objWbemLocator.ConnectServer(strComputer, "\root\default")


'pathname of tvnserver.exe, from local host point of view
dqt = chr(34)
RemoteLocalServerPath = dqt & "C:\Program Files\TightVNC\tvnserver.exe" & dqt
RemoteCopyPath = "\\" & strComputer & "\c$\Program Files\TightVNC\"
LocalProgramBase = "C:\Program Files (x86)\Total Network Inventory 3\VNC\"

'check server is running
Set colServices = objWMIService.ExecQuery ("SELECT * FROM Win32_Service WHERE Name = 'tvnserver'")
For Each objService in colServices
	If objService.Started Then
			'connect without install
			WScript.Echo "connect without install to: " & strComputer
			LClientConnect objWMIService
			WScript.Quit
	End If
Next

'installing remote service...
'registry manipulating
WScript.Echo "configuring registry settings..."
Set objReg = objWMIServiceReg.Get("StdRegProv") 
strKeyPath = "SOFTWARE\TightVNC\Server"
objReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
objReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"ExtraPorts",""
objReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"RunControlInterface",0
objReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"AcceptRfbConnections",0
objReg.SetDWORDValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","SoftwareSASGeneration",1

'copy files
WScript.Echo "copying files..."
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists(RemoteCopyPath) Then
	BitWidth = objWMIService.Get("Win32_Processor='cpu0'").AddressWidth
	objFSO.CreateFolder RemoteCopyPath
	objFSO.CopyFile LocalProgramBase & "install_TightVNCx" & BitWidth & "\*",  RemoteCopyPath,  False
end if

'startup service
WScript.Echo "starting service..."
Set objProc = objWMIService.Get("Win32_Process") 
objProc.Create RemoteLocalServerPath & " -reinstall -silent", null, null, intProcessId
LWaitPID objWMIService, intProcessId
objProc.Create RemoteLocalServerPath & " -start -silent ", null, null, intProcessId
LWaitPID objWMIService, intProcessId
'subroutine for wait while tvnserver ends it's job
Sub LWaitPID(p_objWMIService, p_intProcessId)
	Dim LcolProcessList
	do
		Set LcolProcessList = p_objWMIService.ExecQuery _
			("Select * from Win32_Process Where ProcessID = " & p_intProcessId)
		wscript.sleep 10
		if LcolProcessList.count = 0 then
			exit do
		end if
	loop
End Sub

'connect after install
LClientConnect objWMIService
'end.

Sub Lquit(p_error)
	WScript.Echo "wait keypress..."
    Do While Not WScript.StdIn.AtEndOfLine
        Input = WScript.StdIn.Read(1)
    Loop
	WScript.Quit(p_error)
end sub

Sub LClientConnect(p_objWMIService)
	Dim LobjWMIService, LcolProcess, LobjProcess, LwshShell, LcolItems, LobjItem
	Dim LintProcessId
	Dim LstrLocalHost, LstrIP
	'if viewer hasn't been started on localhost, stating it
	Set LobjWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\.\root\cimv2") 
	Set LcolProcess = LobjWMIService.ExecQuery _
	("Select * from Win32_Process WHERE Name = 'tvnviewer.exe' ")
	If LcolProcess.Count = 0 Then
		Set LobjProcess = LobjWMIService.Get("Win32_Process") 
		if LobjProcess.Create( dqt & LocalProgramBase & "tvnviewer.exe" & dqt & " -listen", null, null, LintProcessId) then
			Wscript.Echo "Error occured, while start tvnviewer.exe on local machine."
			Lquit
		end if
	End If
	'gettin localmachine name
	'Set LwshShell = WScript.CreateObject( "WScript.Shell" )
	'LstrLocalHost = LwshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
	'gettin localmachine ipaddress
	Set LcolItems = LobjWMIService.ExecQuery( _
	"SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True",,48)
	'finding network adapter, which have non empty domain field
	For Each LobjItem in LcolItems
		if not IsNull(LobjItem.DefaultIPGateway) then
			For Each LstrIP In LobjItem.IPAddress 
				'if it is not ipv6, it have length 7-15 characters
				if Len(LstrIP) < 16 then
					LstrLocalHost = LstrIP
				end if
			next     
		end if
	Next
	'remote server will try to connect to localhost
	Set LobjProcess = p_objWMIService.Get("Win32_Process") 
	LobjProcess.Create RemoteLocalServerPath & " -controlservice -connect " & LstrLocalHost, null, null, LintProcessId
End Sub
