' Script to allow ZCM agent to be install in MDT
' Written by Vaughn Miller  September 2012
' ###########################################################
'
' Usage : Place script in the same folder as the PreAgentPkg_AgentCompleteDotNet.exe
'         file and import as an application into MDT.  Make the command 
'         cscript install.vbs
'
' ###########################################################

' Launch the installer with the silent, no reboot switches.
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "PreAgentPkg_AgentCompleteDotNet.exe", "-q " & "-x"

' Set the process name and computer name
strComputer = "." ' local computer
strProcess = "Setup.exe"

' Loop until we see that process has started 
Do Until isProcessRunning(strComputer,strProcess)
  WScript.Sleep(5000)
Loop

' Loop until we see that the process has stopped
Do While isProcessRunning(strComputer,strProcess)
	WScript.Sleep(5000)
Loop

WScript.Quit 0

' Function to check if a process is running
FUNCTION isProcessRunning(BYVAL strComputer,BYVAL strProcessName)
	DIM objWMIService, strWMIQuery
	strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
	SET objWMIService = GETOBJECT("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
			& strComputer & "\root\cimv2") 

	IF objWMIService.ExecQuery(strWMIQuery).Count > 0 THEN
		isProcessRunning = TRUE
	ELSE
		isProcessRunning = FALSE
	END IF
END FUNCTION
