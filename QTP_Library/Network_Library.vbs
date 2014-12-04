'Authored By: Kunal Dey
'Date: 24/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 IsIpReachable
'02 GetMacAdd
'03 GetIpCompname

'************************************************************
'************************************************************

'Checks if a particular IP address is reachable from the host machine
'==========================================================================
Function IsIpReachable(strComputer)
'     On Error Resume Next
    Dim wmiQuery, objWMIService, objPing, objStatus
    
    wmiQuery = "Select * From Win32_PingStatus Where " & _
    "Address = '" & strComputer & "'"
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set objPing = objWMIService.ExecQuery(wmiQuery)
    
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then
            Reachable = False 'if computer is unreacable, return false
        Else
            Reachable = True 'if computer is reachable, return true
        End If
    Next
    
End Function

'Retrieve the Mac Address of a machine from the host machine
'==========================================================================
Function GetMacAdd(strComputer)
'     On Error Resume Next
    Dim wmiQuery, objWMIService, objPing, objStatus
    
    wmiQuery = "Select * From Win32_PingStatus Where " & _
    "Address = '" & strComputer & "'"
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set objPing = objWMIService.ExecQuery(wmiQuery)
    
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then
            ResolveIP = "Computer is Unreachable!"
        Else
            ResolveIP = objStatus.ProtocolAddress
        End If
    Next
    
End Function

'Retrieve IPaddress and Computer Name of LocalHost
'==========================================================================
Function GetIpCompname(strIP,compName)

	Set objWMI = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

		For Each Nic in objWMI
			if Nic.IPEnabled then
				strIP = Nic.IPAddress(i)
				Set WshNetwork = CreateObject("WScript.Network")
				compName= WshNetwork.Computername
			end if
		next

End Function
