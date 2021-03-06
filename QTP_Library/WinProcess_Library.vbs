'Authored By: Kunal Dey
'Date: 02/12/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 IsProcessRunning
'02 TerminateProcess
'03 StartProcess
'04 CountProcess
'05 CountAllProcess
'06 GetAllProcessName
'07 GetAllProcessID
'08 GetThreadList
'09 GetThreadStatus
'10 GetProcessDetail

'************************************************************
'************************************************************

'Checks if a particular window process is running
'=================================================================================
Function IsProcessRunning( strComputer, strProcess )

    Dim Process, strObject
    IsProcessRunning = False
    strObject   = "winmgmts://" & strComputer
	    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
		    If UCase( Process.name ) = UCase( strProcess ) Then
		        IsProcessRunning = True
		        Print strProcess & " process is running on " & strComputer
		        Exit Function
			Else
				procState = strProcess & " process is not running on " & strComputer	
		    End If
	    Next
	  	Print procState 
    
End Function

'Terminate a particular window process is running
'=================================================================================
Function TerminateProcess( strComputer, strProcess )

    Dim Process, strObject
		strObject   = "winmgmts://" & strComputer
		    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
		    	If UCase( Process.name ) = UCase( strProcess ) Then
		    		Process.Terminate
		    		Print strProcess & " process tree is terminated successfully on " & strComputer
		        Exit Function
		        Else
		        	procState = strProcess & " process is not running, so can not terminate on " & strComputer
		        End If
		    Next
			Print procState		    
		    
End Function

'Start any windows process
'=================================================================================
Function StartProcess(strComputer,strProcess)

	Dim objWMIService, objProcess, objCalc
	
	set objWMIService = getobject("winmgmts://" & strComputer & "/root/cimv2") 

		Set objProcess = objWMIService.Get("Win32_Process")
		Set objProgram = objProcess.Methods_("Create").InParameters.SpawnInstance_
		objProgram.CommandLine = strProcess
	
		Set strShell = objWMIService.ExecMethod("Win32_Process", "Create", objProgram) 
		
		Print strProcess & " process is up and running on " & strComputer
		
	Set strShell = Nothing
	Set objWMIService = Nothing
	
End Function

'Count the windows process of same type
'=================================================================================
Function CountProcess( strComputer, strProcess )

    Dim Process, strObject
		strObject   = "winmgmts://" & strComputer
		    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
		    	If UCase( Process.name ) = UCase( strProcess ) Then
					countProc = countProc + 1
		        End If
		    Next
	CountProcess = countProc		    
		    
End Function

'Count all the windows process
'=================================================================================
Function CountAllProcess(strComputer)

    Dim Process, strObject
		strObject   = "winmgmts://" & strComputer
		    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
				countProc = countProc + 1
		    Next
	CountAllProcess = countProc		    
		    
End Function

'List all the process name
'=================================================================================
Function GetAllProcessName(strComputer)

    Dim objWMIService, colServices
    Dim procName()
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colServices = objWMIService.ExecQuery _
	    ("Select * from Win32_Process")
		For Each process In  colServices
			countProc = countProc + 1
			ReDim Preserve procName(countProc) 'Preserving the previous content of Array
			procName(countProc-1) = process.Name
		Next
	
		GetAllProcessName = procName
	Set	colServices = Nothing
	Set objWMIService = Nothing
	
End Function

'List all the process ID
'=================================================================================
Function GetAllProcessID(strComputer)

    Dim objWMIService, colServices
    Dim procID()
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colServices = objWMIService.ExecQuery _
	    ("Select * from Win32_Process")
		For Each process In  colServices
			countProc = countProc + 1
			ReDim Preserve procID(countProc) 'Preserving the previous content of Array
			procID(countProc-1) = process.ProcessId
		Next
	
		GetAllProcessID = procID
	Set	colServices = Nothing
	Set objWMIService = Nothing
	
End Function

'Retuns thread list for a single process (single instance only)
'=================================================================================
Function GetThreadList(strComputer,strProcess)

	Dim arrThread()
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery _
	    ("Select * from Win32_Process Where Name = '" & strProcess & "'")
		If colProcesses.count > 1 Then
			Print "Eroor retrieving threads as " & strProcess & " process has multiple instances"
		Else
			Set colThreads = objWMIService.ExecQuery _
			    ("Select * from Win32_Thread")
				For each objThread in colThreads
					threadCnt = threadCnt + 1 
					ReDim Preserve arrThread(threadCnt)
					arrThread(threadCnt -1) = objThread.Handle

				Next
		End If

		GetThreadList = arrThread
	Set	colServices = Nothing
	Set objWMIService = Nothing
	
End Function

'Retuns thread status list for a single process (single instance only)
'=================================================================================
Function GetThreadStatus(strComputer,strProcess)

	Dim arrThread()
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery _
	    ("Select * from Win32_Process Where Name = '" & strProcess & "'")
		If colProcesses.count > 1 Then
			Print "Eroor retrieving threads as " & strProcess & " process has multiple instances"
		Else
			Set colThreads = objWMIService.ExecQuery _
			    ("Select * from Win32_Thread")
				For each objThread in colThreads
					threadCnt = threadCnt + 1 
					ReDim Preserve arrThread(threadCnt)
					arrThread(threadCnt -1) = objThread.ThreadState

				Next
		End If

		GetThreadStatus = arrThread
	Set	colServices = Nothing
	Set objWMIService = Nothing
	
End Function

'Retrives details of a particular process (single & multiple instances)
'=================================================================================
Function GetProcessDetail(strComputer,strProcess)

    Dim objWMIService, colServices
    Dim arrProcess()
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colServices = objWMIService.ExecQuery _
	    ("Select * from Win32_Process Where Name = '" & strProcess & "'")
			For Each process In  colServices
				instanceID = 6 + itemCount
				ReDim Preserve arrProcess(instanceID)
					arrProcess(0 + itemCount) = process.CreationDate
					arrProcess(1 + itemCount) = process.ProcessId
					arrProcess(2 + itemCount) = process.ExecutablePath
					arrProcess(3 + itemCount) = process.PageFileUsage
					arrProcess(4 + itemCount) = process.SessionId
					arrProcess(5 + itemCount) = process.WorkingSetSize
					Print instanceID
					itemCount = UBound(arrProcess)
			Next
		GetProcessDetail = arrProcess

	Set	colServices = Nothing
	Set objWMIService = Nothing
	
End Function
