'Authored By: Kunal Dey
'Date: 26/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 LaunchExistingTest
'02 CustomizeSaveResult
'03 CreateNewTest
'04 ConnectQC
'05 AssoORToTest
'06 AssoLibToTest
'07 DefineEnvValToTest
'08 GetActionNames
'09 MinimizeQTP
'10 MaximizeQTP
'************************************************************
'************************************************************

'Start  QTP, open an existing test and Run the Test:
'===================================================================================
Function LaunchExistingTest(testPath)
	Dim objQtApp
	Dim qtTest
	
		
		Set objQtApp = CreateObject("QuickTest.Application") 
		
			
			If  objQtApp.launched <> True then 
				objQtApp.Launch 
			End If 

				objQtApp.Visible = True
							
				objQtApp.Options.Run.ImageCaptureForTestResults = "OnError"
				objQtApp.Options.Run.RunMode = "Fast"
				objQtApp.Options.Run.ViewResults = False
				
				'Open the test in read-only mode
				objQtApp.Open testPath, True 

				Set qtTest = objQtApp.Test		
'				qtTest.Settings.Run.OnError = "NextStep" 
				qtTest.Run
'				MsgBox qtTest.LastRunResults.Status
				qtTest.Close 

		objQtApp.quit

	Set qtTest = Nothing
	Set objQtApp = Nothing
	
End Function

'Start  QTP, open an existing test and Run the Test  And Store 
'Run Results in Specified Folder
'===================================================================================
Function CustomizeSaveResult(testPath,resultPath)
	
	Dim objQtApp
	Dim qtTest
	Dim qtResultsOpt
	
		Set objQtApp = CreateObject("QuickTest.Application") 
		
			
			If  objQtApp.launched <> True then 
				objQtApp.Launch 
			End If 
			
				
				objQtApp.Visible = True
				
				'Set QuickTest run options
				objQtApp.Options.Run.ImageCaptureForTestResults = "OnError"
				objQtApp.Options.Run.RunMode = "Fast"
				objQtApp.Options.Run.ViewResults = False
				
				'Open the test in read-only mode
				objQtApp.Open testPath, True 
				'set run settings for the test
				Set qtTest = objQtApp.Test
			'	qtTest.Settings.Run.OnError = "NextStep" 
				'Create the Run Results Options object
				Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions")
				'Set the results location
				qtResultsOpt.ResultsLocation = resultPath
				' Run the test
				qtTest.Run qtResultsOpt 
			'	MsgBox qtTest.LastRunResults.Status
	
		qtTest.Close 
		objQtApp.quit

	Set qtResultsOpt = nothing
	Set qtTest = Nothing
	Set objQtApp = Nothing

End Function

'Start  QTP and open New test
'==================================================================================
Function CreateNewTest()
	
	Dim objQtApp
	Dim qtTest
	
		
		Set objQtApp = CreateObject("QuickTest.Application") 
		
		
		If  objQtApp.launched <> True then 
		objQtApp.Launch 
		End If 
		objQtApp.Visible = True
		
		' Open a new test
		objQtApp.New
	
	Set objQtApp = Nothing

End Function

' Open QTP and Connect to Quality Center and run QC script
'Note: 'qcURL = QC Server path
	   'domainName = Domain name that contains QC project
	   'projName =Project Name in QC you want to connect to
	   'usrName = Username to connect to Project
	   'usrPwd = Password to connect to project
	   'False or True = Whether ‘password is entered in encrypted or normal.
	   'scriptPath = Path of the script and script name in QC 
	   '(Example: "Subject\QCScriptPath\ScriptName")
'Default Inputs:"http://200.168.1.1:8080/qcbin","Default","proj1","qtpworld","qtp",false
'===================================================================================
Function ConnectQC(qcURL,domainName,projName,usrName,usrPwd,scriptPath)
	
	Dim objQtApp
	Set objQtApp = CreateObject("QuickTest.Application") 
	
		
		If  objQtApp.launched <> True then 
			objQtApp.Launch 
		End If 

			objQtApp.Visible = True
				If Not objQtApp.TDConnection.IsConnected Then				
					objQtApp.TDConnection.Connect qcURL,domainName,projName,usrName,usrPwd,False			
				End If			
			'Make Sure about your script path  and script name in QC
			objQtApp.Open "[QualityCenter]" & scriptPath, False
			objQtApp.Test.Run
			objQtApp.TDConnection.Disconnect

		objQtApp.quit
	Set objQtApp = Nothing
	
End Function

'Start  QTP, open an existing test, associate Object Repositories and save the test
'===================================================================================
Function AssoORToTest(testPath,actionName,repoPath)
	
	Dim objQtApp
	Dim qtTest
	Dim qtRepositories
	
	Set objQtApp = CreateObject("QuickTest.Application") 
	
			If  objQtApp.launched <> True then 
				objQtApp.Launch 
			End If 
				
				objQtApp.Visible = True
				objQtApp.Open testPath, False

				Set qtRepositories = objQtApp.Test.Actions(actionName).ObjectRepositories 
					
					If qtRepositories.Find(repoPath) = -1 Then 
					    qtRepositories.Add repoPath, 1
					End If
		
			objQtApp.Test.Save
		
		objQtApp.quit

	Set qtLibraries = Nothing
	Set qtTest = Nothing
	Set objQtApp = Nothing

End Function

' Start  QTP, open an existing test, associate libraries and save the test: 
'===================================================================================
Function AssoLibToTest(testPath,libPath)

	Dim objQtApp
	Dim qtTest
	Dim qtLibraries
		
	Set objQtApp = CreateObject("QuickTest.Application") 
		
		If  objQtApp.launched <> True then 
			objQtApp.Launch 
		End If 
			
			objQtApp.Visible = True
			objQtApp.Open testPath, False
			Set qtLibraries = objQtApp.Test.Settings.Resources.Libraries 
			'If the library file "libraary.vbs" is not assiciates with the Test then associate it
				If qtLibraries.Find(libPath) = -1 Then 
				    qtLibraries.Add libPath, 1  
				End If
		
			objQtApp.Test.Save
	
		objQtApp.quit

	Set qtLibraries = Nothing
	Set qtTest = Nothing 
	Set objQtApp = Nothing
	
End Function

'Start QTP, Open an Existing Test and Define Environment Variables: 
'=================================================================================== 
Function DefineEnvValToTest(testPath,varName(),varVal())

	Dim objQtApp
	Set objQtApp = CreateObject("QuickTest.Application") 
	
		If  objQtApp.launched <> True then 
		 	objQtApp.Launch 
		End If 
		
			objQtApp.Visible = True
						
				objQtApp.Open testPath, False
				' Set some environment variables
				For iCounter = 0 To UBound(varName)
					objQtApp.Test.Environment.Value(varName(iCounter)) = varVal(iCounter)
				Next
				
			objQtApp.Test.Save
		
		objQtApp.quit
	
	Set objQtApp = Nothing	

End Function

'Start QTP, Open an Existing Test and Get All Available Action Names From the Test 
'===================================================================================
Function GetActionNames(testPath)

	Dim objQtApp
	Dim actionNames()
	Set objQtApp = CreateObject("QuickTest.Application") 
	
		If  objQtApp.launched <> True then 
		objQtApp.Launch 
		End If 
		
			objQtApp.Visible = True
			objQtApp.Open testPath, False, False
			
			oActCount = objQtApp.Test.Actions.Count
			ReDim actionNames(oActCount)
				For iCounter=1 to oActCount
					actionNames(iCounter) = objQtApp.Test.Actions(iCounter).Name
				Next
	
		objQtApp.Quit
	Set objQtApp = Nothing
	GetActionNames = actionNames

End Function

'Open and minimize QTP Window
'===================================================================================
Function MinimizeQTP()
	
	Dim objQtApp
	Set objQtApp = CreateObject("QuickTest.Application") 
	
		If  objQtApp.launched <> True then 
			objQtApp.Launch 
		End If 
	
		objQtApp.Visible = True
		objQtApp.WindowState = "Minimized" ' Minimize the QuickTest window
	Set objQtApp = Nothing

End Function

'Open and maximize QTP Window
'===================================================================================
Function MaximizeQTP()
	
	Dim objQtApp
	Set objQtApp = CreateObject("QuickTest.Application") 
	
		If  objQtApp.launched <> True then 
			objQtApp.Launch 
		End If 
	
		objQtApp.Visible = True
		objQtApp.WindowState = "Maximized" ' Maximize the QuickTest window
	Set objQtApp = Nothing
	
	End Function

