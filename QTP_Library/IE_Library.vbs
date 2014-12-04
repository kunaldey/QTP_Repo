'Authored By: Kunal Dey
'Date: 24/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 OpenWebPage
'02 DeleteAFile

'************************************************************
'************************************************************

'Open any webpage or HTML page in IE
'=================================================================
Function OpenWebPage(webPath,pageVisible)
	
	Set objIE = CreateObject("InternetExplorer.Application")
		objIE.Visible = pageVisible
	    objIE.Navigate webPath
	    wait objIE.ReadyState
	    
'	    MsgBox objIE.ReadyState
	    getUrl = objIE.LocationURL
	    Set objDoc = objIE.Document
'	    Set objUserID = objDoc.getElementById("userId")
'	    MsgBox getUrl
'		On Error Resume Next
'		If objUserID Is Nothing Then
'			Print WebPath & " Could not be loaded"
'			Else
'			Print webPath & " Loaded successfully"
'		End If
'	    If objIE.Document.readyState="complete" Or objIE.Document.readyState="loaded" or objIE.Document.readyState="interactive" Then
''	    	If InStr (getUrl,webPath) Then
'	    		Print webPath & " Loaded successfully"
'	    		Else
'	    		Print WebPath & " Could not be loaded"
'	    	End If
		
'		objIE.Quit
	Set objIE = Nothing
	
End Function

Sub CheckIEVisible(objIE)
  Do
  WScript.Sleep 500
  Loop While objIE.ReadyState < 4 And objIE.Busy 
End Sub
