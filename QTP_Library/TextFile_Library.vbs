'Authored By: Kunal Dey
'Date: 19/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 CreateTextFile
'02 ReadTextFile
'03 FindString



'************************************************************
'************************************************************

'Create an XML File
'=============================================================================
Function CreateTextFile(filePath,writeText)

   Dim fso, txtFile
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set txtFile = fso.CreateTextFile(filePath, True)
   txtFile.WriteLine(writeText)
   txtFile.Close
   
End Function

'Read the entire content of a Text file
'=============================================================================
Function ReadTextFile(filePath)

   Dim fso, txtFile
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set txtFile = fso.OpenTextFile(filePath)
	Do Until txtFile.AtEndOfStream
    getLine = txtFile.ReadLine
    Print getLine
    Loop
    
End Function

'Find a string in a text file and return the line number
'Note: Returns an array, while calling this function follow the code snippet below...
'	Dim foundLine
'		foundLine = FindChildNode ("C:\Sample Files\Test.txt","OpenTextFile",totalFound)
'			For i = 0 To UBound(foundLine)
'			 print foundLine(i)
'			Next
'=============================================================================
Function FindString(filePath,targetString,totalFound)

   Dim fso, txtFile
   
	   Set fso = CreateObject("Scripting.FileSystemObject")
	   Set txtFile = fso.OpenTextFile(filePath)
	   lineNum = 1

		Do Until txtFile.AtEndOfStream
		lineNum = lineNum + nextLine
	    getLine = txtFile.ReadLine
	    	If InStr(getLine,targetString) Then
					Print "Line : " & lineNum & " || Match Found : " & targetString
	    			Print "Found Line Is : " & getLine
	    			totalFound = totalFound + 1   			

					If lineNUmber = "" Then
						lineNumber = lineNum
						Else
						lineNumber = lineNum & "," & lineNumber
					End If
	    	End If
	    nextLine = 1
	    Loop
	    foundLine = Split (lineNumber,",")
	FindString = foundLine
	
End Function

