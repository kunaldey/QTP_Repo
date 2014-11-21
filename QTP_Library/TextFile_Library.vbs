'Authored By: Kunal Dey
'Date: 19/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 CreateTextFile



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

