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

