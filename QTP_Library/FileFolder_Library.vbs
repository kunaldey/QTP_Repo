'Authored By: Kunal Dey
'Date: 21/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 CheckFileExistence
'02 DeleteAFile
'03 GetFileNames
'04 DeleteAllFiles
'05 CountAllFiles
'06 CountAllFiles
'07 CheckFolderExistence
'08 DeleteAFolder
'09 CountChildFolder
'10 DeleteChildFolder
'************************************************************
'************************************************************

'Check existence of any file type
'==========================================================================
Function CheckFileExistence(filePath,fileExist)

	Dim fileObj
		Set fileObj = CreateObject("Scripting.FileSystemObject")
		If fileObj.FileExists(filePath) Then
			fileExist = 1
			Print "File Found : " & filePath
			Else
			fileExist = 0
			Print "File Not Found : " & filePath
		End If
		Set CheckFileExistence = fileObj 'Returning fileObj object
		
End Function

'Deletion of a particular file type
'==========================================================================
Function DeleteAFile(filePath)

   Dim fileObj, txtFile
   Set fileObj = CheckFileExistence (filePath,fileExist)   
   	If fileExist = 1 Then 
   		fileObj.DeleteFile filePath
   		Print "Successful Deletion of : " & filePath
   		Else
   		Print "Deletion not Successful of : " & filePath & " | As file can not be found"
   	End If
   
   
End Function

'Retrive name of all the files
'==========================================================================
Function GetFileNames(folderPath)
	
	Dim fileObj, txtFile
	Dim fileName()
	Set fileObj = CheckFolderExistence(folderPath,folderExist)
		Set getFiles = fileObj.GetFolder(folderPath).Files
		fileCount = getFiles.Count
		ReDim fileName(fileCount-1)
		If fileCount = 0 Then
			Print folderPath & "--> Contains no file"
			Else
				For Each childFile In getFiles
					i = i + j
					fileName(i) = childFile.Name
					j = 1
				Next
		End If
	GetFileNames = fileName
End Function

'Deletion of all files under a folder
'==========================================================================
Function DeleteAllFiles(folderPath)
	
	Dim fileObj, txtFile
	Set fileObj = CheckFolderExistence(folderPath,folderExist)
		Set getFiles = fileObj.GetFolder(folderPath).Files
		fileCount = getFiles.Count
		If fileCount = 0 Then
			Print folderPath & "--> Contains no file"
			Else
				For Each childFile In getFiles
					fileName = childFile.Name
					childFile.Delete
					If fileObj.FileExists(fileName) Then
						Print fileName & " : file could not be deleted"
						Else
						Print fileName & " : file is successfully deleted"
					End If
				Next
		End If

	
End Function

'Count the number of files uder a specified parent folder
'==========================================================================
Function CountAllFiles(folderPath,fileCount)
	
	Dim fileObj, txtFile
	Set fileObj = CheckFolderExistence(folderPath,folderExist)
		fileCount = fileObj.GetFolder(folderPath).Files.Count
	
End Function

'Check existence of any folder
'==========================================================================
Function CheckFolderExistence(folderPath,folderExist)

	Dim fileObj
		Set fileObj = CreateObject("Scripting.FileSystemObject")
		If fileObj.FolderExists(folderPath) Then
			folderExist = 1
			Print "Folder Found : " & folderPath
			Else
			folderExist = 0
			Print "Folder Not Found : " & folderPath
		End If
		Set CheckFolderExistence = fileObj
		
End Function

'Deletion of any folder
'==========================================================================
Function DeleteAFolder(folderPath)

   Dim fileObj, txtFile
   Set fileObj = CheckFolderExistence (folderPath,folderExist)

   	If folderExist = 1 Then 
   		fileObj.DeleteFolder folderPath
   		Print "Successful Deletion of : " & folderPath
   		Else
   		Print "Deletion not Successful of : " & folderPath & " | As folder can not be found"
   	End If
   
   
End Function

'Count the number of subfolders by specifying a parent folder
'==========================================================================
Function CountChildFolder(folderPath,folderCount)
	
   Dim fileObj, txtFile
   Set fileObj = CheckFolderExistence (folderPath,folderExist)

   		Set getFolder = fileObj.GetFolder(folderPath)
   		folderCount = getFolder.SubFolders.Count
   
End Function

'Retrive name of all the sub folders
'==========================================================================
Function GetNameChildFolder(folderPath)
	
   Dim fileObj, txtFile
   Dim folderNames()
   Set fileObj = CheckFolderExistence (folderPath,folderExist)

   		Set getFolder = fileObj.GetFolder(folderPath)
   		Set getSubFolder = getFolder.SubFolders
   		folderCount = getSubfolder.Count
   			If folderCount = 0 Then
   				Print "No subfolder found"
   				Else
   				  ReDim folderNames(folderCount-1)
		   			For Each subFol In getSubFolder
		   				i = i + j
		   				folderNames(i) = subFol.Name
		   				j = 1
		   			Next
   			End If

   GetNameChildFolder = folderNames
   
End Function

'Delete all the sub folders
'==========================================================================
Function DeleteChildFolder(folderPath)
	
   Dim fileObj, txtFile
   Set fileObj = CheckFolderExistence (folderPath,folderExist)

   		Set getFolder = fileObj.GetFolder(folderPath)
   		Set getSubFolder = getFolder.SubFolders
   		folderCount = getSubfolder.Count
   			If folderCount = 0 Then
   				Print "No subfolder found"
   				Else
		   			For Each subFol In getSubFolder
		   				folderName = subFol.Name
						subFol.Delete
							If fileObj.FolderExists(folderName)  Then
								Print folderName & " : folder could not be deleted"
								Else
								Print folderName & " : folder is successfully deleted"
							End If						
		   			Next
   			End If
   
End Function
