
'Authored By: Kunal Dey
'Date: 11/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 CreateExcel
'02 AddWorksheet
'03 DeleteWorksheet
'04 RenameWorksheet
'05 MoveTabForwardWorksheet
'06 MoveTabLastWorksheet
'07 AddValWorksheet
'08 ReadValWorksheet
'09 AddFormattedValWorksheet
'10 FindLastRow
'11 FindLastColumn
'12 FindValue
'13 CompareExcelFiles
'14 CopyExcel
'************************************************************
'************************************************************

'Create a Excel File
'==========================================================
Function CreateExcel (xlsPath)

	Set objExl = CreateObject("Excel.Application")

		'To make Excel not visible
		objExl.Application.Visible = false
		objExl.Workbooks.Add


		'Save the Excel file as qtp.xls
		objExl.ActiveWorkbook.SaveAs xlsPath
     
		'close Excel
		objExl.Application.Quit

	Set objExl=nothing

End Function

'Adding Worksheet in an Excel File
'==========================================================
Function AddWorksheet (xlsPath, wsName)
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	
	'New worksheet will be created at the end of the last existing worksheet
	Set objWs = objWb.Sheets.Add( ,objWb.Sheets(objWb.Sheets.Count))
		objWs.Name = wsName
		objExl.ActiveWorkbook.Save
		objExl.DisplayAlerts = False
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Deleting Worksheet in an Excel File
'==========================================================
Function DeleteWorksheet (xlsPath, wsName)
	
	Set objExl = CreateObject("Excel.Application")
		objExl.Workbooks.Open xlsPath
		'Select the worksheet to be deleted
		objExl.Sheets(wsName).Delete
		objExl.DisplayAlerts = False
		objExl.ActiveWorkbook.Save
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Renaming a Worksheet in an Excel File
'==========================================================
Function RenameWorksheet (xlsPath,wsNameOld,wsNameNew)
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	
	'New worksheet will created at the end of the last existing worksheet
	Set objWs = objWb.Worksheets(wsNameOld)
	
		objWs.Name = wsNameNew
		objExl.ActiveWorkbook.Save
		objExl.DisplayAlerts = False 
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Move Worksheet position forward by 1 in an Excel File'
'==========================================================
Function MoveTabForwardWorksheet (xlsPath, wsName)
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	
	objWb.Sheets(objWb.Sheets.Count).Move objWb.Sheets(wsName)
		objExl.ActiveWorkbook.Save
		objExl.DisplayAlerts = False 
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Move Worksheet to last position in an Excel File'
'==========================================================
Function MoveTabLastWorksheet (xlsPath, wsName)
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	
	objWb.Sheets(wsName).Move ,objWb.Sheets( objWb.Sheets.Count )
		objExl.ActiveWorkbook.Save
		objExl.DisplayAlerts = False 
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Add value to a particular cell in an Excel File
'==========================================================
Function AddValWorksheet (xlsPath,wsName,rowNum,colNum,cellVal)
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	Set objWs = objWb.Worksheets(wsName)
	
		objWs.Cells(rowNum, colNum).Value = cellVal
		objExl.ActiveWorkbook.Save
		objExl.DisplayAlerts = False 
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Read value from a particular cell in an Excel File
'==========================================================
Function ReadValWorksheet (xlsPath,wsName,rowNum,colNum,cellVal)
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	Set objWs = objWb.Worksheets(wsName)
	
		cellVal = objWs.Cells(rowNum, colNum).Value
		objExl.DisplayAlerts = False 
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Add formatted value to a particular cell in an Excel File
'==========================================================
Function AddFormattedValWorksheet (xlsPath,wsName,rowNum,colNum,cellVal,isBold,fontSize,fontColor)
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	Set objWs = objWb.Worksheets(wsName)
	
		objWs.Cells(rowNum, colNum).Value = cellVal
		objWs.Cells(rowNum, colNum).Font.Bold = isBold
		objWs.Cells(rowNum, colNum).Font.Size = fontSize
		objWs.Cells(rowNum, colNum).Font.ColorIndex = fontColor
		objExl.ActiveWorkbook.Save
		objExl.DisplayAlerts = False 
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Find the last row
'==========================================================
Function FindLastRow (xlsPath,wsName,getCount)

	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	Set objWs = objWb.Worksheets(wsName)
	objWs.Activate
	objExl.ActiveSheet.UsedRange.Select
	getCount = objExl.Selection.Rows.Count
	objExl.Application.Quit
	Set objExl=nothing

	
End Function

'Find the last Column
'==========================================================
Function FindLastColumn (xlsPath,wsName,getCount)

	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	Set objWs = objWb.Worksheets(wsName)
	objWs.Activate
	objExl.ActiveSheet.UsedRange.Select
	getCount = objExl.Selection.Columns.Count
	objExl.Application.Quit
	Set objExl=nothing

	
End Function

'Find a particular value in an Excel File
'==========================================================
Function FindValue (xlsPath,wsName,findString)
	
	Const xlValues = -4163
	
	Set objExl = CreateObject("Excel.Application")
	Set objWb = objExl.Workbooks.Open (xlsPath)
	Set objWs = objWb.Worksheets(wsName)

		Set objRange = objWs.UsedRange
		
		Set objTarget = objRange.Find(findString)

		If Not objTarget Is Nothing Then
		    print objTarget.AddressLocal(False,False)
		    strFirstAddress = objTarget.AddressLocal(False,False)
		End If
		Do Until (objTarget Is Nothing)
		    Set objTarget = objRange.FindNext(objTarget)
		
		    strHolder = objTarget.AddressLocal(False,False)
		    If strHolder = strFirstAddress Then
		        Exit Do
		    End If
		    print objTarget.AddressLocal (False,False)
		Loop
		objExl.Application.Quit
	Set objExl=nothing
	
End Function

'Compare two Excel files cell by cell
'==========================================================
Function CompareExcelFiles (xlsPath,xlsPath2,wsName,wsName2)
	
	Mismatch=0
	Set objExl = createobject("Excel.Application")
	
	'To make Excel not visible
	objExl.Visible = False
	
	'Open first workbook
	Set Workbook1= objExl.Workbooks.Open(xlsPath)
	
	'Open second workbook
	Set Workbook2= objExl.Workbooks.Open(xlsPath2)
	 
	Set  mysheet1=Workbook1.Worksheets(wsName)
	Set  mysheet2=Workbook2.Worksheets(wsName2)
	 
	'Compare two sheets cell by cell
	For Each cell In mysheet1.UsedRange
	
	'Highlights the cell if  cell values not match
	       If cell.Value <>mysheet2.Range(cell.Address).Value Then
	           'Highlights the cell if  cell values not match
	           cell.Interior.ColorIndex = 3
	              mismatch=1
	       End If
	   Next
	 
	If Mismatch=0 Then
	    Print "No Mismach exists"
	End If

	Workbook1.Activate
	objExl.ActiveWorkbook.Save
	'close the workbooks
	Workbook1.close
	Workbook2.close
	 
	objExl.Quit
	set objExl=nothing
End Function

'Copy one Excel content to another Excel
'==========================================================
Function CopyExcel (xlsPath,xlsPath2,wsName,wsName2)

	Set objExl = createobject("Excel.Application")

		objExl.Visible = False
		Set Workbook1= objExl.Workbooks.Open(xlsPath)
		Set Workbook2= objExl.Workbooks.Open(xlsPath2)
		
		'Copy  the used range of  workbook "qtp1.xls"
			Workbook1.Worksheets(wsName).UsedRange.Copy
			'Paste the copied values in above step in the  A1 cell  of  workbook "qtp2.xls"
			Workbook2.Worksheets(wsName2).Range("A1").PasteSpecial Paste =xlValues
			
			'Save the workbooks
			Workbook1.save
			Workbook2.save
			 
			'close the workbooks
			Workbook1.close
			Workbook2.close
			 
			objExl.Quit
		set objExl=nothing
End Function

'Import Excel file to Datatable
'==========================================================
Function ImportToDT (xlsPath,wsName,sheetDest)
	DataTable.ImportSheet xlsPath, wsName, sheetDest
End Function
