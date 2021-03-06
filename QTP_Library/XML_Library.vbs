
'Authored By: Kunal Dey
'Date: 13/11/2014

'************************************************************
'************************************************************
'List of Functions This Library Can Handle
'01 CreateXML
'02 GenEnvVar
'03 CompareXML
'04 FindParentNode
'05 FindChildNode
'06 GetAllChildNodes
'07 GetNodeCount
'08 FindTextFromRange
'09 AddChildNode
'10 AddAttribute
'11 AddText
'12 CheckXMlLoading
'13 ReadAllAttributeName
'14 ReadAllAttributeValue
'15 ReadAllAttributeCount
'************************************************************
'************************************************************

'Create an XML File
'Note: The following code will generate a xml file of the following format...
		'<?xml version="1.0"?>
		'-<School>
		'	-<Teacher>
		'		<Student>Tester</Student>
		'	</Teacher>
		'</School>
	'*** The XML can have three levels of elements (nodes). One can use array
	'to add multiple nodes under the same level. The Root Node is fixed but can be
	'named according to will
'=============================================================================
Function CreateXML (rootNode,parentNode(),childNode(),childText(),xmlPath)
	
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	Set objRoot = xmlDoc.createElement(rootNode)  
	xmlDoc.appendChild objRoot  
	
	For pn = 0 To UBound(parentNode)-1
		Set objRecord = xmlDoc.createElement(parentNode(pn)) 
		objRoot.appendChild objRecord 
			
			For cn = 0 To UBound(childNode)-1
				Set objName = xmlDoc.createElement(childNode(cn))
					For ct = 0 To UBound(childText)-1
						objName.Text = childText(ct)
					    objRecord.appendChild objName
					Next		
			Next
	Next
		
	Set objIntro = xmlDoc.createProcessingInstruction ("xml","version='1.0'")  
	xmlDoc.insertBefore objIntro,xmlDoc.childNodes(0)  
	
	xmlDoc.Save xmlPath

End Function

'Generate Environment Variable in XML
'Note: User need to provide Variable Name and Value in Array format as parameter and 
	  'can generate as many variables as he/she wants.
'=============================================================================
Function GenEnvVar (xmlPath,variableNum,nameList(),ValueList())
	
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")  
	Set objRoot = xmlDoc.createElement("Environment")  
	xmlDoc.appendChild objRoot  
	
	For nl = 0 To UBound(nameList)-1
		Set objRecord = xmlDoc.createElement("Variable") 
		objRoot.appendChild objRecord 

			Set objName = xmlDoc.createElement("Name")
				objName.Text = nameList(nl)
				objRecord.appendChild objName
			Set objName = xmlDoc.createElement("Value")
				objName.Text = ValueList(nl)
				objRecord.appendChild objName
	Next
		
	Set objIntro = xmlDoc.createProcessingInstruction ("xml","version='1.0'")  
	xmlDoc.insertBefore objIntro,xmlDoc.childNodes(0)  
	
	xmlDoc.Save xmlPath
	
End Function

'Compare two xml files and print mismatch result
'=============================================================================
Function CompareXML (xmlPath1,xmlPath2)
'	Dim description, filepath
	Set xmlDoc1 = CheckXMlLoading(xmlPath1)
	Set xmlDoc2 = CheckXMlLoading(xmlPath2)
	
		Set ElemList1= xmlDoc1.DocumentElement.ChildNodes
		Set ElemList2= xmlDoc2.DocumentElement.ChildNodes

			If ElemList1.length=ElemList2.length Then' check weather both xml file has same number of childnodes
			  Print "Both XML files have same number of Child nodes"
			
			   For i = 0 to ElemList1.length-1
			
			       If ElemList1.item(i).Text=ElemList2.item(i).Text Then
			       Else
			          Print "Mismatch Found at Element#" & i &" file 1 contains:" & ElemList1.item(i).Text & "|| file 2 contains:" & ElemList2.item(i).Text
			     End If
			   Next
			Else
			  Print "XML files can not be compared as mismatch found in the number of childnodes"
			End If
End Function

'Find the parent node of any defined node
'elemIndex defines the index position of the node. If same node exits multiple 
'time, user need to provide index of the node to get the exact parent name
'=============================================================================
Function FindParentNode(xmlPath,childNode,parNode,elemIndex)
	
	Set xmlDoc = CheckXMlLoading(xmlPath)
	If elemIndex = "" Then
		elemIndex = "0"
	End If
	Set getElem = xmlDoc.getElementsByTagName(childNode)(elemIndex)
	set getParent = getElem.parentNode
	parNode = getParent.nodeName

End Function

'Find all the child nodes of any defined node
'Note: Returns an array, while calling this function follow the code snippet below...
	'Dim nodeList
	'nodeList = FindChildNode ("C:\XLS\School.xml","Teacher")
	'For i = 0 To UBound(nodeList)-1
	'	print nodeList(i)
	'Next
'=============================================================================
Function FindChildNode(xmlPath,parentNode)
	
	Set xmlDoc = CheckXMlLoading(xmlPath)
	Dim nodeList()
	
			Set nodes = xmlDoc.selectNodes("//*/" & parentNode)
			For i = 0 to nodes.length-1
    			Set getChilds = nodes.item(i).childNodes
    				ReDim nodeList(getChilds.length)
		    			If getChilds.length>0 Then
		    				For j = 0 To getChilds.length-1
		    					nodeList(j) = getChilds.item(j).nodeName
		    				Next
		    			Else
		    				Print "No Child Node Found"
		    			End If
			Next
	FindChildNode = nodeList
	
End Function

'Find all the child nodes of root node
'Note: Returns an array, while calling this function follow the code snippet below...
	'Dim nodeList
	'nodeList = GetAllChildNodes ("C:\XLS\School.xml")
	'For i = 0 To UBound(nodeList)-1
	'	print nodeList(i)
	'Next
'=============================================================================
Function GetAllChildNodes(xmlPath)

	Set xmlDoc = CheckXMlLoading(xmlPath)
		Dim nodelist()
			set nodes = xmlDoc.selectNodes("//*")
		
			ReDim nodelist(nodes.length)
			For i = 0 to nodes.length-1
    			nodelist(i) = (nodes.item(i).nodeName)
			Next
	
	GetAllChildNodes = nodelist
End Function

'Count all the defined nodes
'=============================================================================
Function GetNodeCount(xmlPath, findNode, countNode)
	
	Set xmlDoc = CheckXMlLoading(xmlPath)
	Set nodes = xmlDoc.selectNodes("//*/" & findNode)
	countNode = nodes.length

End Function

'Determine existance of a value between a text range
'Note: This method finds a text between a given text range. Consider the following XML snippet
'-<SCHOOL>
'-<CLASS>
'-<SUBJECT>
''	<SUB>Math</SUB>
'	<SUB>English</SUB>
'	<SUB>Geaography</SUB>
'	<SUB>Biology</SUB>
'	<SUB>Histoy</SUB>
'</SUBJECT>
'</CLASS>
'</SCHOOL>
'Let's assume we have to find if Geography is present between English and Biology
'Call the method this way... 
'FindTextFromRange (<xmlPath>,"SUB","English","Biology","Geography")
'=============================================================================
Function FindTextFromRange(xmlPath,targetNode,startString,endString,findString)
	
	Set xmlDoc = CheckXMlLoading(xmlPath)
	Set oNodeList = xmlDoc.selectNodes("//*/" & targetNode)
		
		ReDim items(oNodeList.Length - 1)
		i = 0
		For Each node In oNodeList
		  items(i) = node.text
		  i = i + 1
		Next
		
		txt = Join(items, vbLf)
		
		Set msg = New RegExp
			msg.Pattern = startString & "([\s\S]*?)" & endString
		
		For Each m In msg.Execute(txt)
		  If InStr(m.SubMatches(0), findString) > 0 Then
		    Print findString & " Can be found between " & startString & " and " & endString
		  Else
		    Print findString & " Can not be found between " & startString & " and " & endString
		  End If
		Next
	
End Function

'Add a child node uder a defined node
'=============================================================================
Function AddChildNode(xmlPath,newNode,newNodeText,targetNode,nodeIndex)

	Set xmlDoc = CheckXMlLoading(xmlPath)
	Set newElem = xmlDoc.createNode(1,newNode,"")
	If nodeIndex = "" Then
		nodeIndex = "0"
	End If
	Set getNode = xmlDoc.getElementsByTagName(targetNode)(nodeIndex)
	newElem.Text = newNodeText
	getNode.appendChild(newElem)
	xmlDoc.save(xmlPath)
	
End Function

'Add attribute to any specified node
'=============================================================================
Function AddAttribute(xmlPath,targetNode,nodeIndex,atrName,atrValue)

	Set xmlDoc = CheckXMlLoading(xmlPath)
	Set getNode = xmlDoc.getElementsByTagName(targetNode)(nodeIndex)
	getNode.setAttribute atrName,atrValue
	xmlDoc.save(xmlPath)
	
End Function

'Add text to any specified node
'=============================================================================
Function AddText(xmlPath,targetNode,nodeIndex,nodeText)

	Set xmlDoc = CheckXMlLoading(xmlPath)
	If nodeIndex = "" Then
		nodeIndex = "0"
	End If
	Set getNode = xmlDoc.getElementsByTagName(targetNode)(nodeIndex)
	getNode.text = nodeText
	xmlDoc.save(xmlPath)
	
End Function

'Verify if the XML file is loaded successfully or not
'=============================================================================
Function CheckXMlLoading(xmlPath)
		
		Dim xmlDoc
		Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
		xmlDoc.load(xmlPath)
		Dim nodelist()
		'	Checking if document is loaded successfuly 
			If xmlDoc.Load(xmlPath) Then
			   	Print (xmlPath & " is loaded successfully.")
			Else
				Print (xmlPath & " could not be loaded.")
			End If
		Set CheckXMlLoading = xmlDoc
		
End Function

'Retrieve all available attributes name of a defined node
'=============================================================================
Function ReadAllAttributeName(xmlPath,targetNode,nodeIndex)

		Dim xmlDoc
		Dim attrName()
		Set xmlDoc = CheckXMlLoading(xmlPath)
			If nodeIndex = "" Then
				nodeIndex = "0"
			End If
				Set getNode = xmlDoc.getElementsByTagName(targetNode)(nodeIndex)
'				attrCount = getNode.Attributes.Length
				ReDim attrName(getNode.Attributes.Length)
				Dim objAttributeNode
				Dim i
				i = 0
					For Each  objAttributeNode in getNode.Attributes
						i = i + j
						attrName(i) = objAttributeNode.nodeName
						j = 1		
					Next
				ReadAllAttributeName = attrName
		
End Function

'Retrieve all available attributes value of a defined node
'=============================================================================
Function ReadAllAttributeValue(xmlPath,targetNode,nodeIndex)

		Dim xmlDoc
		Dim attrVal()
		Set xmlDoc = CheckXMlLoading(xmlPath)
			If nodeIndex = "" Then
				nodeIndex = "0"
			End If
				Set getNode = xmlDoc.getElementsByTagName(targetNode)(nodeIndex)
'				attrCount = getNode.Attributes.Length
				ReDim attrVal(getNode.Attributes.Length)
				Dim objAttributeNode
				Dim i
				i = 0
					For Each  objAttributeNode in getNode.Attributes
						i = i + j
						attrVal(i) =objAttributeNode.nodeValue
						j = 1		
					Next
				ReadAllAttributeValue = attrVal
		
End Function

'Retrieve all available attributes count of a defined node
'=============================================================================
Function ReadAllAttributeCount(xmlPath,targetNode,nodeIndex,attrCount)

		Dim xmlDoc
		Set xmlDoc = CheckXMlLoading(xmlPath)
			If nodeIndex = "" Then
				nodeIndex = "0"
			End If
				Set getNode = xmlDoc.getElementsByTagName(targetNode)(nodeIndex)
				attrCount = getNode.Attributes.Length

End Function

