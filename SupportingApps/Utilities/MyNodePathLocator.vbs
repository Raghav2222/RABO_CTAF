' Function to find the absolute path in an xml




'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Dim strNodeName, arr(50), firstNode

On error resume next


Msgbox "Please READ and proceed -"&vblf&vblf&"This tool will help you to find the Absolute Path of a node in an XML."&vblf&vblf&"You need to provide the exact drive location of the xml file (e.g. C:\abc\xyz.XML) in the prompt box."&vblf&vblf&"Mention the Node name you want to search in that XML."&vblf&vblf&"It will also prompt if you want to save the Result in text!!!",48, "MyNodePathLocator v0.1"

strXMLPath = trim(inputbox("Please enter your XML location - full path         (e.g. C:\abc\xyz.xml)", "XML Location"))
'strXMLPath = "G:\XMLAbsolutePath\Reserve_BBAN.xml"
''strXMLPath = "G:\XMLAbsolutePath\SampleXML.xml"

strNodeName = trim(inputbox("Enter the NODE name you want to search in the XML", "Node Search"))
'strNodeName = "ns1:ActCd"
''strNodeName = "citem"

'for Yes - 6, for No-7
strSaveTxt = Msgbox ("Would you like to Save the result to C:\MyNodePathLocatorInfo.txt location??", 4, "Saving to notepad")

'Msgbox strNodeName

dim xmlDom
set xmlDom = createobject("MSXML.DOMDocument")
xmlDom.async = false
'xmlDom.load ("G:\XMLAbsolutePath\Reserve_BBAN.xml")
xmlDom.load (strXMLPath)


'strNodeName = "ns1:ActCd"
'strNodeName = "PmtId"

If strSaveTxt=6 Then
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Check that the strDirectory folder exists
	If objFSO.FileExists("C:\MyNodePathLocatorInfo.txt") Then
	   objFSO.GetFile("C:\MyNodePathLocatorInfo.txt")
	Else
	   objFSO.CreateTextFile("C:\MyNodePathLocatorInfo.txt")   
	End If
	
	Set objTxtFile = objFSO.OpenTextFile("C:\MyNodePathLocatorInfo.txt", 8, true)
	
	DtTime = Now()
	strIntro = "--------------------------------------------------------------------------------------------------"
	
	
	If strSaveTxt=6 Then
		objTxtFile.Writeline(strIntro)
		''objTxtFile.wrvbCrLf
	''strNewText = vbCrLf & vbCrLf & vbCrLf
		objTxtFile.Writeline(DtTime)
		objTxtFile.Writeline(strXMLPath)
		objTxtFile.Writeline(strIntro)
	End If
	
End If


'firstNode = "soapenv:Envelope"
firstNode= FindParentNode(strXMLPath)
'firstNode = "SOAP-ENV:Envelope"
temp = strNodeName

Set x=xmlDom.getElementsByTagName(strNodeName)

If x.length = 0 Then
		Msgbox "Oops!! Sorry...your Node is not found in the XML :(", 16, "Node not found"		
End If
NodeCount = x.length

If x.length > 0 Then
	'xmlDom.ownerDocument
	For i = 0 to x.length-1
		Set x=xmlDom.getElementsByTagName(strNodeName)
		str = x.item(i).parentNode.nodeName
		temp = str&"\"&temp
		While str<>firstNode
			str = ImmediateParentRecurrisive(str, i)
			temp = str&"\"&temp
			'Msgbox temp
		Wend
			arr(i) = temp
			'Save the info into C:\MyNodePathLocatorInfo.txt
			If strSaveTxt=6 Then				
				objTxtFile.WriteLine(temp&vblf)
			End If
			temp = strNodeName
	Next
	
	'Msgbox  x.parentNode.nodeName'
	'temp = null
	'For itr = 0 to Ubound(arr)
	'	If arr(itr) <>"" Then
	'		temp = temp&vbLf&arr(itr)
	'			'Msgbox arr(itr)	
	'	End If
	'Next
	
	
	temp = "Total number of - "&strNodeName&" - nodes in the XML :  "& NodeCount&vbLf
	temp1 = temp
	For itr = 0 to Ubound(arr)
		If arr(itr) <>"" Then			'
			temp = temp&vbLf&arr(itr)&vbLf
				'Msgbox arr(itr)	
		End If
	Next

	''Close notepad
		If strSaveTxt=6 Then
				objTxtFile.Close
		End If

	msgbox temp,64,"Absolute node path in XML"

	If strSaveTxt=6 Then
		Msgbox "Result has been saved to C:\MyNodePathLocatorInfo.txt file.",64, "Information on text file"
	End If
	Set xmlDom = null
	Set x = null
	
End If



'--------------------------------------------------------------------------
' Function to find the parent of a node
'--------------------------------------------------------------------------

Function ImmediateParentRecurrisive(strNode, i)

On error resume next
	
	Set x=xmlDom.getElementsByTagName(strNode)
	For y=0 to x.length-1
		str = x.item(y).parentNode.nodeName
		ImmediateParentRecurrisive = str
	Next	

	If err.number <> 0 then
		'msgbox "worked"
		else
		'msgbox "Error is - "&err.description
	End if

End Function



'-----------------------------------------------------------------------------------
'Function to find the Parent node of an xml
'-----------------------------------------------------------------------------------

Function FindParentNode(strXMLPath)

	On error resume next
	set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.async = false
	xmlDoc.load(strXMLPath) 
	
	set root = xmlDoc.documentElement
	'Msgbox "Root: " + root.nodeName

	FindParentNode = root.nodeName
   
End Function

















