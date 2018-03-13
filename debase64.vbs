set args = Wscript.Arguments
count = Wscript.Arguments.Count
If count <= 1 Then
	MsgBox "�������������룺"+vbCrLf+"debase64.vbs  Դ�ļ���  Ŀ���ļ���" ,vbInformation, "����÷�"
	wscript.quit
else 
	Call Base64Decode(args(0), args(1))
end If

'����˵����BASE64����
Function Base64Decode(srcFileName, dstFileName)
	set fs = wscript.CreateObject("Scripting.FileSystemObject")
	If not fs.fileexists(args(0)) Then
		MsgBox  "Ҫ������ļ�:["+srcFileName+"]������" ,vbExclamation, "����"
		wscript.quit
	end If

	Dim fileSystemObj,file
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
	Set file = fileSystemObj.OpenTextFile(srcFileName, 1)
	txtVar = file.ReadAll() 
	
    Set xmlDom = CreateObject("Microsoft.XMLDOM")
    Set xmlNode = xmlDom.createElement("MyNode")
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = Replace(txtVar, vbCrLf, "")
    
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 1  'adTypeBinary
    adoStream.Open()
    adoStream.Write(xmlNode.nodeTypedValue)
    adoStream.Position = 0
	Call adoStream.SaveToFile (dstFileName, 2)
    adoStream.Close()
End Function