
set args = Wscript.Arguments
count = Wscript.Arguments.Count
If count <= 1 Then
	MsgBox "�������������룺"+vbCrLf+"base64.vbs  Դ�ļ���  Ŀ���ļ���" ,vbInformation, "����÷�"
	wscript.quit
else
	Call Base64Encode(args(0), args(1))
end If

'Base64 ����
Function Base64Encode(srcFileName, dstFileName) 
	set fs = wscript.CreateObject("Scripting.FileSystemObject")
	If not fs.fileexists(args(0)) Then
		MsgBox  "Ҫ������ļ�: [ "+srcFileName+" ]������" ,vbExclamation, "����"
		wscript.quit
	end If
	
	Dim stream, xmldom, node
	Set xmldom = CreateObject("Microsoft.XMLDOM")
	Set node = xmldom.CreateElement("tmpNode")
	node.DataType = "bin.base64"
	Set stream = CreateObject("ADODB.Stream")
	stream.Type = 1  'adTypeBinary
	stream.Open()
	stream.LoadFromFile(srcFileName) 
	node.NodeTypedValue = stream.Read()
	stream.Close()
	
    Dim fileSystemObj,txtFile
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fileSystemObj.OpenTextFile(dstFileName, 2, true) 
    txtFile.WriteLine(CStr(node.Text))
    txtFile.Close()
End Function