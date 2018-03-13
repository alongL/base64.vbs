
set args = Wscript.Arguments
count = Wscript.Arguments.Count
If count <= 1 Then
	MsgBox "在命令行中输入："+vbCrLf+"base64.vbs  源文件名  目标文件名" ,vbInformation, "软件用法"
	wscript.quit
else
	Call Base64Encode(args(0), args(1))
end If

'Base64 编码
Function Base64Encode(srcFileName, dstFileName) 
	set fs = wscript.CreateObject("Scripting.FileSystemObject")
	If not fs.fileexists(args(0)) Then
		MsgBox  "要处理的文件: [ "+srcFileName+" ]不存在" ,vbExclamation, "错误"
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