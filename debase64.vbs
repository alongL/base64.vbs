set args = Wscript.Arguments
count = Wscript.Arguments.Count
If count <= 1 Then
	MsgBox "在命令行中输入："+vbCrLf+"debase64.vbs  源文件名  目标文件名" ,vbInformation, "软件用法"
	wscript.quit
else 
	Call Base64Decode(args(0), args(1))
end If

'函数说明：BASE64解码
Function Base64Decode(srcFileName, dstFileName)
	set fs = wscript.CreateObject("Scripting.FileSystemObject")
	If not fs.fileexists(args(0)) Then
		MsgBox  "要处理的文件:["+srcFileName+"]不存在" ,vbExclamation, "错误"
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