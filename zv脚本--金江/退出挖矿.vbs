on error resume next

'by �� 


'��������
'------------------------------begin----------------------------------------
'���ӽڵ�
node="https://api.jokechain.cc:8101"
root="F:\zv\ba2b"
'-------------------------------end-----------------------------------------

'------------��ȡ��Ҫ��Ѻ���˻�------------
Const ForReading = 1
Dim fs, ts
'------------End ��ȡ��ת�˵��˻�------------
'------------����cmd-------------------------
set wshshell=CreateObject("Wscript.shell")
set fso=Createobject("Scripting.FileSystemObject")
wshshell.run "cmd.exe"
wscript.sleep 200

'------------��ȡ��ת���˻�------------------
set fs = CreateObject("Scripting.FileSystemObject")
set ts = fs.OpenTextFile(root+"\jiedong.txt", ForReading)
Do Until ts.AtEndOfStream
    message = ts.ReadLine       '��ʹ�ǿ��У�Ҳ�ᱻ��һ��  
	ss=split(message,",")
	addr=ss(0)
	pass=ss(1)
	path=ss(2)
	stake=0
	


str="gzv.exe --keystore "+path+" console"
s=mid(str,1,len(str))
wshshell.sendkeys s
wshshell.sendkeys "{enter}"

wscript.sleep 100
str="connect -host "+node
s=mid(str,1,len(str))
wshshell.sendkeys s
wshshell.sendkeys "{enter}"

wscript.sleep 100
str="unlock -addr "+addr
s=mid(str,1,len(str))
wshshell.sendkeys s
wshshell.sendkeys "{enter}"

wscript.sleep 1000
str=pass 
for i=0 to len(str)-1
    a=mid(str,i+1,1)
	wshshell.sendkeys a
	wscript.sleep 100
Next
wshshell.sendkeys "{enter}"

wscript.sleep 100
wshshell.sendkeys "minerabort -type 0"
'wshshell.sendkeys "minerabort -type 0 -gaslimit 5000"
wshshell.sendkeys "{enter}"

wscript.sleep 100
wshshell.sendkeys "exit"
wshshell.sendkeys "{enter}"
wscript.sleep 100
Loop
ts.Close
set ts = Nothing
set fs = Nothing

'wshshell.sendkeys "exit"
'wshshell.sendkeys "{enter}"
msgbox "�ⶳ���,���ֶ���ѯ�ⶳ�Ƿ�ɹ�!",,"�ⶳ"

Function GetHttpPage(url, charset) 
    Dim http 
    Set http = CreateObject("Msxml2.ServerXMLHTTP")
http.setOption(2) = 13056
     http.Open "GET", url, false
    http.Send() 
    If http.readystate<>4 Then
        Exit Function 
    End If 
    GetHttpPage = BytesToStr(http.ResponseBody, charset)
    Set http = Nothing
End function

Function BytesToStr(body, charset)
    Dim objStream
    Set objStream = CreateObject("Adodb.Stream")
    objStream.Type = 1
    objStream.Mode = 3
    objStream.Open
    objStream.Write body
    objStream.Position = 0
    objStream.Type = 2
    objStream.Charset = charset
    BytesToStr = objStream.ReadText 
    objStream.Close
    Set objStream = Nothing
End Function