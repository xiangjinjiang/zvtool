on error resume next

'by �� 


'��������
'------------------------------begin----------------------------------------
'���ӽڵ�
node="https://api.jokechain.cc:8101"
'-------------------------------end-----------------------------------------
root="F:\zv\ba2b"
'------------��ȡ���ͷ����˻���Ϣ------------
Const ForReading = 1
Dim account  
Dim fs, ts
set fs = CreateObject("Scripting.FileSystemObject")
set ts = fs.OpenTextFile(root+"\fenfashuliang_fasongfang.txt", ForReading)
	message = ts.ReadLine       '��ʹ�ǿ��У�Ҳ�ᱻ��һ��  
	ss=split(message,",")
	account=ss(0)
	pass=ss(1)
	path=ss(2)
ts.Close
set ts = Nothing
set fs = Nothing
if account = "" Then
        msgbox "�˻�����Ϊ��!",,"ZVChain"
        WScript.Quit
End if
'------------End ��ȡ���ͷ����˻���Ϣ------------

addr=account

weatherData = GetHttpPage("https://tianquanexplorer.zvchain.io:8000/addressdetail?address="+addr+"&pagetype=1","UTF-8")
indedStart= InStr(weatherData,"nonce")
indexEnd=InStr(weatherData,"type")
nonce= Cint(mid(weatherData,indedStart+7,indexEnd-indedStart-9))


'--------------account=inputBox("������Է��˻�","ZVChain")
If account=vbEmpty or account="" Then
	WScript.Quit
end if
if len(account)<>66 then
	msgbox "�˻���������!",,"ZVChain"
	WScript.Quit
end if

set wshshell=CreateObject("Wscript.shell")
set fso=Createobject("Scripting.FileSystemObject")
wshshell.run "cmd.exe"
wscript.sleep 200

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
str="balance -addr "+addr
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
'------------��ȡ��ת���˻�------------------
set fs = CreateObject("Scripting.FileSystemObject")
set ts = fs.OpenTextFile(root+"\fenfashuliang_jieshoufang.txt", ForReading)
Do Until ts.AtEndOfStream
	message = ts.ReadLine       '��ʹ�ǿ��У�Ҳ�ᱻ��һ��  
	ss=split(message,",")
	sendTo=ss(0)
	sendToValue=ss(1)

	nonce=nonce+1
	wshshell.sendkeys "sendtx -to "+sendTo+" -value "+sendToValue+" -nonce "+CStr(nonce)
	wshshell.sendkeys "{enter}"
	wscript.sleep 100
Loop

ts.Close
set ts = Nothing
set fs = Nothing
'�˳�
wscript.sleep 100
wshshell.sendkeys "exit"
wshshell.sendkeys "{enter}"
'msgbox "ת�����,���ֶ���ѯ�����Ƿ�ɹ�!",,"ZVChain"
wshshell.sendkeys "exit"
wshshell.sendkeys "{enter}"
msgbox "ת�����,���ֶ���ѯ�����Ƿ�ɹ�!",,"ZVChain"

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