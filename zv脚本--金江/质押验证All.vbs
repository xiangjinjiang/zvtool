on error resume next

'by 金江


'基本参数
'------------------------------begin----------------------------------------
'连接节点
node="https://api.jokechain.cc:8101"
root="F:\zv\ba2b"
'-------------------------------end-----------------------------------------

'------------读取需要质押的账户------------
Const ForReading = 1
Dim fs, ts
'------------End 读取被转账的账户------------
'------------启动cmd-------------------------
set wshshell=CreateObject("Wscript.shell")
set fso=Createobject("Scripting.FileSystemObject")
wshshell.run "cmd.exe"
wscript.sleep 200

'------------读取待转账账户------------------
set fs = CreateObject("Scripting.FileSystemObject")
set ts = fs.OpenTextFile(root+"\zhiyaAll.txt", ForReading)
Do Until ts.AtEndOfStream
    message = ts.ReadLine       '即使是空行，也会被读一次  
	ss=split(message,",")
	addr=ss(0)
	pass=ss(1)
	path=ss(2)
weatherData = GetHttpPage("https://tianquanexplorer.zvchain.io:8000/addressdetail?address="+addr+"&pagetype=1","UTF-8")
indedStart= InStr(weatherData,"balance")
indexEnd=InStr(weatherData,"nonce")
tasBalance= mid(weatherData,indedStart+9,indexEnd-indedStart-11)
stake=round(CDbl(tasBalance/1000000000))-1



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

wscript.sleep 500
wshshell.sendkeys "stakeadd -type 0 -value "+CStr(stake)
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
msgbox "质押完成,请手动查询质押是否成功!",,"质押"

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