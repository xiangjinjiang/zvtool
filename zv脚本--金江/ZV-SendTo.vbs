on error resume next

'by 金江 


'基本参数
'------------------------------begin----------------------------------------
'钱包地址
addr=""
'连接节点
node="https://api.jokechain.cc:8101"
'钱包密码
pass=""

'-------------------------------end-----------------------------------------
weatherData = GetHttpPage("https://tianquanexplorer.zvchain.io:8000/addressdetail?address="+addr+"&pagetype=1","UTF-8")
indedStart= InStr(weatherData,"balance")
indexEnd=InStr(weatherData,"nonce")
tasBalance= mid(weatherData,indedStart+9,indexEnd-indedStart-11)
value=round(CDbl(tasBalance/1000000000))-1

account=inputBox("请输入对方账户","ZVChain")
If account=vbEmpty or account="" Then
	WScript.Quit
end if
if len(account)<>66 then
	msgbox "账户输入有误!",,"ZVChain"
	WScript.Quit
end if
sendTo=account
ans=inputBox("账户剩余数量【"+CStr(value)+"】","ZVChain",CStr(value))
If ans=vbEmpty or ans="" Then
	WScript.Quit
end if
value=ans
set wshshell=CreateObject("Wscript.shell")
set fso=Createobject("Scripting.FileSystemObject")
wshshell.run "cmd.exe"
wscript.sleep 100

str="gzv.exe console"
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
wshshell.sendkeys "sendtx -to "+sendTo+" -value "+CStr(value)
wshshell.sendkeys "{enter}"

wscript.sleep 100
wshshell.sendkeys "exit"
wshshell.sendkeys "{enter}"
msgbox "转账完成,请手动查询交易是否成功!",,"ZVChain"
wshshell.sendkeys "exit"
wshshell.sendkeys "{enter}"

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