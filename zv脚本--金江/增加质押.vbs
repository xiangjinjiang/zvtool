on error resume next

'by 金江 


'基本参数
'------------------------------begin----------------------------------------
'连接节点
node="https://api.jokechain.cc:8101"
root="F:\zv\ba2b"
maxzhiya=2500
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
set ts = fs.OpenTextFile(root+"\zengjiazhiya.txt", ForReading)
Do Until ts.AtEndOfStream
    message = ts.ReadLine       '即使是空行，也会被读一次  
	ss=split(message,",")
	addr=ss(0)
	pass=ss(1)
	path=ss(2)
	'stake=ss(3)
	set ZVObject=getStake(addr)
	curBalance=round(CDbl(ZVObject.balance/1000000000))-1
	curStake=ZVObject.stake
	if maxzhiya>curStake and curBalance>1  then
		if maxzhiya-curStake>= curBalance then
			stake=curBalance
		else
			stake=maxzhiya-curStake
		end if

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
	end if
Loop
ts.Close
set ts = Nothing
set fs = Nothing

'wshshell.sendkeys "exit"
'wshshell.sendkeys "{enter}"
msgbox "质押完成,请手动查询质押是否成功!",,"质押"

class ZVC
	dim balance,stake
end class

function getStake(addr)
	Set html = CreateObject("htmlfile")
	Set http = CreateObject("Msxml2.ServerXMLHTTP")
 
	http.open "GET", "https://tianquanexplorer.zvchain.io:8000/addressdetail?address="+addr+"&pagetype=1", False
	http.send
	strHtml = http.responseText ' 得到数据
 
	Set window = html.parentWindow
	window.execScript "var json = " & strHtml, "JScript" ' 解析 json
 
	Set zvjson = window.json ' 获取解析后的对象
	set oClass=new ZVC
	oClass.balance=zvjson.data.balance
	oClass.stake = zvjson.data.stake
	set getStake =oClass
end function