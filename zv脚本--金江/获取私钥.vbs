on error resume next

'by 金江 


'基本参数
'------------------------------begin----------------------------------------
 
 
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
set ts = fs.OpenTextFile(root+"\zengjiazhiya.txt", ForReading)
Do Until ts.AtEndOfStream
    message = ts.ReadLine       '即使是空行，也会被读一次  
	ss=split(message,",")
	addr=ss(0)
	pass=ss(1)
	path=ss(2)
	
	str="gzv.exe --keystore "+path+" console"
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
	wshshell.sendkeys "exportkey -addr  "+addr
	wshshell.sendkeys "{enter}"

	wscript.sleep 100
	wshshell.sendkeys "exit"
	wshshell.sendkeys "{enter}"
	wscript.sleep 100	 
Loop
ts.Close
set ts = Nothing
set fs = Nothing
'推出cmd
'wshshell.sendkeys "exit"
'wshshell.sendkeys "{enter}"
msgbox "质押完成,请手动查询质押是否成功!",,"质押"

 