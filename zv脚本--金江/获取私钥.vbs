on error resume next

'by �� 


'��������
'------------------------------begin----------------------------------------
 
 
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
set ts = fs.OpenTextFile(root+"\zengjiazhiya.txt", ForReading)
Do Until ts.AtEndOfStream
    message = ts.ReadLine       '��ʹ�ǿ��У�Ҳ�ᱻ��һ��  
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
'�Ƴ�cmd
'wshshell.sendkeys "exit"
'wshshell.sendkeys "{enter}"
msgbox "��Ѻ���,���ֶ���ѯ��Ѻ�Ƿ�ɹ�!",,"��Ѻ"

 