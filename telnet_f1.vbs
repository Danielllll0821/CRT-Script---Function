#$language = "VBScript"
#$interface = "1.0"
'本脚本示范：从一个文件里面自动读取设备IP地址，密码等，自动输入巡检命令，并记录日志到文件。
'  此版本使用的文件名是提取hostname信息，并作为文件名。
'  1、在用户模式下，通过show ver | in uptime来提取hostname信息；此处因cisco系统版本不同，
'     会有部分结果的第一个字符为空，导致获取不到正确的用户名。
'  2、分别使用Split(str,vbCr)和Mid()函数获得hostname R1;
'  3、使用Mid()函数获得hostname。注意Split(str,vbCr)函数获得的结果,第一个元素为空""，因此不能之间应用到文件名中。

Sub Main
    '打开保存设备管理地址以及密码的文件
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso,file1,line,logfile,params
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file1 = fso.OpenTextFile("d:\device.txt",Forreading, False)    
    crt.Screen.Synchronous = True
    DO While file1.AtEndOfStream <> True
       '读出每行
       line = file1.ReadLine
       '分离每行的参数 IP地址 密码 En密码
       params = Split (line)

	   If params(4) = 0	Then
		       'Telnet到这个设备上
        crt.Session.Connect "/TELNET " & params(0)
	   '----------------------------------------------------------
	   '创建目录存放日志文件,根据ip地址命名文件	 
	   'logfile = "d:\logfile\" & params(0) & "-%Y%M%D%h%m%s.txt"
	   '----------------------------------------------------------
	    '输入telnet密码
	    crt.Screen.WaitForString "Username:"
        crt.Screen.Send params(1) & vbcr
        crt.Screen.WaitForString "Password:"
        crt.Screen.Send params(2) & vbcr
       '进特权模式
       'crt.Screen.Send "enable" & vbcr
       'crt.Screen.WaitForString "Password:"
       'crt.Screen.Send params(3) & vbcr
	    PrivilegeLevel(">") 
		
		'crt.Quit
        crt.Session.Disconnect
	   
		
		
		Else
		       'Telnet到这个设备上
        crt.Session.Connect "/TELNET " & params(0)
	   '----------------------------------------------------------
	   '创建目录存放日志文件,根据ip地址命名文件	 
	   'logfile = "d:\logfile\" & params(0) & "-%Y%M%D%h%m%s.txt"
	   '----------------------------------------------------------
	           '输入telnet密码
	    crt.Screen.WaitForString "Username:"
        crt.Screen.Send params(1) & vbcr
        crt.Screen.WaitForString "Password:"
        crt.Screen.Send params(2) & vbcr
       '进特权模式
       'crt.Screen.Send "enable" & vbcr
       'crt.Screen.WaitForString "Password:"
       'crt.Screen.Send params(3) & vbcr
	      

		PrivilegeLevel("#")
		'crt.Quit
		crt.Session.Disconnect
	   
		End If
		
	 	 
    loop
    crt.Screen.Synchronous = False
	'crt.Quit 	
End Sub



Function PrivilegeLevel(str1)
		 
       'crt.Screen.waitForString "#"
	   
	   '从"show run | in host" 的结果中提取用户名，并代入文件中
	   crt.Screen.Send "show ver | in uptime" & vbCr 
	   Result = crt.Screen.ReadString(str1)	   
       
	   '第一种用换行来分割结果，获取hostname R1,但是返回的是数组的元素，无法加入到文件名中，需要Mid函数再次提取
	   strHN = Split(Result,vbCr)(1) 
       '''msgbox strHN
	   '''msgbox Mid(Result, 21)
	   '''strHN = Split(Mid(Result, 21),vbCr)(1)	'第2种方法，用Mid函数提取，在用Split提取R1   
	   '''msgbox strHN
	   hn = Split(strHN)(0)
	   '''msgbox hn	   
	   HN = Mid(hn,2)  '''hn中包含两行，第一行为空，但是第二行为HOSTNAME，但是第一行要占用一个字符，第二行从2开始。
  	   'msgbox HN
	   
       '设置文件存放目录及文件名，文件名中包含日期
       logfile = "d:\logfile\" & HN & "_%Y.%M.%D_%h.%m.%s_.log"
	  
	   '开启记录日志
	   crt.Session.LogFileName = logfile
	   crt.Session.Log(true) 	   
	   
	   SendCommand(str1)
	   
       '备份完成后退出
       crt.Screen.waitForString str1,2
End Function

Function SendCommand(str1)
	
	commands = Array("terminal length 0","show ver"," show env ala","show env stat","show process cpu ","show process memory","show module","show logging","show clock" ,"show ntp status")
	
	for each c in commands
		crt.Screen.Send c & vbCr
		crt.Screen.waitForString str1
	Next
	
End Function