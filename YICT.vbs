#$language = "VBScript"
#$interface = "1.0"

Sub Main
    '打开保存设备管理地址以及密码的文件
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso,file1,line,logfile,params,ipaddr,username,password
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file1 = fso.OpenTextFile("D:\YICT盐田国际\自动巡检脚本\hostlist.txt",Forreading, False)    
    crt.Screen.Synchronous = True
    DO While file1.AtEndOfStream <> True
       '读出每行
       line = file1.ReadLine
       '分离每行的参数 IP地址 密码 En密码
       params = Split (line)
	   ipaddr = params(0)
	   username = params(1)
	   password = params(2)
	   
	   
		
			'调用Telnet_Login函数
		Telnet_Login ipaddr,username,password
		crt.Screen.Send "show privilege" & vbCr
		crt.Screen.Send vbCr
		crt.Screen.Send "show ver | in uptime"
		outPut = crt.Screen.ReadString ("show ver | in uptime")
	
		if InStr(output,">")  Then
			PrivilegeLevel ">",ipaddr
			crt.Session.Disconnect
		Elseif InStr(outPut,"#") Then
			PrivilegeLevel "#",ipaddr
			crt.Session.Disconnect
		End if
	
	 	 
    loop
    crt.Screen.Synchronous = False
	'crt.Quit 	
End Sub

Function Telnet_Login(ipaddress,username,password)

	   'Telnet到这个设备上
        crt.Session.Connect "/TELNET " & ipaddress

	   '输入telnet密码
	    crt.Screen.WaitForString "Username:"
        crt.Screen.Send username & vbcr
        crt.Screen.WaitForString "Password:"
        crt.Screen.Send password & vbcr
       '进特权模式
       'crt.Screen.Send "enable" & vbcr
       'crt.Screen.WaitForString "Password:"
       'crt.Screen.Send params(3) & vbcr	
End Function





Function PrivilegeLevel(str1,ipaddress)
		 
       'crt.Screen.waitForString "#"
	   
	   '从"show run | in host" 的结果中提取用户名，并代入文件中
	   crt.Screen.Send "show ver | in uptime" & vbCr 
	   Result = crt.Screen.ReadString(str1)

       'msgbox Result
	   
	   '第一种用换行来分割结果，获取hostname R1,但是返回的是数组的元素，无法加入到文件名中，需要Mid函数再次提取
	   strHN = Split(Result,vbCr)(1) 
       'msgbox strHN
	   '''msgbox Mid(Result, 21)
	   '''strHN = Split(Mid(Result, 21),vbCr)(1)	'第2种方法，用Mid函数提取，在用Split提取R1   
	   '''msgbox strHN
	   hn = Split(strHN)(0)
	   'msgbox hn	   
	   HN = Mid(hn,2)  '''hn中包含两行，第一行为空，但是第二行为HOSTNAME，但是第一行要占用一个字符，第二行从2开始。
  	   'msgbox HN
	   
       '设置文件存放目录及文件名，文件名中包含日期
       logfile = "D:\YICT盐田国际\自动巡检-Log\" & ipaddress & " " & HN & " .log"
	  
	   '开启记录日志
	   crt.Session.LogFileName = logfile
	   crt.Session.Log(true) 	   
	   
	   SendCommand(str1)	 
End Function

Function SendCommand(str1)
	
	commands = Array("terminal length 0","show ver"," show env ala","show env stat","show process cpu ","show process memory","show module","show logging","show clock" ,"show ntp status")
	
	for each c in commands
		crt.Screen.Send c & vbCr 
		'crt.Sleep 1000
		crt.Screen.waitForString str1
	Next
	
End Function