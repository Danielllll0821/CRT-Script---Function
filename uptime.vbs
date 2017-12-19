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
    Dim fso,file1,file2,f_w,line,logfile,params,ipaddr,username,password
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file1 = fso.OpenTextFile("d:\device.txt",Forreading, False)
	
		'Set file2 = fso.CreateTextFile("d:\uptime.txt")
		Set f_w = fso.OpenTextFile("d:\uptime.txt", ForAppending, True)
	
    'crt.Screen.Synchronous = True
    DO While file1.AtEndOfStream <> True
       '读出每行
       line = file1.ReadLine
       '分离每行的参数 IP地址 密码 En密码
       params = Split (line)
	   'If params(4) = 0	Then
		
	     crt.Session.Connect "/TELNET " & params(0)
	   

	   
       '输入telnet密码
	   	 crt.Screen.WaitForString "Username:"
       crt.Screen.Send params(1) & vbcr
       crt.Screen.WaitForString "Password:"
       crt.Screen.Send params(2) & vbcr

			 crt.Screen.Send "show ver | in uptime" & vbCr
			  crt.Screen.Send "show ver | in uptime" & vbCr
			 output = crt.Screen.ReadString ("show ver | in uptime")
			 f_w.WriteLine(output)
			 f_w.Close()
			crt.Session.Disconnect
	 	 
    loop
    crt.Screen.Synchronous = False
	'crt.Quit 	
End Sub






