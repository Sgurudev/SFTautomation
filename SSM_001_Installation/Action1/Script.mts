'*********************************************************************************************************************************************
'Name: 	001_SSM_Softpaq_Installation					
'Description: 
'Pre-Requisite:1. SSM software located Path 
'Created By: 	Jaffar Vali			
'Creation Date: 					
'Application Name:					
'Changed By         Date @ Time:    		Description:						
'*********************************************************************************************************************************************
on error resume next 
Dim testData
Set testData = DataReader()
''''''''''''''''''''Variable declaration'''''''''''''''''''''''
strlogFileName="C:\CMIT\SSM\Logs"&"\"&Environment.Value("TestName")&"_Log.txt"
Datatable.ImportSheet "C:\CMIT\SSM\TestData\SSM_Data.xlsx","SSM",1
StrSSMExePath= Datatable.Value("str_SSM_ExePath",1)
strFolderPath=Datatable.Value("str_FolderPath",1)
SSM_InstallFilePath=Datatable.Value("SSM_InstallFilePath",1)
SSM_CheckFiles=Datatable.Value("SSM_CheckFiles",1)

'SSM_InstallFilePath="C:\SSM_TestFolder\"
'strSSM_ExePath ="C:\Users\valij\Desktop\ssm-3.2.4.1\setup.exe"
'strCheckFiles="SSM Release Document.url;SSM Users Guide.url;ssm.cab;SSM.exe"
'SSM_InstallFilePath="C:\Program Files (x86)\Hewlett-Packard\System Software Manager/"
'environment.Value(

Call Generate_LogFile(strlogFileName)

Call Write_Log(strlogFileName,"---------------"&Environment.Value("ActionName")&"-------Script Starts--------")
Systemutil.Run "C:\CMIT\SSM\Applications\ssm-3.2.6.1\setup.exe"
''''''''''--------------


'InvokeApplication "C:\Users\hpadmin\Desktop\ssm-3.2.4.1\setup.exe"

If window("InstallShield Wizard").Exist(8) Then
	Call Write_Log(strlogFileName,"HP System software Manager- Install Sheild Wizard window is opened")
	Reporter.ReportEvent micPass,"HP System software Manager- Install Sheild Wizard window is opened", "HP System software Manager- Install Sheild Wizard window is opened"
	window("InstallShield Wizard").WinButton("Next").Click
	Wait 3
		If  window("InstallShield Wizard").WinRadioButton("Modify").Exist(10) Then
			Call Write_Log(strlogFileName,"InstallShield Wizard software is already Installed")
			Reporter.ReportEvent micPass,"InstallShield Wizard software is already Installed","InstallShield Wizard software is already Installed"
			window("InstallShield Wizard").WinRadioButton("Modify").Click
			window("InstallShield Wizard").WinButton("Next").Click
			Wait 2
			window("InstallShield Wizard").WinButton("Next").Click
			Else
			Reporter.ReportEvent micPass,"InstallShield Wizard software is not Installed","InstallShield Wizard software is already Installed"
			Call Write_Log(strlogFileName,"InstallShield Wizard software is not Installed")
		End If
		
		
		window("InstallShield Wizard").WinButton("Install").Click
		'Window("regexpwndtitle:=InstallShield Wizard").WinButton("regexpwndtitle:=&Finish").CaptureBitmap ""
		window("InstallShield Wizard").WinButton("Finish").Click
		Wait 3
			If window("InstallShield Wizard").Exist(3) Then
				Call Write_Log(strlogFileName,"HP System software Manager- Install Sheild Wizard Installation is completed")
				Reporter.ReportEvent micFail,"HP System software Manager- Install Sheild Wizard Installation is completed", "HP System software Manager- Install Sheild Wizard window is opened"
				Else
				Call Write_Log(strlogFileName,"HP System software Manager- Install Sheild Wizard Installation is completed")
				Reporter.ReportEvent micPass,"HP System software Manager- Install Sheild Wizard Installation is completed", "HP System software Manager- Install Sheild Wizard window is opened"
			End If
		Call FilesSearch("C:\Program Files (x86)\HP\System Software Manager\","SSM Release Document.url;SSM Users Guide.url;ssm.cab;SSM.exe",strlogFileName)
			
	Else
	Reporter.ReportEvent micFail,"HP System software Manager- Install Sheild Wizard window is not opened", "HP System software Manager- Install Sheild Wizard window is not opened"
End If



Call CheckSW_Existence("HP System Software Manager","Y",strlogFileName)

Call Write_Log(strlogFileName,"---------------"&Environment.Value("ActionName")&"-------Script Ends--------")

resultGenerator err.number , true ,testData

Function FilesSearch(strfilepath,strCheckFiles,strlogFilePath)


Set Fso=createobject("Scripting.FileSystemObject")

strFiles=Split(strCheckFiles,";")
For intTemp=0  To Ubound(strFiles)
	strFileName=strfilepath&strFiles(intTemp)
	If  Fso.FileExists(strFileName) Then
		Call Write_Log(strlogFilePath,"File is present in Path"&strfilepath& "FileName:"&strFiles(intTemp))
		Reporter.ReportEvent micPass,"File is present in Path"&Vbnewline&strfilepath&Vbnewline&  "FileName:"&strFiles(intTemp),"File is present in Path"&strfilepath& "FileName:"&strFiles(intTemp)
		Else
		Call Write_Log(strlogFilePath,"File is not  present in Path"&strfilepath& "FileName:"&strFiles(intTemp))
		Reporter.ReportEvent micFail ,"File is not  present in Path"&Vbnewline& strfilepath&Vbnewline& "FileName:"&strFiles(intTemp),"File is not present in Path"&strfilepath& "FileName:"&strFiles(intTemp)
	End If 
Next

Set Fso= Nothing
End Function

Function Generate_LogFile(strlogFileName)
	Set objFile = CreateObject("Scripting.FileSystemObject")
	If objFile.FileExists(strlogFileName) Then
		Reporter.ReportEvent micPass,"Log file already exist in the expected path","Log file already exist in the expected path"
		Else
		Set textFile = objFile.CreateTextFile(strlogFileName,2)
		textFile.WriteLine "Log File created"
		textFile.Close
	End If
	
	Set objFile = Nothing
End Function


 'Environment.Value("SystemTempDir")&"\"&Environment.Value("TestName")&".txt"

Function Write_Log(strFilePath,strData)
Set objlogfso = CreateObject( "Scripting.FileSystemObject" )
Set textFile = objlogfso.OpenTextFile( strFilePath,8)
strLine = Date&""& Time & "-----" & strData
textFile.WriteLine strLine
textFile.Close
End Function


