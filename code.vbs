Dim local_ctp_path
Dim destination_ctp_path 
Dim source_ctp_path
Dim start_up_path
Dim autoscript_path
Dim launch_QTP_path
Dim Test_file_Path
Dim username
Dim Password
Dim objFile
Dim Foldername
Dim Steps,step

read_enviroment_xml

set Steps  = Read_From_Workflow_xml()

for each Step in Steps
	On error resume next
	Select Case Step.text
	Case "Copy deliverables(Zip) from Server"
		Call CopyDeliverables()
		Foldername = Split(objFile.Path, ".")(0)
	Case "UnZip the Commpressed File"
		Call UnzipFiles(objFile.Path)
	Case "create bat file"
		Call create_mp_bat(Foldername&"\")
	Case "replace start up property file"
		Call replace_start_up(Foldername&"\")
	Case "Send Deliverables to remote machines and Execute QTP test"
		Call send_remote_system(Foldername)
	end select
	if Err.number <>0 then
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set file = fso.CreateTextFile("C:\CTP\Workflow.log")
		file.WriteLine Err.number & " : " & Err.Description
		file.Close
		exit for 
	end if
next




'Clean up Function 
Sub zCleanUp(file, count)   
        'Clean up
        Dim i, fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        For i = 1 To count
           If fso.FolderExists(fso.GetSpecialFolder(2) & "\Temporary Directory " & i & " for " & file) = True Then
           text = fso.DeleteFolder(fso.GetSpecialFolder(2) & "\Temporary Directory " & i & " for " & file, True)
           Else
              Exit For
           End If
        Next
    
End Sub


'Unzip files
Sub UnzipFiles(strFile)
    Dim arrFile
    Dim file_name
    Set fso = CreateObject("Scripting.FileSystemObject")
    arrFile = Split(strFile, ".")
    file_name =Split(strFile,"\")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(arrFile(0)&"\") then 
	  fso.DeleteFolder(arrFile(0))
	  WScript.Echo "deleted"
	End If
	
	
	fso.CreateFolder(arrFile(0)&"\")

    Dim sa, filesInzip, zfile, fso, i : i = 1
    Set sa = CreateObject("Shell.Application")
    Set filesInzip=sa.NameSpace(strFile).items
    
    For Each zfile In filesInzip
        If Not fso.FileExists(arrFile(0)& "\" & zfile) Then
             sa.NameSpace(arrFile(0)&"\").CopyHere(zfile), &H100 
             i = i + 1
        End If
           
        If i = 99 Then
            zCleanUp file_name(UBound(file_name)), i
            i = 1
        End If
    Next
        
    'create_mp_bat(arrFile(0)&"\")
    'replace_start_up(arrFile(0)&"\")
    'send_remote_system(arrFile(0))
        
    If i > 1 Then 
        zCleanUp file_name(UBound(file_name)), i
    End If
End Sub
    
    
'creating bat file 
Sub create_mp_bat(folder)
	dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
   		
    Set file = fso.CreateTextFile(local_ctp_path & "\mp.bat")
    file.WriteLine "cd """ & folder & """"
    file.WriteLine "java -jar ctp.jar"
    file.Close
End Sub


'replacing start up
Sub replace_start_up(folder)
	dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If(fso.FileExists (folder &"Startup.properties")) Then
   		fso.DeleteFile(folder &"Startup.properties")
  	End If
  	WScript.Echo "startup deleted"
    fso.CopyFile start_up_path,folder
   
  	WScript.Echo "new startup copied"
End Sub


'Copy  build folder to remote system and then trigger QTP in all the machines
Sub send_remote_system(folder)
    Dim array_file
    Dim folder_name
    Dim machine_name
    Dim upgrade_version,Lock,Compare_slk
    
    username=InputBox("Enter username (eg Cisco\aysingh): ")
    Password=InputBox("Enter your password : ")
    
    array_file = Split(folder,"\")
    folder_name = array_file(UBound(array_file))
    Set fso = CreateObject("Scripting.FileSystemObject")
    
   	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.load("C:\CTP\Machine.xml")

	Set Root = objXMLDoc.documentElement 
	Set NodeList1 = Root.getElementsByTagName("Variable/machine") 
	Set NodeList2 = Root.getElementsByTagName("Variable/UpgradeVersion")
	Set NodeList3 = Root.getElementsByTagName("Variable/Lock")
	Set NodeList4 = Root.getElementsByTagName("Variable/CompareSlk")

	Dim a
	For i = 0 To NodeList1.length-1
	
	machine_name=  NodeList1(i).text
	upgrade_version= NodeList2(i).text
	Lock = NodeList3(i).text
	Compare_slk = NodeList4(i).text
    		
    create_CTP_Env upgrade_version , Lock , Compare_slk ,machine_name
    Create_features_to_test machine_name
    
    If fso.FileExists ("\\" & machine_name & "\"& destination_ctp_path & "\mp.bat") Then
    	fso.DeleteFile "\\" & machine_name & "\"& destination_ctp_path & "\mp.bat"
    End if
    	
    If fso.FolderExists ("\\" & machine_name & "\" &destination_ctp_path&"\"&folder_name) Then 
    	WScript.Echo "folder deleted"
    	fso.DeleteFolder ("\\" & machine_name & "\" &destination_ctp_path&"\"&folder_name)
    End If
    	
    If fso.FileExists ("\\" & machine_name & "\" & destination_ctp_path &"\" & folder_name & ".zip") Then
    	WScript.Echo "zip deleted"
    	fso.DeleteFile ("\\" & machine_name & "\" & destination_ctp_path &"\" & folder_name & ".zip")
    End if
	
    fso.CopyFolder local_ctp_path,"\\" & machine_name & "\"& destination_ctp_path
    WScript.Echo "copied"
    		
    trigerring_QTP(machine_name)
    
    Next
    
End Sub

'create enviroment file eniviroment.ini
Sub create_CTP_Env( upgrade_version , Lock ,Compare_slk ,machine_name)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists ("\\" & machine_name & "\" & autoscript_path & "\CTPEnv.ini") then
	    fso.DeleteFile ("\\" & machine_name & "\" & autoscript_path & "\CTPEnv.ini")
	End If
	WScript.Echo "CTP_env deleted"
	Set file = fso.CreateTextFile ("\\" & machine_name & "\" & autoscript_path & "\CTPEnv.ini")
	file.WriteLine "[Environment]"
	file.WriteLine "CTPUpgradeVersion = " & upgrade_version
	file.WriteLine "Lock = " & Lock
	file.WriteLine "CompareSLK = " & Compare_slk	
	file.Close
	
	WScript.Echo "CTP_env_created"
End Sub

'creating feature to Test file
Sub Create_features_to_test( machine_name)
   Set fso = CreateObject("Scripting.FileSystemObject")
   If fso.FileExists ("\\" & machine_name & "\" & autoscript_path & "\FeaturesToTest.txt") then
	    fso.DeleteFile ("\\" & machine_name & "\" & autoscript_path & "\FeaturesToTest.txt")
   End If
 
   fso.CopyFile "\\" & machine_name & "\" & autoscript_path & "\FeaturesToTest_Backup.txt" , "\\" & machine_name & "\" & autoscript_path & "\FeaturesToTest.txt" 
   
End Sub


'running psexec command
Sub trigerring_QTP(machine_name)
    Set wshShell = WScript.CreateObject ("WSCript.shell")
    wshShell.Run "psexec \\" & machine_name & " -u " & username &" -p "& Password & " -i 1 cscript """  & launch_QTP_path & "\launch_qtp.vbs"" " & Test_file_Path
End Sub


'Reading all the variable from xml file
Sub read_enviroment_xml()
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.load("C:\CTP\Enviroment.xml")

	Set Root = objXMLDoc.documentElement 
	Set NodeList = Root.getElementsByTagName("Variable") 
	
	source_ctp_path = NodeList(0).Text
	destination_ctp_path = NodeList(1).Text
	start_up_path = NodeList(2).Text
	local_ctp_path = NodeList(3).Text
	autoscript_path = NodeList(4).Text
	launch_QTP_path = NodeList(5).Text
	Test_file_Path = NodeList(6).Text
	

End Sub

function Read_From_Workflow_xml()
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.load("C:\CTP\WorkFlow.xml")
	Set Root = objXMLDoc.documentElement 
	set Read_From_Workflow_xml = Root.getElementsByTagName("FunctionName")	
end function


Sub CopyDeliverables()
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(local_ctp_path) Then
objFSO.DeleteFolder (local_ctp_path)
End If

objFSO.CreateFolder(local_ctp_path )

objFso.CopyFile source_ctp_path , local_ctp_path &"\" 

Set objFolder = objFSO.GetFolder(local_ctp_path)

Set colFiles = objFolder.Files

For Each objFile in colFiles
	If StrComp("zip",objFSO.GetExtensionName(ObjFile.Path))=0 Then
    Exit For
    End if 
Next
end sub


​'http://www.automationrepository.com/tutorials-for-qtp-beginners/
'http://www.qtpworld.com/index.php?cid=44

'##########################################################################################################################################################################
'Capture Error
Function CaptureImageonError()
	Dim moment
	Dim sMyfile
	moment=Now() 'displays the current date and time
	sMyfile= moment&".png"
	sMyfile = Replace(sMyfile,"/","-") 'removing special char / & :
	sMyfile = Replace(sMyfile,":","-")
	sMyfile= "C:\"&sMyfile 'Path where file would be saved
	REM Desktop.CaptureBitmap sMyfile,true
	Msgbox sMyfile
End Function

'##########################################################################################################################################################################
'Minimizes QTP window:
Sub MinimizeQTPWindow ()
    Set     qtApp = getObject("","QuickTest.Application")
    qtApp.WindowState = "Minimized"
    Set qtApp = Nothing
End Sub

'##########################################################################################################################################################################
' Associate library from script
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'1 - Using AOM (Automation Object Model)
' QTP AOM is a mechanism using which you can control various QTP operations from outside QTP. Using QTP Automation Object Model, you can write a code 
' which would open a QTP test and associate a function library to that test.
' Example: Using the below code, you can open QTP, then open any test case and associate a required function library to that test case.
' To do so, copy paste the below code in a notepad and save it with a .vbs extension.

'Open QTP
	Set objQTP = CreateObject("QuickTest.Application")
	objQTP.Launch
	objQTP.Visible = True
'Open a test and associate a function library to the test
	objQTP.Open "C:\Automation\SampleTest", False, False
	Set objLib = objQTP.Test.Settings.Resources.Libraries
'If the library is not already associated with the test case, associate it..
	If objLib.Find("C:\SampleFunctionLibrary.vbs") = -1 Then ' If library is not already added
		objLib.Add "C:\SampleFunctionLibrary.vbs", 1 ' Assoc​​iate the library to the test case
	End
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Using ExecuteFile Method
'ExecuteFile statement executes all the VBScript statements in a specified file. 
'After the file has been executed, all the functions, subroutines and other elements from the file (function library) are available 
'to the action as global entities. Simply put, once the file is executed, its functions can be used by the action. Y
'ou can use the below mentioned logic to use ExecuteFile method to associate function libraries to your script.
'Action begins
ExecuteFile "C:\YourFunctionLibrary.vbs"
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Using LoadFunctionLibrary Method
' LoadFunctionLibrary, a new method introduced in QTP 11 allows you to load a function library when a step runs. 
' You can load multiple function libraries from a single line by using a comma delimiter.
LoadFunctionLibrary "C:\YourFunctionLibrary_1.vbs" 'Associate a single function library
LoadFunctionLibrary "C:\FuncLib_1.vbs", "C:\FuncLib_2.vbs" 'Associate more than 1 function libraries
 
'##########################################################################################################################################################################
'Speak from QTP
Set oVoice = CreateObject("sapi.spvoice")
oVoice.Speak "Agent1 was able to login successfully"
Set oVoice = Nothing
'##########################################################################################################################################################################
' DataTable
	Dim record_count
	record_count = DataTable.GetSheet("Action1").GetRowCount
	For i = 1 to record_count
		MsgBox DataTable("Column1Name",dtlocalsheet)
		DataTable.GetSheet("Action1").SetNextRow 
	Next
'##########################################################################################################################################################################
' reading data from excel sheet in QTP
' 1st method:
	datatable.importsheet "path of the excel file.xls",source sheetID,desination sheetID
	n = datatable.getsheet("desination sheetname").getrowcount
	for i = 1 to n
		columnname = datatable.getsheet("destination sheetname").getparameter(i).name
		if colunmname = knowncolumnname then
			value = datatable.getsheet(destinationsheetname).getparameter(i)
		end if
	next
' DataTable("Name",dtlocalsheet) = datatable.GetSheet("Action1").GetParameter("Name").Value
' To get Column Name for 1st Column : datatable.GetSheet("Action1").GetParameter(1).Name

' 2nd method:
Dim appExcel, objWorkBook, objSheet, columncount, rowcount
Set appExcel = CreateObject("Excel.Application")
Set objWorkBook = appExcel.Workbooks.open("C:\Test1.xls")
Set objSheet = appExcel.Sheets("MainSheet")
columncount = objSheet.usedrange.columns.count
rowcount = objSheet.usedrange.rows.count
Find_Details="345"
For a= 1 to rowcount
	For b = 1 to columncount
		fieldvalue =objSheet.cells(a,b)
		If  cstr(fieldvalue)= Cstr(Find_Details) Then
			msgbox "Found"
		Exit For
		End If
	Next
Next


'##########################################################################################################################################################################
'##########################################################################################################################################################################
'##########################################################################################################################################################################
'##########################################################################################################################################################################
'##########################################################################################################################################################################
'##########################################################################################################################################################################

