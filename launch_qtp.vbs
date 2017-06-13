testPath = WScript.Arguments(0)
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
DoesFolderExist = objFSO.FolderExists(testPath)
Set objFSO = Nothing

If DoesFolderExist Then
    Dim qtApp 'Declare the Application object variable
    Dim qtTest 'Declare a Test object variable
    Set qtApp = CreateObject("QuickTest.Application")
    qtApp.Launch 'Start QuickTest
    qtApp.Visible = True 
    qtApp.Open testPath, False
    Set qtTest = qtApp.Test
    qtTest.Run 'Run the test
    qtTest.Close 'Close the test
    qtApp.Quit
Else
msgbox "error"
    'Couldn't find the test folder. That's bad. Guess we'll have to report on how we couldn't find the test.
    'Insert reporting mechanism here.
End If