Option Explicit

'-------------------------------------------------------------------------------------------------------------------
' check out raw.githubusercontent.com
' https://stackoverflow.com/questions/39065921/what-do-raw-githubusercontent-com-urls-represent
' https://stackoverflow.com/questions/2466735/how-to-sparsely-checkout-only-one-single-file-from-a-git-repository
' for instance already works: https://raw.githubusercontent.com/porentief49/try-public/main/README.md
'-------------------------------------------------------------------------------------------------------------------


Private Const LOCAL_REPO_BASE_PATH As String = "C:\MyData\Sandboxes\vba-code-vault\"
'Private Const GITHUB_RAW_BASE_URL As String = "https://raw.githubusercontent.com/porentief49/vba-code-vault/main/Mitarbeiterauslagen/Main.bas"
Private Const GITHUB_RAW_BASE_URL As String = "https://raw.githubusercontent.com/porentief49/vba-code-vault/main/"

'Public Sub UpdateCode()
'    'https://drive.google.com/file/d/1J8bdRrYpTWF-G1KwxkrqQykNETR8BJ1a/view?usp=sharing
'    Const MODULE_NAME As String = "Main"
''    Const CODE_FILE As String = "c:\temp\Mitarbeiterauslagen.txt"
''    Debug.Print ImportModuleFromFile(CODE_FILE)
'
'    Dim lCode As String
'    Dim lResult As String
''    lResult = ReadGoogleDrive("1J8bdRrYpTWF-G1KwxkrqQykNETR8BJ1a", lCode)
'    lResult = ReadGitHubRaw("https://raw.githubusercontent.com/porentief49/vba-code-vault/main/Mitarbeiterauslagen/Main.bas", lCode)
'    If LenB(lResult) = 0 Then
'        If LenB(lCode) > 0 Then
'            lResult = UpdateModule(MODULE_NAME, lCode)
'            If LenB(lResult) = 0 Then
'                Debug.Print "Module successfully updated"
'            Else
'                Debug.Print "UpdateModule did not work: " & lResult
'            End If
'        Else
'            Debug.Print "ReadGoogleDrive worked, but no code in module"
'        End If
'    Else
'        Debug.Print "ReadGoogleDrive did not work: " & lResult
'    End If
'End Sub

Public Sub ExportAll()
    Dim lComponent As VBComponent
    Dim lFso As New FileSystemObject
    Dim lStream As TextStream
    For Each lComponent In ThisWorkbook.VBProject.VBComponents
        If lComponent.Type < 2 Then
            Set lStream = lFso.CreateTextFile(LOCAL_REPO_BASE_PATH & GetWorkbookName & "\" & GetFileName(lComponent))
            Call lStream.WriteLine(lComponent.CodeModule.Lines(1, lComponent.CodeModule.CountOfLines))
            Call lStream.Close
            Set lStream = Nothing
        End If
    Next lComponent
End Sub

Public Sub UpdateAll()
    Dim lComponent As VBComponent
    Dim lResult As String
    Dim lCode As String
    For Each lComponent In ThisWorkbook.VBProject.VBComponents
        If lComponent.Name <> "Loader" Then
            lResult = ReadGitHubRaw(GITHUB_RAW_BASE_URL & GetWorkbookName & "/" & GetFileName(lComponent), lCode)
            If LenB(lResult) = 0 Then
                If LenB(lCode) > 0 Then
                    lResult = UpdateModule(lComponent.Name, lCode)
                    If LenB(lResult) = 0 Then
                        Debug.Print "Module successfully updated"
                    Else
                        Debug.Print "UpdateModule did not work: " & lResult
                    End If
                Else
                    Debug.Print "ReadGoogleDrive worked, but no code in module"
                End If
            Else
                Debug.Print "ReadGoogleDrive did not work: " & lResult
            End If
        End If
    Next lComponent
End Sub

Private Function GetFileName(aComponent As VBComponent) As String
    GetFileName = aComponent.Name & IIf(aComponent.Type = vbext_ct_StdModule, ".bas", ".cls")
End Function

Private Function GetWorkbookName() As String
    GetWorkbookName = Split(ActiveWorkbook.Name, ".")(0)
End Function



'Sub ListModulesAndClasses()
'    Dim VBComp As Object
'    Dim VBProj As Object
'    Dim VBCompType As Long
'    Dim ModuleName As String
'
'    ' Set a reference to the VBA project in which you want to list modules and classes.
'    Set VBProj = ThisWorkbook.VBProject ' Change "ThisWorkbook" to the appropriate workbook or other VBA project.
'
'    ' Loop through all components in the project.
'    For Each VBComp In VBProj.VBComponents
'        VBCompType = VBComp.Type
'        ModuleName = VBComp.Name
'
'        ' Check if the component is a module or a class and print its name.
'        If VBCompType = 1 Then
'            Debug.Print "Module: " & ModuleName
'        ElseIf VBCompType = 2 Then
'            Debug.Print "Class: " & ModuleName
'        Else
'            Debug.Print "Type " & CStr(VBCompType) & ": " & ModuleName
'        End If
'    Next VBComp
'End Sub


'Private Sub ReadGoogleDriveOLD() '2023-10-09, rAiner Gruber
'' credit: https://chat.openai.com/share/d3dd39f3-abb9-4233-aa19-7c3cef294b50
'' link: https://drive.google.com/file/d/18D2GscIRnO286zlWqNTSL06jcMgtTeon/view?usp=sharing
'    ' Define your Google Drive file URL
'    Dim fileURL As String
'' this is the link format needed: https://drive.google.com/uc?id=YOUR_FILE_ID"
'' when you share in google drive, you will get this link: https://drive.google.com/file/d/18D2GscIRnO286zlWqNTSL06jcMgtTeon/view?usp=sharing
'' now grab the doc ID and put it into the link above: https://drive.google.com/uc?id=18D2GscIRnO286zlWqNTSL06jcMgtTeon
''
''    fileURL = "https://drive.google.com/file/d/18D2GscIRnO286zlWqNTSL06jcMgtTeon/view?usp=sharing"
''    fileURL = "https://docs.google.com/document/d/e/2PACX-1vQuF1Kaw3WoehwpHqFFJAc46A51gbDUhoNqes4kHYcB86wacRTJxgBoplC3lrhK0ugTxNSWiwIfPZ9N/pub"
'
''    fileURL = "https://drive.google.com/uc?id=18D2GscIRnO286zlWqNTSL06jcMgtTeon"
'    fileURL = "https://drive.google.com/uc?id=1J8bdRrYpTWF-G1KwxkrqQykNETR8BJ1a"
'
'
'    'https://drive.google.com/file/d/1J8bdRrYpTWF-G1KwxkrqQykNETR8BJ1a/view?usp=sharing
'
''    https://drive.google.com/file/d/1Sr8H-HY1yFLhmcXSh6_HDD2MTnyoeVpG/view?usp=sharing
'
'    ' Create a HTTP request
'    Dim xmlHttp As Object
'    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
'
'    ' Send a GET request to the Google Drive file
'    xmlHttp.Open "GET", fileURL, False
'    xmlHttp.send
'
'    ' Check if the request was successful
'    If xmlHttp.Status = 200 Then
'        ' Read the response text (contents of the file)
'        Dim fileContents As String
'        fileContents = xmlHttp.responseText
'
'        ' Display the file contents (you can modify this part as needed)
'        MsgBox fileContents
'        Open "c:\temp\download.txt" For Output As #1
'        Print #1, fileContents
'        Close #1
'    Else
'        ' Handle errors (e.g., file not found)
'        MsgBox "Error: " & xmlHttp.Status & " - " & xmlHttp.statusText
'    End If
'
'    ' Clean up
'    Set xmlHttp = Nothing
'
'End Sub

'Private Function ReadGoogleDrive(aFileId As String, ByRef aCodeModule As String) As String '2023-10-09, rAiner Gruber
'    ' credit: https://chat.openai.com/share/d3dd39f3-abb9-4233-aa19-7c3cef294b50
'    ' this is the link format needed: https://drive.google.com/uc?id=YOUR_FILE_ID"
'    ' when you share in google drive, you will get this link: https://drive.google.com/file/d/18D2GscIRnO286zlWqNTSL06jcMgtTeon/view?usp=sharing
'    ' now grab the doc ID and put it into the link above: https://drive.google.com/uc?id=18D2GscIRnO286zlWqNTSL06jcMgtTeon
'    Dim fileURL As String
'    Dim xmlHttp As Object
'    On Error GoTo hell
'    fileURL = "https://drive.google.com/uc?id=" & aFileId
'    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
'    Call xmlHttp.Open("GET", fileURL, False)  ' Send a GET request to the Google Drive file
'    Call xmlHttp.send
'    If xmlHttp.Status = 200 Then ' Check if the request was successful
'        aCodeModule = xmlHttp.responseText ' Read the response text (contents of the file)
''        MsgBox fileContents ' Display the file contents (you can modify this part as needed)
''        Open "c:\temp\download.txt" For Output As #1
''        Print #1, fileContents
''        Close #1
'        ReadGoogleDrive = vbNullString
'    Else
'        ' Handle errors (e.g., file not found)
'        ReadGoogleDrive = "HTTP Error: " & xmlHttp.Status & " - " & xmlHttp.statusText
'    End If
'    Set xmlHttp = Nothing ' Clean up
'    Exit Function
'hell:
'    Set xmlHttp = Nothing ' Clean up
'    ReadGoogleDrive = "Error: " & Err.Description
'End Function

Private Function ReadGitHubRaw(aUrl As String, ByRef aCodeModule As String) As String '2023-10-09, rAiner Gruber
    ' credit: https://chat.openai.com/share/d3dd39f3-abb9-4233-aa19-7c3cef294b50
    ' this is the link format needed: https://drive.google.com/uc?id=YOUR_FILE_ID"
    ' when you share in google drive, you will get this link: https://drive.google.com/file/d/18D2GscIRnO286zlWqNTSL06jcMgtTeon/view?usp=sharing
    ' now grab the doc ID and put it into the link above: https://drive.google.com/uc?id=18D2GscIRnO286zlWqNTSL06jcMgtTeon
'    Dim fileURL As String
    Dim xmlHttp As Object
    On Error GoTo hell
'    fileURL = "https://drive.google.com/uc?id=" & aFileId
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    Call xmlHttp.Open("GET", aUrl, False)  ' Send a GET request to the Google Drive file
    Call xmlHttp.send
    If xmlHttp.Status = 200 Then ' Check if the request was successful
        aCodeModule = xmlHttp.responseText ' Read the response text (contents of the file)
'        MsgBox fileContents ' Display the file contents (you can modify this part as needed)
'        Open "c:\temp\download.txt" For Output As #1
'        Print #1, fileContents
'        Close #1
        ReadGitHubRaw = vbNullString
    Else
        ' Handle errors (e.g., file not found)
        ReadGitHubRaw = "HTTP Error: " & xmlHttp.Status & " - " & xmlHttp.statusText
    End If
    Set xmlHttp = Nothing ' Clean up
    Exit Function
hell:
    Set xmlHttp = Nothing ' Clean up
    ReadGitHubRaw = "Error: " & Err.Description
End Function


'Private Sub ExportModule(aComponent As VBComponent, aFilePath As String) '2023-10-09, rAiner Gruber
'    Dim lFso As New FileSystemObject
'    Dim lStream As TextStream
''    Dim lProject As VBProject
''    Dim lComponent As VBComponent
'    Dim lModule As CodeModule
''    Dim lFile As String
''    Dim lExtension As String
''    Set lProject = ThisWorkbook.VBProject
''    Set lComponent = lProject.VBComponents(aModule)
'    Set lModule = aComponent.CodeModule
''    If aComponent.Type = 1 Then lExtension = ".bas"
''    else if aComponent
''    lFile = "C:\temp\" & aModule & ".txt"
''    lFile = "G:\My Drive\CodeBase\" & aModule & ".txt"
'    Set lStream = lFso.CreateTextFile(aFilePath)
'    Call lStream.WriteLine(lModule.Lines(1, lModule.CountOfLines))
'    Call lStream.Close
'    Set lStream = Nothing
''    Debug.Print "VBA module exported to " & lFile
'End Sub

'Private Function ImportModuleFromFile(aFile As String) As String '2023-10-09, rAiner Gruber
'    Dim lFso As New FileSystemObject
'    Dim lStream As TextStream
'    Dim lProject As VBProject
'    Dim lComponent As VBComponent
'    Dim lModule As CodeModule
''    Dim lFile As String
'    Dim lLines As String
'    Dim lModuleName As String
'    On Error GoTo hell
'    lModuleName = Split(lFso.GetFileName(aFile), ".")(0)
'    Set lProject = ThisWorkbook.VBProject
'    Set lComponent = lProject.VBComponents(lModuleName)
'    Set lModule = lComponent.CodeModule
''    lFile = "C:\temp\" & lModuleName & ".txt"
'    Set lStream = lFso.OpenTextFile(aFile)
'    lLines = lStream.ReadAll
'    Call lStream.Close
'    Set lStream = Nothing
'    Call lModule.DeleteLines(1, lModule.CountOfLines)
'    Call lModule.InsertLines(1, lLines)
'    ImportModuleFromFile = "VBA module imported from '" & aFile & "'"
'    Exit Function
'hell:
'    Set lStream = Nothing
'    ImportModuleFromFile = "Error: " & Err.Description
'End Function


Private Function UpdateModule(aModuleName As String, aCode As String) As String '2023-10-09, rAiner Gruber
'    Dim lFso As New FileSystemObject
'    Dim lStream As TextStream
    Dim lProject As VBProject
    Dim lComponent As VBComponent
    Dim lModule As CodeModule
'    Dim lFile As String
'    Dim lLines As String
'    Dim lModuleName As String
    On Error GoTo hell
'    lModuleName = Split(lFso.GetFileName(aFile), ".")(0)
    Set lProject = ThisWorkbook.VBProject
    Set lComponent = lProject.VBComponents(aModuleName)
    Set lModule = lComponent.CodeModule
'    lFile = "C:\temp\" & lModuleName & ".txt"
'    Set lStream = lFso.OpenTextFile(aFile)
'    lLines = lStream.ReadAll
'    Call lStream.Close
'    Set lStream = Nothing
    Call lModule.DeleteLines(1, lModule.CountOfLines)
    Call lModule.InsertLines(1, aCode)
    UpdateModule = vbNullString
    Exit Function
hell:
'    Set lStream = Nothing
    UpdateModule = "Error: " & Err.Description
End Function

