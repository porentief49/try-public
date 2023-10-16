'2023-10-16 GRR: add update from local repo in developer mode
'2023-10-15 GRR: add rev. check, update only if required
'2023-10-13 GRR: initial creation

Option Explicit

'Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'
'Private Const LOCAL_REPO_BASE_PATH As String = "C:\MyData\Sandboxes\vba-code-vault\"
Private Const LOCAL_REPO_BASE_PATH_FILE As String = "vba-code-vault.txt"
Private Const GITHUB_RAW_BASE_URL As String = "https://raw.githubusercontent.com/porentief49/vba-code-vault/main/" ' full path like: https://raw.githubusercontent.com/porentief49/vba-code-vault/main/Mitarbeiterauslagen/Main.bas

Private Enum eDoWeUpdate
    WhatDoIKnow = 0
    YeahGoForIt = 1
    NahWhatIHaveIsGood = 0
End Enum

Public Sub UpdateAll()
    Dim lComponent As VBComponent
    Dim lResult As String
    Dim lGitHubCode As String
    Dim lThisRevDate As String
    Dim lGitHubRevDate As String
    Dim lDoWeUpdate As eDoWeUpdate
    Dim lReadFromLocal As Boolean
    Dim lLocalPath As String
    lLocalPath = GetLocalRepoPath
    lReadFromLocal = (LenB(lLocalPath) > 0)
    lDoWeUpdate = WhatDoIKnow
    For Each lComponent In ThisWorkbook.VBProject.VBComponents
        If lComponent.Type <= 2 Then
'            If lComponent.Name <> "Loader" Then
                If lReadFromLocal Then
                    lResult = ReadFromLocal(lLocalPath & GetWorkbookName & "\" & GetFileName(lComponent), lGitHubCode)
                Else
                    lResult = ReadFromGitHub(GITHUB_RAW_BASE_URL & GetWorkbookName & "/" & GetFileName(lComponent), lGitHubCode)
                End If
                If LenB(lResult) = 0 Then
                    If LenB(lGitHubCode) > 0 Then
                        lThisRevDate = GetRevDate(lComponent.CodeModule.Lines(1, lComponent.CodeModule.CountOfLines))
                        lGitHubRevDate = GetRevDate(lGitHubCode)
                        If lGitHubRevDate <> lThisRevDate Then
                            If lDoWeUpdate = WhatDoIKnow Then lDoWeUpdate = IIf(MsgBox(IIf(lReadFromLocal, "local repo", "GitHub") & " - different version found - update?", vbYesNo, "Auto-Update") = vbYes, YeahGoForIt, NahWhatIHaveIsGood)
                            If lDoWeUpdate = YeahGoForIt Then
                                lResult = UpdateModule(lComponent, lGitHubCode)
                                Call LogMessage(lComponent, IIf(LenB(lResult) = 0, IIf(lReadFromLocal, "local repo", "GitHub") & " - successfully updated with rev. " & lGitHubRevDate, "update failed - " & lResult))
                            Else
                                Call LogMessage(lComponent, IIf(lReadFromLocal, "local repo", "GitHub") & " - newer version available (" & lGitHubRevDate & "), but update declined")
                            End If
                        Else
                            Call LogMessage(lComponent, IIf(lReadFromLocal, "local repo", "GitHub") & " - already up-to-date (rev. " & lGitHubRevDate & ")")
                        End If
                    Else
                        Call LogMessage(lComponent, IIf(lReadFromLocal, "local repo", "GitHub") & " - read worked, but code module is empty - not updated")
                    End If
                Else
                    Call LogMessage(lComponent, IIf(lReadFromLocal, "local repo", "GitHub") & " - read failed - " & lResult)
                End If
'            End If
        End If
    Next lComponent
End Sub

Private Sub ExportAll(aPath As String)
    Dim lComponent As VBComponent
    Dim lFso As New FileSystemObject
    Dim lStream As TextStream
    For Each lComponent In ThisWorkbook.VBProject.VBComponents
        If lComponent.Type < 2 Then
            Set lStream = lFso.CreateTextFile(aPath & GetWorkbookName & "\" & GetFileName(lComponent))
            Call lStream.Write(lComponent.CodeModule.Lines(1, lComponent.CodeModule.CountOfLines))
            Call lStream.Close
            Set lStream = Nothing
        End If
    Next lComponent
End Sub

Public Sub ExportIfLocalGitRepoPresent()
    Dim lFso As New FileSystemObject
    Dim lLocalPath As String
'    If lFso.FolderExists(LOCAL_REPO_BASE_PATH & "\.git") Then
    lLocalPath = GetLocalRepoPath
    If LenB(lLocalPath) > 0 Then
        Call ExportAll(lLocalPath)
        Debug.Print "local repo - all sheets exported"
    End If
End Sub

'Public Sub ExportIfShiftKeyPressed()
'    If IsShiftKeyPressed Then
'        Call ExportAll
'        Debug.Print "all sheets exported"
'    End If
'End Sub

Private Function GetLocalRepoPath()
    Dim lFso As New FileSystemObject
    Dim lFile As String
    Static lDone As Boolean
    Static lRepoPath As String
    If Not lDone Then
        lFile = Application.ActiveWorkbook.Path & "\" & LOCAL_REPO_BASE_PATH_FILE
        If lFso.FileExists(lFile) Then lRepoPath = Trim$(lFso.OpenTextFile(lFile).ReadAll)
        lDone = True
    End If
    GetLocalRepoPath = lRepoPath
End Function

Private Sub LogMessage(aComponent As VBComponent, aMessage As String)
    Dim lModuleClass As String
    lModuleClass = IIf(aComponent.Type = vbext_ct_StdModule, "Module", "Class")
    Debug.Print lModuleClass & " '" & aComponent.Name & "': " & aMessage
End Sub

Private Function GetFileName(aComponent As VBComponent) As String
    GetFileName = aComponent.Name & IIf(aComponent.Type = vbext_ct_StdModule, ".bas", ".cls")
End Function

Private Function GetWorkbookName() As String
    GetWorkbookName = Split(ActiveWorkbook.Name, ".")(0)
End Function

Private Function GetRevDate(aCodeAllLines As String) As String
    Dim i As Long
    Dim lLine As String
    Dim lLines() As String
    Dim lDateMaybe As String
    Dim lLatestDate As String
    Dim lDone
    i = 0
    lDone = False
    lLines = Split(Replace$(aCodeAllLines, vbCr, vbNullString), vbLf)
    Do
        lLine = Trim$(lLines(i))
        If Left$(lLine, 1) = "'" Then
            lDateMaybe = Split(Trim$(Mid$(lLine, 2, 999)), " ")(0)
            If Len(lDateMaybe) = 10 Then
                If LenB(ReplaceAny(lDateMaybe, "0123456789-", vbNullString)) = 0 Then
                    If lDateMaybe > lLatestDate Then lLatestDate = lDateMaybe
                Else
                    'error format!!!
                End If
            Else
                'error format!!!
            End If
        Else
            lDone = True
        End If
        i = i + 1
    Loop Until lDone Or i > UBound(lLines)
    GetRevDate = lLatestDate
End Function

Private Function ReplaceAny(aIn As String, aReplaceChars As String, aWith As String) As String
    Dim lResult As String
    Dim i As Long
    lResult = aIn
    For i = 1 To Len(aReplaceChars)
        lResult = Replace$(lResult, Mid$(aReplaceChars, i, 1), aWith)
    Next i
    ReplaceAny = lResult
End Function

Private Function ReadFromGitHub(aUrl As String, ByRef aCode As String) As String 'credit: https://chat.openai.com/share/d3dd39f3-abb9-4233-aa19-7c3cef294b50
    Dim xmlHttp As Object
    On Error GoTo hell
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    Call xmlHttp.Open("GET", aUrl, False)  ' Send a GET request to the Google Drive file
    Call xmlHttp.send
    If xmlHttp.Status = 200 Then ' Check if the request was successful
        aCode = xmlHttp.responseText ' Read the response text (contents of the file)
        ReadFromGitHub = vbNullString
    Else
        ReadFromGitHub = "HTTP Error: " & xmlHttp.Status & " - " & xmlHttp.statusText ' Handle errors (e.g., file not found)
    End If
    Set xmlHttp = Nothing ' Clean up
    Exit Function
hell:
    Set xmlHttp = Nothing ' Clean up
    ReadFromGitHub = "Error: " & Err.Description
End Function

Private Function ReadFromLocal(aFilePath As String, ByRef aCode As String) As String
    Dim lFso As New FileSystemObject
    On Error GoTo hell
    aCode = lFso.OpenTextFile(aFilePath).ReadAll
    ReadFromLocal = vbNullString
    Exit Function
hell:
    ReadFromLocal = "Error: " & Err.Description
End Function

Private Function UpdateModule(aComponent As VBComponent, aCode As String) As String
    On Error GoTo hell
    With aComponent.CodeModule
        Call .DeleteLines(1, .CountOfLines)
        Call .InsertLines(1, aCode)
    End With
    UpdateModule = vbNullString
    Exit Function
hell:
    UpdateModule = "Error: " & Err.Description
End Function

