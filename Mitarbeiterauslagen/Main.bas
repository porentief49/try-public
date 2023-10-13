' 2023-10-10 little code history ...

Option Explicit

Public Const EXP_ROW_TIMERANGE As Long = 1
Public Const EXP_ROW_BALANCE_MONTH As Long = 3
Public Const EXP_ROW_BALANCEL_TOTAL As Long = 4
Public Const EXP_ROW_RANGE_FROM As Long = 3
Public Const EXP_ROW_RANGE_TO As Long = 4
Public Const EXP_ROW_STATUS As Long = 6
Public Const EXP_ROW_DATA_FIRST As Long = 9
Public Const EXP_ROW_DATA_LAST As Long = 9999
Public Const EXP_COL_TIMERANGE As Long = 4
Public Const COL_DATE As Long = 1
Public Const COL_STATUS As Long = 1
Public Const COL_EXPENSE As Long = 2
Public Const COL_VENDOR As Long = 3
Public Const COL_AMOUNT As Long = 4
Public Const COL_COMMENT As Long = 5
Public Const SHEET_EXPENSES As String = "Auslagen"
Public Const SHEET_BALANCE As String = "Abrechnung"
Public Const BAL_ROW_HEADLINE As Long = 1
Public Const BAL_ROW_RECEIVER As Long = 3
Public Const BAL_ROW_IBAN As Long = 6
Public Const BAL_ROW_AMOUNT As Long = 7
Public Const COL_HEADLINE As Long = 1
Public Const BAL_COL_DATA As Long = 2

Public Function FindFreeRow() As Long '2020-04-27 rAiner Gruber
    Const EMPTY_ROW_THRESHOLD As Long = 100
    Dim lRow As Long
    Dim lEmptyCount As Long
    lRow = EXP_ROW_DATA_FIRST
    lEmptyCount = 0
    Do While (lEmptyCount < EMPTY_ROW_THRESHOLD) And (lRow < EXP_ROW_DATA_LAST) '@@@ this is not perfectly correct, because close to the end, it will start overwriting rows
        lEmptyCount = IifLong((Len(Trim$(Cells(lRow, COL_DATE).Value2) & Trim$(Cells(lRow, COL_EXPENSE).Value2) & Trim$(Cells(lRow, COL_VENDOR).Value2) & Trim$(Cells(lRow, COL_AMOUNT).Value2) & Trim$(Cells(lRow, COL_COMMENT).Value2)) > 0), 0, lEmptyCount + 1)
        lRow = lRow + 1
    Loop
    FindFreeRow = lRow - EMPTY_ROW_THRESHOLD
End Function

Public Sub UpdateStatus(aStatus As String, aOk As Boolean) '2023-10-09, rAiner Gruber
    Dim lColLetterFrom As String
    Dim lColLetterTo As String
    With Cells(EXP_ROW_STATUS, COL_STATUS)
        .Value2 = aStatus
        .Font.Color = IifLong(aOk, &HAA00&, &HCC&)
    End With
    lColLetterFrom = ConvertColToLetter(COL_STATUS)
    lColLetterTo = ConvertColToLetter(COL_COMMENT)
    Range(lColLetterFrom & CStr(EXP_ROW_STATUS) & ":" & lColLetterTo & CStr(EXP_ROW_STATUS)).Interior.Color = IifLong(aOk, &HDDFFDD, &HDDDDFF)
End Sub

Public Sub CreateBalance() '2023-10-08 rAiner Gruber
    Const BALANCE_TITLE As String = "Abrechnung Mitarbeiterauslagen "
    Dim lBalance As Double
    Dim lTimeRange As String
    Dim lRow As Long
    Call SortEntries
    lBalance = Cells(EXP_ROW_BALANCE_MONTH, COL_AMOUNT).Value2
    lTimeRange = Cells(EXP_ROW_TIMERANGE, EXP_COL_TIMERANGE).Value2
    If Abs(lBalance) > 0.004 Then
        lRow = FindFreeRow
        Cells(lRow, COL_DATE).Value2 = DateFromTo(lTimeRange, False)
        Cells(lRow, COL_EXPENSE).Value2 = BALANCE_TITLE & lTimeRange
        Cells(lRow, COL_AMOUNT).Value2 = -lBalance
'        Call UpdateStatus("OK - Abrechnung fuer " & lTimeRange & " erstellt.", True)
        Call UpdateStatus("OK - " & lTimeRange & " cleared", True)
    Else
'        Call MsgBox("Abrechnung f? lTimeRange & " nicht m?ch - Saldo ist bereits 0.00?. Wurde vielleicht bereits abgerechnet?", vbExclamation, "Abrechnung")
'        Call UpdateStatus("Abrechnung fuer " & lTimeRange & " nicht moeglich - Saldo ist bereits 0.00 EUR. Wurde vielleicht bereits abgerechnet?", False)
        Call UpdateStatus("Clearing " & lTimeRange & " not possible - balance is already 0.00EUR", False)
    End If
End Sub

Public Sub ExportPDF() '2015-05-11 rAiner Gruber
    Dim lFileName As String
    Dim lCurrentSheet As Object
    Dim lBalance As Double
    Dim lTimeRange As String
    On Error GoTo hell
    lBalance = Cells(EXP_ROW_BALANCE_MONTH, COL_AMOUNT).Value2
    lTimeRange = Cells(EXP_ROW_TIMERANGE, EXP_COL_TIMERANGE).Value2
    If Abs(lBalance) < 0.004 Then
        Application.ScreenUpdating = False
        Set lCurrentSheet = ActiveSheet
        Call Sheets(SHEET_BALANCE).Select
        lFileName = Replace$(Environ("userprofile") & "\Desktop\" & Cells(1, 1).Value, " ", "_") & ".pdf"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=lFileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Call lCurrentSheet.Select
        Application.ScreenUpdating = True
'        Call UpdateStatus("OK - Abrechnung fuer " & lTimeRange & " exportiert.", True)
        Call UpdateStatus("OK - report created for " & lTimeRange, True)
    Else
'        Call MsgBox("Export Abrechnung " & lTimeRange & " nicht m?ch - Saldo ungleich 0.00?. Bitte erst abrechnen.", vbExclamation, "Abrechnung")
'        Call UpdateStatus("Export Abrechnung " & lTimeRange & " nicht m?ch - Saldo ungleich 0.00?. Bitte erst abrechnen.", False)
        Call UpdateStatus("Report for " & lTimeRange & " not created - balance != 0.00EUR - please clear first", False)
    End If
    Exit Sub
hell:
    Application.ScreenUpdating = True
    Err.Raise (Err.Number)
End Sub

Public Function EpcQrText() '2023-10-08 rAiner Gruber
    Const NEW_LINE As String = "\r\n"
    Dim l1_ServiceTag As String
    Dim l2_Version As String
    Dim l3_Encoding As String
    Dim l4_Id As String
    Dim l5_BIC As String
    Dim l6_Receiver As String
    Dim l7_IBAN As String
    Dim l8_Amount As String
    Dim l9_Code As String
    Dim l10_Ref As String
    Dim l11_Title As String
    Dim l12_Comment As String
    l1_ServiceTag = "BCD"
    l2_Version = "002"
    l3_Encoding = "1"
    l4_Id = "SCT"
    l5_BIC = "" 'optional
    l6_Receiver = Left$(Cells(BAL_ROW_RECEIVER, BAL_COL_DATA).Value2, 60)
    l7_IBAN = Replace$(Cells(BAL_ROW_IBAN, BAL_COL_DATA).Value2, " ", vbNullString)
    l8_Amount = "EUR" & Format$(Cells(BAL_ROW_AMOUNT, BAL_COL_DATA).Value2, "0.00")
    l9_Code = ""
    l10_Ref = ""
    l11_Title = Left$(Cells(BAL_ROW_HEADLINE, COL_HEADLINE).Value2, 140)
    l12_Comment = ""
    EpcQrText = l1_ServiceTag & NEW_LINE & l2_Version & NEW_LINE & l3_Encoding & NEW_LINE & l4_Id & NEW_LINE & l5_BIC & NEW_LINE & l6_Receiver & NEW_LINE & l7_IBAN & NEW_LINE & l8_Amount & NEW_LINE & l9_Code & NEW_LINE & l10_Ref & NEW_LINE & l11_Title & NEW_LINE & l12_Comment
End Function

Public Sub GenerateQRCode(inputString As String, outputPath As String) '2023-10-08 rAiner Gruber
    ' credit: https://chat.openai.com/share/4a3043e0-024f-499b-a270-3426e18e9f1a
    ' check out: https://goqr.me/api/
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Replace "Your-API-Endpoint" with the actual API endpoint you want to use
    Dim apiEndpoint As String
'    apiEndpoint = "https://api.qr-code-generator.com/v1/create/?data=" & inputString
    apiEndpoint = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&ecc=M&data=" & inputString
    
    ' Send GET request to the API
    xmlHttp.Open "GET", apiEndpoint, False
    xmlHttp.send ""
    
    ' Check if the request was successful
    If xmlHttp.Status = 200 Then
        ' Create a binary stream for the image
        Dim imageStream As Object
        Set imageStream = CreateObject("ADODB.Stream")
        imageStream.Open
        imageStream.Type = 1 ' Binary
        imageStream.Write xmlHttp.responseBody
        
        ' Save the image to the specified file path
        imageStream.SaveToFile outputPath, 2 ' Overwrite existing file
        
        ' Clean up
        imageStream.Close
        Set imageStream = Nothing
    Else
        MsgBox "Failed to generate QR code. HTTP Status: " & xmlHttp.Status
    End If
    
    ' Clean up
    xmlHttp.abort
    Set xmlHttp = Nothing
End Sub

Public Sub TestQr() '2023-10-08 rAiner Gruber
    ' Usage example:
    Dim inputString As String
    Dim outputPath As String
    inputString = "BCD|002|1|SCT||Paul|DE30702501500027525005|EUR1.00||Zweck"
    outputPath = "C:\temp\QRCode.png"
    Call GenerateQRCode(Replace$(inputString, "|", "\n"), outputPath)
    Debug.Print "QR code generated and saved as " & outputPath
End Sub

Public Sub SortEntries() '2020-04-29 rAiner Gruber
    Dim lSelectCache As Range
    Dim lColLetter As String
    lColLetter = ConvertColToLetter(COL_DATE)
    Application.ScreenUpdating = False
    Set lSelectCache = Selection
    Call Rows(CStr(EXP_ROW_DATA_FIRST) & ":" & CStr(EXP_ROW_DATA_LAST)).Select
    With ActiveWorkbook.ActiveSheet.Sort
        Call .SortFields.Clear
        Call .SortFields.Add2(Key:=Range(lColLetter & CStr(EXP_ROW_DATA_FIRST) & ":" & lColLetter & CStr(EXP_ROW_DATA_LAST)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal)
        Call .SetRange(Range(CStr(EXP_ROW_DATA_FIRST) & ":" & CStr(EXP_ROW_DATA_LAST)))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        Call .Apply
    End With
    Call lSelectCache.Select
    Application.ScreenUpdating = True
End Sub

Public Sub CopyExpenses() '2023-10-09, rAiner Gruber
    Const MAX_COUNT As Long = 53
    Dim lCount As Long
    Dim lDate As Double
    Dim lFrom As Double
'    Dim lTo As Double
'    lCount = 0
'    llast = FindFreeRow
'    lFrom = Cells(EXP_ROW_RANGE_FROM, COL_DATE).Value2
'    lTo = Cells(EXP_ROW_RANGE_TO, COL_DATE).Value2
'    For i = EXP_ROW_DATA_FIRST To llast
'        lDate = Cells(i, COL_DATE).Value2
'        If lDate >= lFrom And lDate <= lTo Then
'            if lcount < MAX_COUNT
'        End If
'    Next i
End Sub

Public Function IifLong(ByVal aExpression As Boolean, ByVal aTruePart As Long, ByVal aFalsePart As Long) As Long '2009-11-12 rAiner Gruber
    If aExpression Then IifLong = aTruePart Else IifLong = aFalsePart
End Function

Public Function ConvertColToLetter(aColumn As Long) As String '2015-05-11 rAiner Gruber
    ConvertColToLetter = Chr$(64 + aColumn)
End Function

Public Function DateFromTo(aKey As String, Optional aFrom As Boolean = True) As Double ' 2023-10-08 rAiner Gruber
    Dim lYear As Long
    Dim lMonth As Long
    Dim lSplit() As String
    Dim lNextMonth As Double
    lSplit = Split(aKey, "M")
    lYear = Val(lSplit(0)) + 2000
    lMonth = IIf(UBound(lSplit) > 0, Val(lSplit(1)), 0)
    If aFrom Then
        DateFromTo = DateSerial(lYear, lMonth, 1)
    Else
        lNextMonth = DateSerial(lYear, lMonth, 1) + 35
        DateFromTo = DateSerial(year(lNextMonth), month(lNextMonth), 1) - 0.0001 '-1 would get the correct day, but midnight of that. In case a date has a time with it that's later, that would fall out. So creating "a few seconds before the next 1st" date/time stamp here ...
    End If
End Function











