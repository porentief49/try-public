'2023-10-10 little code history ...

Option Explicit

Private Const EXP_ROW_TIMERANGE As Long = 1
Private Const EXP_ROW_RANGE_FROM As Long = 5
Private Const EXP_ROW_BALANCE_MONTH As Long = 5
Private Const EXP_ROW_BALANCEL_TOTAL As Long = 6
Private Const EXP_ROW_RANGE_TO As Long = 6
Private Const EXP_ROW_STATUS As Long = 8
Private Const EXP_ROW_DATA_FIRST As Long = 11
Private Const EXP_ROW_DATA_LAST As Long = 9999
Private Const EXP_COL_TIMERANGE As Long = 4
Private Const COL_DATE As Long = 1
Private Const COL_STATUS As Long = 1
Private Const COL_EXPENSE As Long = 2
Private Const COL_VENDOR As Long = 3
Private Const COL_AMOUNT As Long = 4
Private Const COL_COMMENT As Long = 5
Private Const SHEET_EXPENSES As String = "Auslagen"
Private Const SHEET_BALANCE As String = "Abrechnung"
Private Const BAL_ROW_HEADLINE As Long = 1
Private Const BAL_ROW_RECEIVER As Long = 3
Private Const BAL_ROW_IBAN As Long = 6
Private Const BAL_ROW_AMOUNT As Long = 7
Private Const COL_HEADLINE As Long = 1
Private Const BAL_COL_DATA As Long = 2

Private Function FindFreeRow() As Long
    Const EMPTY_ROW_THRESHOLD As Long = 100
    Dim lRow As Long
    Dim lEmptyCount As Long
    lRow = EXP_ROW_DATA_FIRST
    lEmptyCount = 0
    Do While (lEmptyCount < EMPTY_ROW_THRESHOLD) And (lRow < EXP_ROW_DATA_LAST) '@@@ this is not perfectly correct, because close to the end, it will start overwriting rows
        lEmptyCount = IIf((Len(Trim$(Cells(lRow, COL_DATE).Value2) & Trim$(Cells(lRow, COL_EXPENSE).Value2) & Trim$(Cells(lRow, COL_VENDOR).Value2) & Trim$(Cells(lRow, COL_AMOUNT).Value2) & Trim$(Cells(lRow, COL_COMMENT).Value2)) > 0), 0, lEmptyCount + 1)
        lRow = lRow + 1
    Loop
    FindFreeRow = lRow - EMPTY_ROW_THRESHOLD
End Function

Private Sub UpdateStatus(aStatus As String, aOk As Boolean)
    Dim lColLetterFrom As String
    Dim lColLetterTo As String
    With Cells(EXP_ROW_STATUS, COL_STATUS)
        .Value2 = aStatus
        .Font.Color = IIf(aOk, &HAA00&, &HCC&)
    End With
    lColLetterFrom = ConvertColToLetter(COL_STATUS)
    lColLetterTo = ConvertColToLetter(COL_COMMENT)
    Range(lColLetterFrom & CStr(EXP_ROW_STATUS) & ":" & lColLetterTo & CStr(EXP_ROW_STATUS)).Interior.Color = IIf(aOk, &HDDFFDD, &HDDDDFF)
End Sub

Public Sub CreateBalance()
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
        Call UpdateStatus("OK - " & lTimeRange & " cleared", True)
    Else
        Call UpdateStatus("Clearing " & lTimeRange & " not possible - balance is already 0.00EUR", False)
    End If
End Sub

Public Sub CreateReport()
    Const QR_FILE_PATH As String = "C:\temp\QRCode.png"
    Dim lFileName As String
    Dim lCurrentSheet As Object
    Dim lReportSheet As Worksheet
    Dim lBalance As Double
    Dim lTimeRange As String
    Dim lQrString As String
    On Error GoTo hell
    lBalance = Cells(EXP_ROW_BALANCE_MONTH, COL_AMOUNT).Value2
    lTimeRange = Cells(EXP_ROW_TIMERANGE, EXP_COL_TIMERANGE).Value2
    If Abs(lBalance) < 0.004 Then
        Application.ScreenUpdating = False
        Set lReportSheet = Sheets(SHEET_BALANCE)
        
        'generate QR code
        lQrString = EpcQrString(lReportSheet)
        Call GenerateQRCode(lQrString, QR_FILE_PATH)
        
        'place QR code on sheet
        
        
        lFileName = Replace$(Environ("userprofile") & "\Desktop\" & lReportSheet.Cells(1, 1).Value, " ", "_") & ".pdf"
        Call lReportSheet.ExportAsFixedFormat(xlTypePDF, lFileName, xlQualityStandard, True, False, , , True)
        Call UpdateStatus("OK - report created for " & lTimeRange, True)
        Application.ScreenUpdating = True
    Else
        Call UpdateStatus("Report for " & lTimeRange & " not created - balance != 0.00EUR - please clear first", False)
    End If
    Exit Sub
hell:
    Application.ScreenUpdating = True
    Err.Raise (Err.Number)
End Sub

Private Function EpcQrString(aSheet As Worksheet) As String
    Const NEW_LINE As String = "%0A"
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
    l6_Receiver = Left$(aSheet.Cells(BAL_ROW_RECEIVER, BAL_COL_DATA).Value2, 60)
    l7_IBAN = Replace$(aSheet.Cells(BAL_ROW_IBAN, BAL_COL_DATA).Value2, " ", vbNullString)
    l8_Amount = "EUR" & Format$(aSheet.Cells(BAL_ROW_AMOUNT, BAL_COL_DATA).Value2, "0.00")
    l9_Code = ""
    l10_Ref = ""
    l11_Title = Left$(aSheet.Cells(BAL_ROW_HEADLINE, COL_HEADLINE).Value2, 140)
    l12_Comment = ""
    EpcQrString = l1_ServiceTag & NEW_LINE & l2_Version & NEW_LINE & l3_Encoding & NEW_LINE & l4_Id & NEW_LINE & l5_BIC & NEW_LINE & l6_Receiver & NEW_LINE & l7_IBAN & NEW_LINE & l8_Amount & NEW_LINE & l9_Code & NEW_LINE & l10_Ref & NEW_LINE & l11_Title & NEW_LINE & l12_Comment
End Function

Private Sub GenerateQRCode(inputString As String, outputPath As String) 'credit: https://chat.openai.com/share/4a3043e0-024f-499b-a270-3426e18e9f1a
    Dim xmlHttp As Object
    Dim apiEndpoint As String
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    apiEndpoint = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&ecc=M&data=" & inputString ' check out: https://goqr.me/api/
    xmlHttp.Open "GET", apiEndpoint, False ' Send GET request to the API
    xmlHttp.send ""
    If xmlHttp.Status = 200 Then ' Check if the request was successful
        Dim imageStream As Object ' Create a binary stream for the image
        Set imageStream = CreateObject("ADODB.Stream")
        Call imageStream.Open
        imageStream.Type = 1 ' Binary
        Call imageStream.Write(xmlHttp.responseBody)
        Call imageStream.SaveToFile(outputPath, 2)  ' Save the image to the specified file path ' Overwrite existing file
        imageStream.Close ' Clean up
        Set imageStream = Nothing
    Else
        MsgBox "Failed to generate QR code. HTTP Status: " & xmlHttp.Status
    End If
    xmlHttp.abort ' Clean up
    Set xmlHttp = Nothing
End Sub

Public Sub LoadAndDisplayImage()
'    Const IMAGE_WIDTH_HEIGHT As Double = 100
    Const IMAGE_HEIGHT_CORRECTION As Double = 1.118 ' weird - Excel will display squares a little taller than they should be
    Const IMAGE_COL_MARGIN As Double = 5
    Const IMAGE_NAME As String = "QrCode"
    Dim imagePath As String
    Dim ws As Worksheet
    Dim pic As Picture
    Dim shp As Shape
    Dim lHeight As Double
    Dim lWidth As Double

    ' Specify the path to your PNG image file
    imagePath = "C:\temp\QRCode.png" ' Replace with the actual file path

    ' Set the worksheet where you want to display the image
    Set ws = ThisWorkbook.Sheets("Abrechnung") ' Change "Sheet1" to your target sheet

    For Each shp In ws.Shapes
        If shp.Name = IMAGE_NAME Then
            shp.Delete
            Exit For
        End If
    Next shp

    ' Create a new Picture object and load the image
    Set pic = ws.Pictures.Insert(imagePath)
    pic.Name = IMAGE_NAME
    
    ActiveSheet.Shapes.Range(Array(IMAGE_NAME)).Select
    Selection.ShapeRange.LockAspectRatio = msoFalse
    
    lHeight = ws.Cells(8, 3).Top - ws.Cells(3, 3).Top
    lWidth = lHeight / IMAGE_HEIGHT_CORRECTION
    ' Position and resize the image as needed
    With pic
        .Left = ws.Cells(3, 3).Left - lWidth - IMAGE_COL_MARGIN ' Change the coordinates as needed
        .Top = ws.Cells(3, 3).Top
        .Locked = False
        .Width = lWidth
        .Height = lHeight
'        .Width = 100
'        .Height = 100 * 1.118
'        .LockAspectRatio = msoFalse
        ' You can set .Width and .Height properties to resize the image
'        .Name = IMAGE_NAME
    End With
'    Selection.ShapeRange.ScaleWidth 1.4423076923, msoFalse, msoScaleFromTopLeft
'    Selection.ShapeRange.ScaleHeight 2.5089445438, msoFalse, msoScaleFromTopLeft

'    With pic
'        .Left = ws.Cells(1, 1).Left ' Change the coordinates as needed
'        .Top = ws.Cells(1, 1).Top
'        .Locked = False
'        .LockAspectRatio = msoFalse
        ' You can set .Width and .Height properties to resize the image
'        .Name = "huhu"
'    End With
    
    
    
    ' Clean up
    Set pic = Nothing
End Sub

Sub UnlockAspectRatioOfPicture()
    Dim ws As Worksheet
    Dim pic As Picture

    ' Set the worksheet and the picture you want to work with
    Set ws = ThisWorkbook.Sheets("Abrechnung") ' Change "Sheet1" to your target sheet
    Set pic = ws.Pictures("huhu") ' Change "Picture 1" to the name of your picture

    ' Check if the picture exists
    If Not pic Is Nothing Then
        ' Unlock the aspect ratio
        pic.LockAspectRatio = msoFalse
    Else
        MsgBox "Picture not found on the specified sheet."
    End If
End Sub

Public Sub TestQr()
    ' Usage example:
    Dim inputString As String
    Dim outputPath As String
    inputString = "BCD|002|1|SCT||Paul|DE30702501500027525005|EUR1.00||Zweck"
    outputPath = "C:\temp\QRCode.png"
    Call GenerateQRCode(Replace$(inputString, "|", "%0A"), outputPath)
    Debug.Print "QR code generated and saved as " & outputPath
End Sub

Private Sub SortEntries()
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

Private Sub CopyExpenses()
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

Private Function ConvertColToLetter(aColumn As Long) As String
    ConvertColToLetter = Chr$(64 + aColumn)
End Function

Private Function DateFromTo(aKey As String, Optional aFrom As Boolean = True) As Double
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












