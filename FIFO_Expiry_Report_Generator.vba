Sub SegmentExpiryItems()
    On Error GoTo ErrorHandler
    
    Dim wsSource As Worksheet
    Dim wsExpired As Worksheet, ws1Month As Worksheet, ws2Month As Worksheet
    Dim ws3Month As Worksheet, wsTotal As Worksheet
    Dim lastRow As Long, i As Long
    Dim expiryCol As Long
    Dim expiryValue As Variant
    Dim expiryDate As Date
    Dim todayDate As Date
    Dim daysRemaining As Long
    Dim sourceFilePath As String, outputFilePath As String
    Dim finalFileName As String
    
    Dim rowExpired As Long, row1Month As Long, row2Month As Long
    Dim row3Month As Long, rowTotal As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    todayDate = Date
    sourceFilePath = ThisWorkbook.Path
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("Total")
    On Error GoTo ErrorHandler
    
    If wsSource Is Nothing Then
        MsgBox "Error: Could not find sheet named 'Total'.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    If wsSource.Cells(1, 1).Value = "" Then
        MsgBox "No data found in 'Total' sheet.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data to process!", vbExclamation
        GoTo CleanupAndExit
    End If
    
    expiryCol = 14
    
    Dim wbOutput As Workbook
    Set wbOutput = Workbooks.Add
    
    Application.DisplayAlerts = False
    Do While wbOutput.Worksheets.Count > 1
        wbOutput.Worksheets(wbOutput.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    Set wsExpired = wbOutput.Worksheets(1)
    wsExpired.Name = ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
    wsExpired.Tab.Color = RGB(255, 102, 102)
    
    Set ws1Month = wbOutput.Worksheets.Add(After:=wsExpired)
    ws1Month.Name = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws1Month.Tab.Color = RGB(255, 153, 153)
    
    Set ws2Month = wbOutput.Worksheets.Add(After:=ws1Month)
    ws2Month.Name = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
    ws2Month.Tab.Color = RGB(255, 192, 128)
    
    Set ws3Month = wbOutput.Worksheets.Add(After:=ws2Month)
    ws3Month.Name = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws3Month.Tab.Color = RGB(255, 230, 153)
    
    Set wsTotal = wbOutput.Worksheets.Add(After:=ws3Month)
    wsTotal.Name = ChrW(&H6A9) & ChrW(&H627) & ChrW(&H645) & ChrW(&H644)
    wsTotal.Tab.Color = RGB(146, 208, 80)
    
    Call SetupReorganizedHeaders(wsExpired)
    Call SetupReorganizedHeaders(ws1Month)
    Call SetupReorganizedHeaders(ws2Month)
    Call SetupReorganizedHeaders(ws3Month)
    Call SetupReorganizedHeaders(wsTotal)
    
    rowExpired = 2
    row1Month = 2
    row2Month = 2
    row3Month = 2
    rowTotal = 2
    
    For i = 2 To lastRow
        expiryValue = wsSource.Cells(i, expiryCol).Value
        
        If expiryValue <> "" And Not IsEmpty(expiryValue) Then
            On Error Resume Next
            expiryDate = DateValue(expiryValue)
            On Error GoTo ErrorHandler
            
            daysRemaining = expiryDate - todayDate
            
            If daysRemaining < 0 Then
                Call CopyReorganizedRow(wsSource, wsExpired, i, rowExpired, daysRemaining)
                Call CopyReorganizedRow(wsSource, wsTotal, i, rowTotal, daysRemaining)
                rowExpired = rowExpired + 1
                rowTotal = rowTotal + 1
            ElseIf daysRemaining < 30 Then
                Call CopyReorganizedRow(wsSource, ws1Month, i, row1Month, daysRemaining)
                Call CopyReorganizedRow(wsSource, wsTotal, i, rowTotal, daysRemaining)
                row1Month = row1Month + 1
                rowTotal = rowTotal + 1
            ElseIf daysRemaining < 60 Then
                Call CopyReorganizedRow(wsSource, ws2Month, i, row2Month, daysRemaining)
                Call CopyReorganizedRow(wsSource, wsTotal, i, rowTotal, daysRemaining)
                row2Month = row2Month + 1
                rowTotal = rowTotal + 1
            ElseIf daysRemaining < 90 Then
                Call CopyReorganizedRow(wsSource, ws3Month, i, row3Month, daysRemaining)
                Call CopyReorganizedRow(wsSource, wsTotal, i, rowTotal, daysRemaining)
                row3Month = row3Month + 1
                rowTotal = rowTotal + 1
            End If
        End If
    Next i
    
    Dim ws As Worksheet
    Dim lastRowSheet As Long
    
    For Each ws In wbOutput.Worksheets
        Dim skipSheet As Boolean
        skipSheet = False
        
        If InStr(ws.Name, "10") > 0 Or InStr(ws.Name, "RSM") > 0 Or InStr(ws.Name, ChrW(&H623) & ChrW(&H62F) & ChrW(&H627) & ChrW(&H621)) > 0 Or InStr(ws.Name, ChrW(&H645) & ChrW(&H644) & ChrW(&H62E) & ChrW(&H635)) > 0 Then
            skipSheet = True
        End If
        
        If Not skipSheet Then
            With ws
                lastRowSheet = .Cells(.Rows.Count, 1).End(xlUp).Row
                
                If lastRowSheet > 1 Then
                    .Sort.SortFields.Clear
                    .Sort.SortFields.Add Key:=.Range("E2:E" & lastRowSheet), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    With .Sort
                        .SetRange ws.Range("A1:G" & lastRowSheet)
                        .Header = xlYes
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                
                Dim totalQtySum As Long
                totalQtySum = Application.WorksheetFunction.Sum(.Range("E2:E" & lastRowSheet))
                lastRowSheet = lastRowSheet + 1
                
                .Cells(lastRowSheet, 1).Value = "TOTAL"
                .Cells(lastRowSheet, 5).Value = totalQtySum
                .Cells(lastRowSheet, 5).NumberFormat = "#,##0"
                .Cells(lastRowSheet, 5).Font.Bold = True
                .Cells(lastRowSheet, 5).Interior.Color = RGB(146, 208, 80)
                .Cells(lastRowSheet, 5).HorizontalAlignment = xlCenter
                .Cells(lastRowSheet, 5).VerticalAlignment = xlCenter
                .Cells(lastRowSheet, 5).Borders.LineStyle = xlContinuous
                .Cells(lastRowSheet, 5).Borders.Weight = xlMedium
                .Cells(lastRowSheet, 5).Borders.Color = RGB(0, 0, 0)
                
                .Range("A" & lastRowSheet & ":D" & lastRowSheet).Interior.ColorIndex = xlNone
                .Range("A" & lastRowSheet & ":D" & lastRowSheet).Borders.LineStyle = xlNone
                .Range("F" & lastRowSheet & ":G" & lastRowSheet).Interior.ColorIndex = xlNone
                .Range("F" & lastRowSheet & ":G" & lastRowSheet).Borders.LineStyle = xlNone
                
                With .Range(.Cells(1, 1), .Cells(lastRowSheet - 1, 7))
                    .Font.Name = "Calibri"
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = False
                    .ShrinkToFit = False
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(200, 200, 200)
                End With
                
                With .Range("A1:G1")
                    .Font.Name = "Calibri"
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = False
                    .ShrinkToFit = False
                    .Interior.Color = RGB(68, 114, 196)
                End With
                
                .Cells(1, 5).Interior.Color = RGB(255, 255, 0)
                .Cells(1, 5).Font.Color = RGB(0, 0, 0)
                .Cells(1, 5).Font.Bold = True
                
                .Cells(1, 7).Interior.Color = RGB(255, 165, 0)
                .Cells(1, 7).Font.Color = RGB(0, 0, 0)
                .Cells(1, 7).Font.Bold = True
                
                Dim col As Long
                For col = 1 To 7
                    .Columns(col).AutoFit
                    .Columns(col).ColumnWidth = .Columns(col).ColumnWidth + 3
                Next col
                
                .Cells.EntireRow.AutoFit
                .Cells.EntireColumn.AutoFit
                
                .Range("A1").AutoFilter
                .Activate
                .Range("A2").Select
                ActiveWindow.FreezePanes = True
                .Range("A1").Select
            End If
        End With
        End If
    Next ws
    
    ' CREATE GENERAL SUMMARY SHEET (ملخص عام) - NEW AND FIRST
    Dim wsGeneralSummary As Worksheet
    Set wsGeneralSummary = wbOutput.Worksheets.Add(Before:=wbOutput.Worksheets(1))
    wsGeneralSummary.Name = ChrW(&H645) & ChrW(&H644) & ChrW(&H62E) & ChrW(&H635) & " " & ChrW(&H639) & ChrW(&H627) & ChrW(&H645)
    wsGeneralSummary.Tab.Color = RGB(255, 192, 0)
    
    Call CreateGeneralSummarySheet(wsGeneralSummary, rowExpired - 2, row1Month - 2, row2Month - 2, row3Month - 2, wsExpired, ws1Month, ws2Month, ws3Month)
    
    ' CREATE TOP 10 STORES SHEET (أعلى 10 وکلاء) - SEPARATE FROM SUMMARY
    Dim wsSummary As Worksheet
    Set wsSummary = wbOutput.Worksheets.Add(Before:=wsExpired)
    wsSummary.Name = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621)
    wsSummary.Tab.Color = RGB(68, 114, 196)
    
    Call CreateTop10StoresOnly(wsSummary, wsExpired)
    
    ' CREATE TOP 10 PRODUCTS ANALYSIS
    Dim wsProductAnalysis As Worksheet
    Set wsProductAnalysis = wbOutput.Worksheets.Add(Before:=wsSummary)
    wsProductAnalysis.Name = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A)
    wsProductAnalysis.Tab.Color = RGB(76, 175, 80)
    
    Call CreateProductAnalysis(wsProductAnalysis, wsExpired, ws1Month, ws2Month, ws3Month)
    
    ' CREATE RSM DASHBOARD
    Dim wsDash As Worksheet
    Set wsDash = wbOutput.Worksheets.Add(Before:=wsProductAnalysis)
    wsDash.Name = ChrW(&H623) & ChrW(&H62F) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H645) & ChrW(&H62F) & ChrW(&H631) & ChrW(&H627) & ChrW(&H621) & " RSM"
    wsDash.Tab.Color = RGB(192, 0, 0)
    
    Call CreateDashboard(wsDash, wsExpired, ws1Month, ws2Month, ws3Month)
    
    ' Activate General Summary as the first sheet
    wsGeneralSummary.Activate
    
    Dim baseFileName As String
    Dim fileExtension As String
    
    baseFileName = "FIFO Expiry Report - " & Format(Date, "dd-mmm-yyyy")
    fileExtension = ".xlsx"
    finalFileName = baseFileName & fileExtension
    outputFilePath = sourceFilePath & "\" & finalFileName
    
    On Error Resume Next
    If Dir(outputFilePath) <> "" Then
        Dim archivePath As String
        Dim versionNum As Long
        Dim archiveFile As String
        
        archivePath = sourceFilePath & "\M. FIFO Archive"
        If Dir(archivePath, vbDirectory) = "" Then MkDir archivePath
        
        versionNum = 1
        Do While Dir(archivePath & "\" & baseFileName & " v" & versionNum & fileExtension) <> ""
            versionNum = versionNum + 1
        Loop
        
        archiveFile = archivePath & "\" & baseFileName & " v" & versionNum & fileExtension
        FileCopy outputFilePath, archiveFile
        SetAttr outputFilePath, vbNormal
        Kill outputFilePath
    End If
    On Error GoTo ErrorHandler
    
    wbOutput.SaveAs Filename:=outputFilePath, FileFormat:=xlOpenXMLWorkbook
    
    On Error Resume Next
    wbOutput.Close SaveChanges:=False
    On Error GoTo ErrorHandler
    
    wbOutput.SaveAs Filename:=outputFilePath, FileFormat:=xlOpenXMLWorkbook
    wbOutput.Close SaveChanges:=False
    
CleanupAndExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    On Error Resume Next
    If Err.Number = 0 Then
        MsgBox "Success!" & vbCrLf & vbCrLf & _
               "Saved as: " & finalFileName & vbCrLf & _
               "Location: " & sourceFilePath & vbCrLf & vbCrLf & _
               "SUMMARY:" & vbCrLf & _
               "• Expired: " & (rowExpired - 2) & vbCrLf & _
               "• Less than 1 month: " & (row1Month - 2) & vbCrLf & _
               "• Less than 2 months: " & (row2Month - 2) & vbCrLf & _
               "• Less than 3 months: " & (row3Month - 2), vbInformation
    End If
    On Error GoTo 0
    
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    If Dir(outputFilePath) = "" Then
        MsgBox "Error: " & Err.Description, vbCritical
    Else
        MsgBox "Success!" & vbCrLf & vbCrLf & _
               "Saved as: " & finalFileName & vbCrLf & _
               "Location: " & sourceFilePath & vbCrLf & vbCrLf & _
               "SUMMARY:" & vbCrLf & _
               "• Expired: " & (rowExpired - 2) & vbCrLf & _
               "• Less than 1 month: " & (row1Month - 2) & vbCrLf & _
               "• Less than 2 months: " & (row2Month - 2) & vbCrLf & _
               "• Less than 3 months: " & (row3Month - 2), vbInformation
    End If
    On Error GoTo 0
End Sub

Private Sub SetupReorganizedHeaders(ws As Worksheet)
    ws.Cells.Font.Name = "Arial"
    
    ws.Cells(1, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H6CC) & ChrW(&H644)
    ws.Cells(1, 2).Value = ChrW(&H645) & ChrW(&H62F) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H64A)
    ws.Cells(1, 3).Value = ChrW(&H6A9) & ChrW(&H648) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(1, 4).Value = ChrW(&H625) & ChrW(&H633) & ChrW(&H645) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(1, 5).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629)
    ws.Cells(1, 6).Value = ChrW(&H62A) & ChrW(&H627) & ChrW(&H631) & ChrW(&H64A) & ChrW(&H62E) & " " & ChrW(&H627) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
    ws.Cells(1, 7).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H623) & ChrW(&H64A) & ChrW(&H627) & ChrW(&H645) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H62A) & ChrW(&H628) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H629)
End Sub

Private Sub CopyReorganizedRow(wsSource As Worksheet, wsTarget As Worksheet, sourceRow As Long, targetRow As Long, daysRemaining As Long)
    wsTarget.Cells(targetRow, 1).Value = Trim(CStr(wsSource.Cells(sourceRow, 2).Value))
    wsTarget.Cells(targetRow, 1).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 1).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 2).Value = Trim(CStr(wsSource.Cells(sourceRow, 17).Value))
    wsTarget.Cells(targetRow, 2).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 2).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 3).Value = Trim(CStr(wsSource.Cells(sourceRow, 7).Value))
    wsTarget.Cells(targetRow, 3).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 3).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 4).Value = Trim(CStr(wsSource.Cells(sourceRow, 10).Value))
    wsTarget.Cells(targetRow, 4).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 4).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 5).Value = Trim(CStr(wsSource.Cells(sourceRow, 13).Value))
    wsTarget.Cells(targetRow, 5).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 5).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 6).Value = wsSource.Cells(sourceRow, 14).Value
    wsTarget.Cells(targetRow, 6).NumberFormat = "yyyy-mm-dd"
    wsTarget.Cells(targetRow, 6).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 6).VerticalAlignment = xlCenter
    
    wsTarget.Cells(targetRow, 7).Value = daysRemaining
    wsTarget.Cells(targetRow, 7).HorizontalAlignment = xlCenter
    wsTarget.Cells(targetRow, 7).VerticalAlignment = xlCenter
    
    If daysRemaining < 0 Then
        wsTarget.Cells(targetRow, 7).Interior.Color = RGB(255, 102, 102)
    ElseIf daysRemaining < 30 Then
        wsTarget.Cells(targetRow, 7).Interior.Color = RGB(255, 153, 153)
    ElseIf daysRemaining < 60 Then
        wsTarget.Cells(targetRow, 7).Interior.Color = RGB(255, 192, 128)
    ElseIf daysRemaining < 90 Then
        wsTarget.Cells(targetRow, 7).Interior.Color = RGB(255, 230, 153)
    End If
End Sub

Private Sub CreateGeneralSummarySheet(ws As Worksheet, expiredCount As Long, month1Count As Long, month2Count As Long, month3Count As Long, wsExpired As Worksheet, ws1Month As Worksheet, ws2Month As Worksheet, ws3Month As Worksheet)
    Dim totalItems As Long
    totalItems = expiredCount + month1Count + month2Count + month3Count
    
    Dim qtyExpired As Long, qty1Month As Long, qty2Month As Long, qty3Month As Long
    Dim lastRowCalc As Long
    
    qtyExpired = 0: qty1Month = 0: qty2Month = 0: qty3Month = 0
    
    On Error Resume Next
    If expiredCount > 0 Then
        lastRowCalc = wsExpired.Cells(wsExpired.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 2 Then
            qtyExpired = Application.WorksheetFunction.Sum(wsExpired.Range("E2:E" & (lastRowCalc - 1)))
        End If
    End If
    
    If month1Count > 0 Then
        lastRowCalc = ws1Month.Cells(ws1Month.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 2 Then
            qty1Month = Application.WorksheetFunction.Sum(ws1Month.Range("E2:E" & (lastRowCalc - 1)))
        End If
    End If
    
    If month2Count > 0 Then
        lastRowCalc = ws2Month.Cells(ws2Month.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 2 Then
            qty2Month = Application.WorksheetFunction.Sum(ws2Month.Range("E2:E" & (lastRowCalc - 1)))
        End If
    End If
    
    If month3Count > 0 Then
        lastRowCalc = ws3Month.Cells(ws3Month.Rows.Count, 1).End(xlUp).Row
        If lastRowCalc > 2 Then
            qty3Month = Application.WorksheetFunction.Sum(ws3Month.Range("E2:E" & (lastRowCalc - 1)))
        End If
    End If
    On Error GoTo 0
    
    Dim totalQty As Long
    totalQty = qtyExpired + qty1Month + qty2Month + qty3Month
    
    ws.Cells.Font.Name = "Arial"
    
    With ws
        .Range("A1:D1").Merge
        .Range("A1").Value = ChrW(&H62A) & ChrW(&H642) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629) & " FIFO"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 255, 255)
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").VerticalAlignment = xlCenter
        .Range("A1:D1").Borders.LineStyle = xlContinuous
        .Range("A1:D1").Borders.Weight = xlMedium
        .Range("A1:D1").Borders.Color = RGB(0, 0, 0)
        
        .Range("A2").Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H641) & ChrW(&H648) & ChrW(&H639)
        .Range("B2").Value = ChrW(&H639) & ChrW(&H62F) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A)
        .Range("C2").Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629)
        .Range("D2").Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62D) & ChrW(&H627) & ChrW(&H644) & ChrW(&H629)
        
        .Range("A2:D2").Font.Bold = True
        .Range("A2:D2").Font.Size = 11
        .Range("A2:D2").Font.Color = RGB(255, 255, 255)
        .Range("A2:D2").Interior.Color = RGB(68, 114, 196)
        .Range("A2:D2").HorizontalAlignment = xlCenter
        .Range("A2:D2").VerticalAlignment = xlCenter
        .Range("A2:D2").WrapText = False
        .Range("A2:D2").ShrinkToFit = False
        .Range("A2:D2").Borders.LineStyle = xlContinuous
        .Range("A2:D2").Borders.Weight = xlMedium
        .Range("A2:D2").Borders.Color = RGB(0, 0, 0)
        
        .Cells(3, 1).Value = ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
        .Cells(3, 2).Value = expiredCount
        .Cells(3, 3).Value = qtyExpired
        .Cells(3, 4).Value = ChrW(&H62E) & ChrW(&H637) & ChrW(&H631)
        .Range("A3:D3").Interior.Color = RGB(255, 102, 102)
        .Range("A3:D3").WrapText = False
        .Range("A3:D3").ShrinkToFit = False
        .Range("A3:D3").Borders.LineStyle = xlContinuous
        .Range("A3:D3").Borders.Weight = xlThin
        .Range("A3:D3").Borders.Color = RGB(0, 0, 0)
        
        .Cells(4, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & " " & ChrW(&H648) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H62F)
        .Cells(4, 2).Value = month1Count
        .Cells(4, 3).Value = qty1Month
        .Cells(4, 4).Value = ChrW(&H639) & ChrW(&H627) & ChrW(&H62C) & ChrW(&H644)
        .Range("A4:D4").Interior.Color = RGB(255, 153, 153)
        .Range("A4:D4").WrapText = False
        .Range("A4:D4").ShrinkToFit = False
        .Range("A4:D4").Borders.LineStyle = xlContinuous
        .Range("A4:D4").Borders.Weight = xlThin
        .Range("A4:D4").Borders.Color = RGB(0, 0, 0)
        
        .Cells(5, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
        .Cells(5, 2).Value = month2Count
        .Cells(5, 3).Value = qty2Month
        .Cells(5, 4).Value = ChrW(&H62A) & ChrW(&H62D) & ChrW(&H630) & ChrW(&H64A) & ChrW(&H631)
        .Range("A5:D5").Interior.Color = RGB(255, 192, 128)
        .Range("A5:D5").WrapText = False
        .Range("A5:D5").ShrinkToFit = False
        .Range("A5:D5").Borders.LineStyle = xlContinuous
        .Range("A5:D5").Borders.Weight = xlThin
        .Range("A5:D5").Borders.Color = RGB(0, 0, 0)
        
        .Cells(6, 1).Value = ChrW(&H623) & ChrW(&H642) & ChrW(&H644) & " " & ChrW(&H645) & ChrW(&H646) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(6, 2).Value = month3Count
        .Cells(6, 3).Value = qty3Month
        .Cells(6, 4).Value = ChrW(&H62A) & ChrW(&H628) & ChrW(&H639) & ChrW(&H64A) & ChrW(&H62F)
        .Range("A6:D6").Interior.Color = RGB(255, 230, 153)
        .Range("A6:D6").WrapText = False
        .Range("A6:D6").ShrinkToFit = False
        .Range("A6:D6").Borders.LineStyle = xlContinuous
        .Range("A6:D6").Borders.Weight = xlThin
        .Range("A6:D6").Borders.Color = RGB(0, 0, 0)
        
        .Range("B3:C6").NumberFormat = "#,##0"
        .Range("A3:D6").HorizontalAlignment = xlCenter
        .Range("A3:D6").VerticalAlignment = xlCenter
        .Range("A3:D6").Font.Size = 11
        
        .Cells(7, 2).Value = totalItems
        .Cells(7, 3).Value = totalQty
        
        .Range("B7:C7").NumberFormat = "#,##0"
        .Range("B7:C7").HorizontalAlignment = xlCenter
        .Range("B7:C7").VerticalAlignment = xlCenter
        .Range("B7:C7").Font.Size = 11
        .Range("B7:C7").Interior.Color = RGB(255, 255, 0)
        .Range("B7:C7").Font.Bold = True
        .Range("B7:C7").WrapText = False
        .Range("B7:C7").ShrinkToFit = False
        .Range("B7:C7").Borders.LineStyle = xlContinuous
        .Range("B7:C7").Borders.Weight = xlMedium
        .Range("B7:C7").Borders.Color = RGB(0, 0, 0)
        
        .Range("A7").Interior.ColorIndex = xlNone
        .Range("A7").Borders.LineStyle = xlNone
        .Range("D7").Interior.ColorIndex = xlNone
        .Range("D7").Borders.LineStyle = xlNone
        
        .Columns("A:D").AutoFit
        Dim colAdjust As Long
        For colAdjust = 1 To 4
            .Columns(colAdjust).ColumnWidth = .Columns(colAdjust).ColumnWidth + 3
        Next colAdjust
        
        .Range("A1:D7").WrapText = False
        .Range("A1:D7").ShrinkToFit = False
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        
        .Range("A3").Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
    End With
End Sub

Private Sub CreateTop10StoresOnly(ws As Worksheet, wsE As Worksheet)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lr As Long, i As Long
    lr = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    
    If lr > 1 Then
        For i = 2 To lr
            Dim store As String, qty As Long
            store = Trim(CStr(wsE.Cells(i, 1).Value))
            On Error Resume Next
            qty = CLng(wsE.Cells(i, 5).Value)
            On Error GoTo 0
            
            If Len(store) > 0 And qty > 0 And UCase(store) <> "TOTAL" Then
                If Not dict.Exists(store) Then
                    dict.Add store, Array(wsE.Cells(i, 2).Value, 0, 0)
                End If
                Dim arr As Variant
                arr = dict(store)
                arr(1) = arr(1) + qty
                arr(2) = arr(2) + 1
                dict(store) = arr
            End If
        Next i
    End If
    
    Dim sortedKeys() As String, sortedVals() As Long
    ReDim sortedKeys(0 To dict.Count - 1)
    ReDim sortedVals(0 To dict.Count - 1)
    
    Dim j As Long, k As Variant
    j = 0
    For Each k In dict.Keys
        sortedKeys(j) = CStr(k)
        sortedVals(j) = dict(k)(1)
        j = j + 1
    Next k
    
    Dim m As Long, n As Long, tempK As String, tempV As Long
    For m = 0 To UBound(sortedVals) - 1
        For n = m + 1 To UBound(sortedVals)
            If sortedVals(m) < sortedVals(n) Then
                tempV = sortedVals(m)
                sortedVals(m) = sortedVals(n)
                sortedVals(n) = tempV
                
                tempK = sortedKeys(m)
                sortedKeys(m) = sortedKeys(n)
                sortedKeys(n) = tempK
            End If
        Next n
    Next m
    
    ws.Cells.Font.Name = "Arial"
    
    With ws
        .Range("A1:D1").Merge
        .Cells(1, 1).Value = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H644) & ChrW(&H627) & ChrW(&H621)
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Interior.Color = RGB(68, 114, 196)
        .Cells(1, 1).Font.Color = RGB(255, 255, 255)
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(1, 1).VerticalAlignment = xlCenter
        .Cells(1, 1).WrapText = False
        .Cells(1, 1).ShrinkToFit = False
        .Range("A1:D1").Borders.LineStyle = xlContinuous
        .Range("A1:D1").Borders.Weight = xlMedium
        .Range("A1:D1").Borders.Color = RGB(0, 0, 0)
        
        Dim sr As Long
        sr = 2
        
        .Cells(sr, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62A) & ChrW(&H631) & ChrW(&H62A) & ChrW(&H64A) & ChrW(&H628)
        .Cells(sr, 2).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H648) & ChrW(&H6A9) & ChrW(&H6CC) & ChrW(&H644)
        .Cells(sr, 3).Value = ChrW(&H645) & ChrW(&H62F) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H64A)
        .Cells(sr, 4).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H643) & ChrW(&H645) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H629)
        .Range("A" & sr & ":D" & sr).Font.Bold = True
        .Range("A" & sr & ":D" & sr).Interior.Color = RGB(217, 217, 217)
        .Range("A" & sr & ":D" & sr).HorizontalAlignment = xlCenter
        .Range("A" & sr & ":D" & sr).VerticalAlignment = xlCenter
        .Range("A" & sr & ":D" & sr).WrapText = False
        .Range("A" & sr & ":D" & sr).ShrinkToFit = False
        .Range("A" & sr & ":D" & sr).Borders.LineStyle = xlContinuous
        .Range("A" & sr & ":D" & sr).Borders.Weight = xlMedium
        .Range("A" & sr & ":D" & sr).Borders.Color = RGB(0, 0, 0)
        sr = sr + 1
        
        Dim rank As Long
        For rank = 0 To 9
            If rank > UBound(sortedKeys) Then Exit For
            arr = dict(sortedKeys(rank))
            .Cells(sr, 1).Value = rank + 1
            .Cells(sr, 2).Value = sortedKeys(rank)
            .Cells(sr, 3).Value = arr(0)
            .Cells(sr, 4).Value = arr(1)
            
            .Range("A" & sr & ":D" & sr).Interior.ColorIndex = xlNone
            .Range("A" & sr & ":D" & sr).HorizontalAlignment = xlCenter
            .Range("A" & sr & ":D" & sr).VerticalAlignment = xlCenter
            .Range("A" & sr & ":D" & sr).WrapText = False
            .Range("A" & sr & ":D" & sr).ShrinkToFit = False
            .Range("A" & sr & ":D" & sr).Borders.LineStyle = xlContinuous
            .Range("A" & sr & ":D" & sr).Borders.Weight = xlThin
            .Range("A" & sr & ":D" & sr).Borders.Color = RGB(0, 0, 0)
            .Cells(sr, 4).NumberFormat = "#,##0"
            sr = sr + 1
        Next rank
        
        .Cells(sr, 4).Formula = "=SUM(D3:D12)"
        .Range("A" & sr & ":C" & sr).Interior.ColorIndex = xlNone
        .Range("D" & sr).Font.Bold = True
        .Range("D" & sr).Interior.Color = RGB(146, 208, 80)
        .Range("D" & sr).HorizontalAlignment = xlCenter
        .Range("D" & sr).VerticalAlignment = xlCenter
        .Range("D" & sr).WrapText = False
        .Range("D" & sr).ShrinkToFit = False
        .Range("D" & sr).Borders.LineStyle = xlContinuous
        .Range("D" & sr).Borders.Weight = xlMedium
        .Range("D" & sr).Borders.Color = RGB(0, 0, 0)
        .Cells(sr, 4).NumberFormat = "#,##0"
        
        .Columns("A:D").AutoFit
        Dim colNum As Long
        For colNum = 1 To 4
            .Columns(colNum).ColumnWidth = .Columns(colNum).ColumnWidth + 3
        Next colNum
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
    End With
End Sub

Private Sub CreateDashboard(ws As Worksheet, wsExp As Worksheet, ws1M As Worksheet, ws2M As Worksheet, ws3M As Worksheet)
    ws.Cells.Font.Name = "Arial"
    
    Call AddRSMPerformance(ws, wsExp, ws1M, ws2M, ws3M, 1)
    
    ws.Cells.EntireColumn.AutoFit
    ws.Cells.EntireRow.AutoFit
    
    ws.Columns("A:F").AutoFit
    Dim colRSM As Long
    For colRSM = 1 To 6
        ws.Columns(colRSM).ColumnWidth = ws.Columns(colRSM).ColumnWidth + 3
    Next colRSM
    
    ws.Rows.AutoFit
    ws.Range("A1").Select
End Sub

Private Sub CreateProductAnalysis(ws As Worksheet, wsExp As Worksheet, ws1M As Worksheet, ws2M As Worksheet, ws3M As Worksheet)
    ws.Cells.Font.Name = "Arial"
    
    ws.Range("A1:G1").Merge
    ws.Cells(1, 1).Value = ChrW(&H623) & ChrW(&H639) & ChrW(&H644) & ChrW(&H649) & " 10 " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C) & ChrW(&H627) & ChrW(&H62A) & " " & ChrW(&H645) & ChrW(&H639) & ChrW(&H631) & ChrW(&H636) & ChrW(&H629) & " " & ChrW(&H644) & ChrW(&H644) & ChrW(&H627) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H627) & ChrW(&H621) & " (" & ChrW(&H623) & ChrW(&H648) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H64A) & ChrW(&H629) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629) & ")"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Interior.Color = RGB(76, 175, 80)
    ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Cells(1, 1).VerticalAlignment = xlCenter
    ws.Cells(1, 1).WrapText = False
    ws.Cells(1, 1).ShrinkToFit = False
    ws.Range("A1:G1").Borders.LineStyle = xlContinuous
    ws.Range("A1:G1").Borders.Weight = xlMedium
    
    ws.Rows(2).RowHeight = 10
    
    Dim dictProducts As Object
    Set dictProducts = CreateObject("Scripting.Dictionary")
    
    Call CollectProductDataWithID(dictProducts, wsExp, 0)
    Call CollectProductDataWithID(dictProducts, ws1M, 1)
    Call CollectProductDataWithID(dictProducts, ws2M, 2)
    Call CollectProductDataWithID(dictProducts, ws3M, 3)
    
    Dim sortedProducts As Object
    Set sortedProducts = SortProductsByTotal(dictProducts)
    
    ws.Cells(3, 1).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H62A) & ChrW(&H631) & ChrW(&H62A) & ChrW(&H6CC) & ChrW(&H628)
    ws.Cells(3, 2).Value = ChrW(&H6A9) & ChrW(&H648) & ChrW(&H62F) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(3, 3).Value = ChrW(&H625) & ChrW(&H633) & ChrW(&H645) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H62C)
    ws.Cells(3, 4).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & ChrW(&H629)
    ws.Cells(3, 5).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws.Cells(3, 6).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
    ws.Cells(3, 7).Value = ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
    ws.Cells(3, 8).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
    
    ws.Range("A3:H3").Font.Bold = True
    ws.Range("A3:H3").Font.Size = 11
    ws.Range("A3:H3").Font.Color = RGB(255, 255, 255)
    ws.Range("A3:H3").Interior.Color = RGB(68, 114, 196)
    ws.Range("A3:H3").HorizontalAlignment = xlCenter
    ws.Range("A3:H3").VerticalAlignment = xlCenter
    ws.Range("A3:H3").WrapText = False
    ws.Range("A3:H3").ShrinkToFit = False
    ws.Range("A3:H3").Borders.LineStyle = xlContinuous
    ws.Range("A3:H3").Borders.Weight = xlMedium
    
    ws.Cells(3, 4).Interior.Color = RGB(255, 102, 102)
    ws.Cells(3, 5).Interior.Color = RGB(255, 153, 153)
    ws.Cells(3, 8).Interior.Color = RGB(255, 255, 0)
    ws.Cells(3, 8).Font.Color = RGB(0, 0, 0)
    
    Dim rank As Long, rowNum As Long
    Dim productKey As Variant
    
    rank = 1
    rowNum = 4
    
    For Each productKey In sortedProducts.Keys
        If rank > 10 Then Exit For
        
        Dim prodData As Variant
        prodData = sortedProducts(productKey)
        
        Dim totalQty As Long
        totalQty = prodData(1) + prodData(2) + prodData(3) + prodData(4)
        
        ws.Cells(rowNum, 1).Value = rank
        ws.Cells(rowNum, 2).Value = prodData(0)
        ws.Cells(rowNum, 3).Value = productKey
        ws.Cells(rowNum, 4).Value = prodData(1)
        ws.Cells(rowNum, 5).Value = prodData(2)
        ws.Cells(rowNum, 6).Value = prodData(3)
        ws.Cells(rowNum, 7).Value = prodData(4)
        
        ws.Range("A" & rowNum & ":G" & rowNum).Interior.ColorIndex = xlNone
        ws.Range("A" & rowNum & ":G" & rowNum).HorizontalAlignment = xlCenter
        ws.Range("A" & rowNum & ":G" & rowNum).VerticalAlignment = xlCenter
        ws.Range("A" & rowNum & ":G" & rowNum).Font.Size = 11
        ws.Range("A" & rowNum & ":G" & rowNum).WrapText = False
        ws.Range("A" & rowNum & ":G" & rowNum).ShrinkToFit = False
        ws.Range("A" & rowNum & ":G" & rowNum).Borders.LineStyle = xlContinuous
        ws.Range("A" & rowNum & ":G" & rowNum).Borders.Weight = xlThin
        
        ws.Cells(rowNum, 4).NumberFormat = "#,##0"
        ws.Cells(rowNum, 5).NumberFormat = "#,##0"
        ws.Cells(rowNum, 6).NumberFormat = "#,##0"
        ws.Cells(rowNum, 7).NumberFormat = "#,##0"
        
        rank = rank + 1
        rowNum = rowNum + 1
    Next productKey
    
    ws.Cells.EntireColumn.AutoFit
    ws.Cells.EntireRow.AutoFit
    
    ws.Columns("A:G").AutoFit
    Dim colNum As Long
    For colNum = 1 To 7
        ws.Columns(colNum).ColumnWidth = ws.Columns(colNum).ColumnWidth + 3
    Next colNum
    
    ws.Rows.AutoFit
    
    ws.Range("A1:G" & rowNum).HorizontalAlignment = xlCenter
    ws.Range("A1:G" & rowNum).VerticalAlignment = xlCenter
    ws.Range("A1:G" & rowNum).WrapText = False
    ws.Range("A1:G" & rowNum).ShrinkToFit = False
    
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

Private Sub CollectProductDataWithID(dict As Object, ws As Worksheet, categoryIndex As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    Dim i As Long
    For i = 2 To lastRow
        Dim productName As String, itemID As String, qty As Long
        productName = Trim(CStr(ws.Cells(i, 4).Value))
        itemID = Trim(CStr(ws.Cells(i, 3).Value))
        
        On Error Resume Next
        qty = CLng(ws.Cells(i, 5).Value)
        On Error GoTo 0
        
        If Len(productName) > 0 And qty > 0 And UCase(productName) <> "TOTAL" Then
            If Not dict.Exists(productName) Then
                Dim newArr(0 To 4) As Variant
                newArr(0) = itemID
                newArr(1) = 0: newArr(2) = 0: newArr(3) = 0: newArr(4) = 0
                dict.Add productName, newArr
            End If
            
            Dim arr As Variant
            arr = dict(productName)
            arr(categoryIndex + 1) = arr(categoryIndex + 1) + qty
            dict(productName) = arr
        End If
    Next i
End Sub

Private Function SortProductsByTotal(sourceDict As Object) As Object
    Dim sortedDict As Object
    Set sortedDict = CreateObject("Scripting.Dictionary")
    
    If sourceDict.Count = 0 Then
        Set SortProductsByTotal = sortedDict
        Exit Function
    End If
    
    Dim keys() As Variant, totals() As Long
    ReDim keys(0 To sourceDict.Count - 1)
    ReDim totals(0 To sourceDict.Count - 1)
    
    Dim i As Long, k As Variant
    i = 0
    For Each k In sourceDict.Keys
        Dim arr As Variant
        arr = sourceDict(k)
        keys(i) = k
        totals(i) = arr(1) + arr(2) + arr(3) + arr(4)
        i = i + 1
    Next k
    
    Dim j As Long, tempKey As Variant, tempTotal As Long
    For i = 0 To UBound(totals) - 1
        For j = i + 1 To UBound(totals)
            If totals(i) < totals(j) Then
                tempTotal = totals(i)
                totals(i) = totals(j)
                totals(j) = tempTotal
                
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
    
    For i = 0 To UBound(keys)
        sortedDict.Add keys(i), sourceDict(keys(i))
    Next i
    
    Set SortProductsByTotal = sortedDict
End Function

Private Sub AddRSMPerformance(ws As Worksheet, wsE As Worksheet, ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, sr As Long)
    ws.Cells.Font.Name = "Arial"
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lr As Long, i As Long
    
    lr = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            Dim rsm As String, qty As Long
            rsm = Trim(CStr(wsE.Cells(i, 2).Value))
            On Error Resume Next
            qty = CLng(wsE.Cells(i, 5).Value)
            On Error GoTo 0
            
            If Len(rsm) > 0 And UCase(wsE.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0)
                End If
                Dim arr As Variant
                arr = dict(rsm)
                arr(0) = arr(0) + 1
                arr(1) = arr(1) + qty
                dict(rsm) = arr
            End If
        Next i
    End If
    
    lr = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            rsm = Trim(CStr(ws1.Cells(i, 2).Value))
            On Error Resume Next
            qty = CLng(ws1.Cells(i, 5).Value)
            On Error GoTo 0
            
            If Len(rsm) > 0 And UCase(ws1.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0)
                End If
                arr = dict(rsm)
                arr(2) = arr(2) + 1
                arr(3) = arr(3) + qty
                dict(rsm) = arr
            End If
        Next i
    End If
    
    lr = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            rsm = Trim(CStr(ws2.Cells(i, 2).Value))
            On Error Resume Next
            qty = CLng(ws2.Cells(i, 5).Value)
            On Error GoTo 0
            
            If Len(rsm) > 0 And UCase(ws2.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0)
                End If
                arr = dict(rsm)
                arr(4) = arr(4) + 1
                arr(5) = arr(5) + qty
                dict(rsm) = arr
            End If
        Next i
    End If
    
    lr = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row
    If lr > 1 Then
        For i = 2 To lr
            rsm = Trim(CStr(ws3.Cells(i, 2).Value))
            On Error Resume Next
            qty = CLng(ws3.Cells(i, 5).Value)
            On Error GoTo 0
            
            If Len(rsm) > 0 And UCase(ws3.Cells(i, 1).Value) <> "TOTAL" Then
                If Not dict.Exists(rsm) Then
                    dict.Add rsm, Array(0, 0, 0, 0, 0, 0, 0, 0)
                End If
                arr = dict(rsm)
                arr(6) = arr(6) + 1
                arr(7) = arr(7) + qty
                dict(rsm) = arr
            End If
        Next i
    End If
    
    Dim sortedKeys() As String, sortedVals() As Long
    ReDim sortedKeys(0 To dict.Count - 1)
    ReDim sortedVals(0 To dict.Count - 1)
    
    Dim j As Long, k As Variant
    j = 0
    For Each k In dict.Keys
        sortedKeys(j) = CStr(k)
        arr = dict(k)
        sortedVals(j) = arr(1) + arr(3) + arr(5) + arr(7)
        j = j + 1
    Next k
    
    Dim m As Long, n As Long, tempK As String, tempV As Long
    For m = 0 To UBound(sortedVals) - 1
        For n = m + 1 To UBound(sortedVals)
            If sortedVals(m) < sortedVals(n) Then
                tempV = sortedVals(m)
                sortedVals(m) = sortedVals(n)
                sortedVals(n) = tempV
                
                tempK = sortedKeys(m)
                sortedKeys(m) = sortedKeys(n)
                sortedKeys(n) = tempK
            End If
        Next n
    Next m
    
    With ws
        .Range("A" & sr & ":F" & sr).Merge
        .Cells(sr, 1).Value = ChrW(&H62A) & ChrW(&H642) & ChrW(&H64A) & ChrW(&H64A) & ChrW(&H645) & " " & ChrW(&H623) & ChrW(&H62F) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H645) & ChrW(&H62F) & ChrW(&H631) & ChrW(&H627) & ChrW(&H621) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & " (RSMs)"
        .Cells(sr, 1).Font.Bold = True
        .Cells(sr, 1).Interior.Color = RGB(68, 114, 196)
        .Cells(sr, 1).Font.Color = RGB(255, 255, 255)
        .Cells(sr, 1).HorizontalAlignment = xlCenter
        .Cells(sr, 1).VerticalAlignment = xlCenter
        .Cells(sr, 1).WrapText = False
        .Cells(sr, 1).ShrinkToFit = False
        .Cells(sr, 1).Borders.LineStyle = xlContinuous
        .Cells(sr, 1).Borders.Weight = xlMedium
        .Cells(sr, 1).Borders.Color = RGB(0, 0, 0)
        sr = sr + 1
        
        .Cells(sr, 1).Value = ChrW(&H645) & ChrW(&H62F) & ChrW(&H6CC) & ChrW(&H631) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H642) & ChrW(&H644) & ChrW(&H64A) & ChrW(&H645) & ChrW(&H64A)
        .Cells(sr, 2).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H645) & ChrW(&H646) & ChrW(&H62A) & ChrW(&H647) & ChrW(&H6CC) & " " & ChrW(&H627) & ChrW(&H644) & ChrW(&H635) & ChrW(&H644) & ChrW(&H627) & ChrW(&H62D) & ChrW(&H64A) & ChrW(&H629)
        .Cells(sr, 3).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(sr, 4).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H634) & ChrW(&H647) & ChrW(&H631) & ChrW(&H6CC) & ChrW(&H646)
        .Cells(sr, 5).Value = ChrW(&H6A9) & ChrW(&H645) & ChrW(&H6CC) & ChrW(&H629) & " " & ChrW(&H628) & ChrW(&H642) & ChrW(&H6CC) & " " & ChrW(&H663) & " " & ChrW(&H623) & ChrW(&H634) & ChrW(&H647) & ChrW(&H631)
        .Cells(sr, 6).Value = ChrW(&H627) & ChrW(&H644) & ChrW(&H625) & ChrW(&H62C) & ChrW(&H645) & ChrW(&H627) & ChrW(&H644) & ChrW(&H64A)
        
        .Range("A" & sr & ":F" & sr).Font.Bold = True
        .Range("A" & sr & ":F" & sr).Interior.Color = RGB(217, 217, 217)
        .Range("A" & sr & ":F" & sr).HorizontalAlignment = xlCenter
        .Range("A" & sr & ":F" & sr).VerticalAlignment = xlCenter
        .Range("A" & sr & ":F" & sr).WrapText = False
        .Range("A" & sr & ":F" & sr).ShrinkToFit = False
        .Range("A" & sr & ":F" & sr).Borders.LineStyle = xlContinuous
        .Range("A" & sr & ":F" & sr).Borders.Weight = xlMedium
        .Range("A" & sr & ":F" & sr).Borders.Color = RGB(0, 0, 0)
        .Cells(sr, 2).Interior.Color = RGB(255, 102, 102)
        .Cells(sr, 3).Interior.Color = RGB(255, 153, 153)
        .Cells(sr, 4).Interior.Color = RGB(255, 192, 128)
        .Cells(sr, 5).Interior.Color = RGB(255, 230, 153)
        .Cells(sr, 6).Interior.Color = RGB(255, 255, 0)
        .Cells(sr, 6).Font.Color = RGB(0, 0, 0)
        sr = sr + 1
        
        For j = 0 To UBound(sortedKeys)
            arr = dict(sortedKeys(j))
            Dim rsmTotal As Long
            rsmTotal = arr(1) + arr(3) + arr(5) + arr(7)
            
            .Cells(sr, 1).Value = sortedKeys(j)
            .Cells(sr, 2).Value = arr(1)
            .Cells(sr, 3).Value = arr(3)
            .Cells(sr, 4).Value = arr(5)
            .Cells(sr, 5).Value = arr(7)
            .Cells(sr, 6).Value = rsmTotal
            
            If arr(1) > 10000 Then
                .Range("A" & sr & ":E" & sr).Interior.Color = RGB(255, 102, 102)
            ElseIf arr(1) > 5000 Then
                .Range("A" & sr & ":E" & sr).Interior.Color = RGB(255, 192, 128)
            End If
            
            .Range("A" & sr & ":F" & sr).HorizontalAlignment = xlCenter
            .Range("A" & sr & ":F" & sr).VerticalAlignment = xlCenter
            .Range("A" & sr & ":F" & sr).WrapText = False
            .Range("A" & sr & ":F" & sr).ShrinkToFit = False
            .Range("A" & sr & ":F" & sr).Borders.LineStyle = xlContinuous
            .Range("A" & sr & ":F" & sr).Borders.Weight = xlThin
            .Range("A" & sr & ":F" & sr).Borders.Color = RGB(0, 0, 0)
            .Cells(sr, 2).NumberFormat = "#,##0"
            .Cells(sr, 3).NumberFormat = "#,##0"
            .Cells(sr, 4).NumberFormat = "#,##0"
            .Cells(sr, 5).NumberFormat = "#,##0"
            .Cells(sr, 6).NumberFormat = "#,##0"
            sr = sr + 1
        Next j
        
        Dim grandTotal As Long
        grandTotal = 0
        For Each k In dict.Keys
            arr = dict(k)
            grandTotal = grandTotal + arr(1) + arr(3) + arr(5) + arr(7)
        Next k
        
        .Cells(sr, 6).Value = grandTotal
        .Range("F" & sr).Font.Bold = True
        .Range("F" & sr).Interior.Color = RGB(146, 208, 80)
        .Range("F" & sr).HorizontalAlignment = xlCenter
        .Range("F" & sr).VerticalAlignment = xlCenter
        .Range("F" & sr).WrapText = False
        .Range("F" & sr).ShrinkToFit = False
        .Range("F" & sr).Borders.LineStyle = xlContinuous
        .Range("F" & sr).Borders.Weight = xlMedium
        .Range("F" & sr).Borders.Color = RGB(0, 0, 0)
        .Cells(sr, 6).NumberFormat = "#,##0"
        
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    End With
End Sub
