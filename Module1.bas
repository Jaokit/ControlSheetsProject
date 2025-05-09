Attribute VB_Name = "Module1"
Option Explicit

Function CleanFileName(str As String) As String
    Dim invalidChars As Variant, i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        str = Replace(str, invalidChars(i), "_")
    Next i
    CleanFileName = str
End Function

Sub ExportBoxForms()
    Dim wsData As Worksheet, wsTemplate As Worksheet, wsBoxDetails As Worksheet, wsMatrix As Worksheet
    Dim lastRow As Long, i As Long
    Dim customer As String, partNo As String, customerLot As String, opmLot As String
    Dim theDate As String, tagIssue As String, pieces As Variant
    Dim folderPath As String, formCount As Long, groupIndex As Long
    Dim newWb As Workbook, newWs As Worksheet, boxRow As Long, lastBoxRow As Long
    Dim fileName As String, rowOffset As Long
    Dim boxNo As Long, currentLot As String, currentOPM As String
    Dim formIndex As Long
    Dim totalBoxes As Long

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first.", vbExclamation
        Exit Sub
    End If

    Set wsData = ThisWorkbook.Sheets("SourceData")
    Set wsTemplate = ThisWorkbook.Sheets("BoxFormTemplate")
    Set wsBoxDetails = ThisWorkbook.Sheets("BoxDetails")
    Set wsMatrix = ThisWorkbook.Sheets("InspectionMatrix")

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    folderPath = ThisWorkbook.Path & "\ControlBox\"
    'If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    lastBoxRow = wsBoxDetails.Cells(wsBoxDetails.Rows.Count, 1).End(xlUp).Row

    totalBoxes = wsData.Range("F2").Value
    formIndex = 0

    For i = 2 To lastRow
        customer = wsData.Cells(i, 1).Value
        partNo = wsData.Cells(i, 2).Value
        customerLot = wsData.Cells(i, 3).Value
        opmLot = wsData.Cells(i, 4).Value
        theDate = wsData.Cells(i, 5).Value
        tagIssue = wsData.Cells(i, 8).Value

        formCount = 0
        groupIndex = 1
        Set newWb = Workbooks.Add(xlWBATWorksheet)
        Set newWs = newWb.Sheets(1)

        With newWs.PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .TopMargin = Application.InchesToPoints(0.3)
            .BottomMargin = Application.InchesToPoints(0.3)
            .LeftMargin = Application.InchesToPoints(0.3)
            .RightMargin = Application.InchesToPoints(0.3)
        End With

        For boxRow = 2 To lastBoxRow
            currentLot = wsBoxDetails.Cells(boxRow, 1).Value
            currentOPM = wsBoxDetails.Cells(boxRow, 2).Value

            If currentLot = customerLot And currentOPM = opmLot Then
                boxNo = wsBoxDetails.Cells(boxRow, 3).Value
                pieces = wsBoxDetails.Cells(boxRow, 4).Value

                rowOffset = (formCount Mod 4) * 13
                If rowOffset = 0 And formCount > 0 Then
                    fileName = folderPath & CleanFileName(customerLot) & "_Group" & groupIndex & ".xlsx"
                    Application.DisplayAlerts = False
                    newWb.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
                    newWb.Close SaveChanges:=False
                    Application.DisplayAlerts = True

                    groupIndex = groupIndex + 1
                    Set newWb = Workbooks.Add(xlWBATWorksheet)
                    Set newWs = newWb.Sheets(1)
                    With newWs.PageSetup
                        .Orientation = xlPortrait
                        .PaperSize = xlPaperA4
                        .Zoom = False
                        .FitToPagesWide = 1
                        .FitToPagesTall = False
                        .TopMargin = Application.InchesToPoints(0.3)
                        .BottomMargin = Application.InchesToPoints(0.3)
                        .LeftMargin = Application.InchesToPoints(0.3)
                        .RightMargin = Application.InchesToPoints(0.3)
                    End With
                End If

                formIndex = formIndex + 1
                Call FillBoxForm(newWs, rowOffset, customer, partNo, tagIssue, theDate, _
                                 customerLot, opmLot, pieces, formIndex, totalBoxes, wsMatrix)
                formCount = formCount + 1
            End If
        Next boxRow

        If formCount > 0 Then
            fileName = folderPath & CleanFileName(customerLot) & "_Group" & groupIndex & ".xlsx"
            Application.DisplayAlerts = False
            newWb.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
            newWb.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    Next i

    MsgBox "Success go to Folder ControlBox", vbInformation
End Sub

Sub QuickSortNumeric(arr As Variant, first As Long, last As Long)
    Dim low As Long, high As Long, midValue As Variant, temp As Variant
    low = first
    high = last
    midValue = arr((first + last) \ 2)

    Do While low <= high
        Do While arr(low) < midValue
            low = low + 1
        Loop
        Do While arr(high) > midValue
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortNumeric arr, first, high
    If low < last Then QuickSortNumeric arr, low, last
End Sub

Sub FillBoxForm(ws As Worksheet, rowOffset As Long, _
                customer As String, partNo As String, tagIssue As String, theDate As String, _
                customerLot As String, opmLot As String, pieces As Variant, _
                boxIndex As Long, totalBoxes As Long, wsMatrix As Worksheet)

    Dim templateWs As Worksheet
    Set templateWs = ThisWorkbook.Sheets("BoxFormTemplate")

    templateWs.Range("A1:M12").Copy
    ws.Cells(1 + rowOffset, 1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    ws.Paste
    ws.Columns("A").ColumnWidth = 10
    ws.Rows("2").RowHeight = 20

    Application.CutCopyMode = False
    ws.Range(ws.Cells(1 + rowOffset, 1), ws.Cells(12 + rowOffset, 13)).Font.Name = "Calibri"

    With ws
        .Cells(4 + rowOffset, 2).Value = customer
        .Cells(5 + rowOffset, 2).Value = partNo
        .Cells(6 + rowOffset, 2).Value = tagIssue
        .Cells(4 + rowOffset, 7).NumberFormat = "dd/mm/yy"
            If IsDate(theDate) Then
                .Cells(4 + rowOffset, 7).Value = CDate(theDate)
            Else
                .Cells(4 + rowOffset, 7).Value = theDate
            End If
        .Cells(5 + rowOffset, 7).Value = customerLot
        .Cells(6 + rowOffset, 7).Value = opmLot
        .Cells(2 + rowOffset, 3).Value = pieces
        .Cells(4 + rowOffset, 12).Value = "[" & boxIndex & " of " & totalBoxes & "]"

        Dim matrixRow As Long, matrixLastCol As Long, writeCol As Long, i As Long
        writeCol = 2

        For i = 2 To wsMatrix.Cells(wsMatrix.Rows.Count, 1).End(xlUp).Row
            If wsMatrix.Cells(i, 1).Value = partNo Then
                matrixRow = i
                Exit For
            End If
        Next i

        If matrixRow = 0 Then
            MsgBox "Part No '" & partNo & "' not found in InspectionMatrix!", vbCritical
            Exit Sub
        End If

        matrixLastCol = wsMatrix.Cells(1, wsMatrix.Columns.Count).End(xlToLeft).Column

        Dim stepDict As Object
        Set stepDict = CreateObject("Scripting.Dictionary")

        For i = 2 To matrixLastCol
            If Trim(wsMatrix.Cells(matrixRow, i).Value) <> "" And IsNumeric(wsMatrix.Cells(matrixRow, i).Value) Then
                stepDict(wsMatrix.Cells(matrixRow, i).Value) = wsMatrix.Cells(1, i).Value
            End If
        Next i

        If stepDict.Count > 0 Then
            Dim sortedKeys As Variant
            sortedKeys = stepDict.Keys
            Call QuickSortNumeric(sortedKeys, LBound(sortedKeys), UBound(sortedKeys))

            Dim key
            For Each key In sortedKeys
                .Cells(8 + rowOffset, writeCol).Value = stepDict(key)
                writeCol = writeCol + 1
            Next key
        End If
    End With
End Sub


