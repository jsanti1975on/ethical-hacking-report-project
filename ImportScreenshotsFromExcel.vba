Sub ImportScreenshotsFromExcel()
    Dim fd As FileDialog
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim row As Long
    Dim imgPath As String, captionText As String
    Dim figureCount As Long
    Dim docPath As String
    Dim pic As InlineShape

    ' Prompt user to select Excel workbook
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the Excel file with screenshot data"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        If .Show <> -1 Then Exit Sub
        docPath = .SelectedItems(1)
    End With

    ' Launch Excel in background
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    If excelApp Is Nothing Then Set excelApp = CreateObject("Excel.Application")
    On Error GoTo 0

    excelApp.Visible = False
    Set wb = excelApp.Workbooks.Open(docPath)
    Set ws = wb.Sheets("Screenshots")

    figureCount = 1
    row = 2 ' Assuming row 1 is header

    Do While ws.Cells(row, 1).Value <> ""
        imgPath = wb.Path & "\" & ws.Cells(row, 1).Value
        captionText = ws.Cells(row, 2).Value

        ' Validate file exists
        If Dir(imgPath) <> "" Then
            ' Insert image
            Set pic = Selection.InlineShapes.AddPicture(FileName:=imgPath, _
                     LinkToFile:=False, SaveWithDocument:=True)

            ' Resize to max page width (optional)
            With pic
                .LockAspectRatio = True
                If .Width > InchesToPoints(6.5) Then
                    .Width = InchesToPoints(6.5)
                End If
            End With

            ' Insert caption
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.TypeParagraph
            Selection.Font.Italic = True
            Selection.Font.Size = 10
            Selection.TypeText "Figure " & figureCount & ": " & captionText
            Selection.TypeParagraph
            Selection.TypeParagraph
            figureCount = figureCount + 1
        Else
            ' Insert missing file notice
            Selection.TypeText "[Missing file: " & imgPath & "]"
            Selection.TypeParagraph
        End If

        row = row + 1
    Loop

    wb.Close SaveChanges:=False
    excelApp.Quit
    Set excelApp = Nothing

    MsgBox "Screenshots imported successfully!", vbInformation
End Sub
