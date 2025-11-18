Sub GenerateDailyReport()
    Dim reportDate As String
    Dim evidenceFolder As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim templatePath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim contentRange As Object
    Dim shot As String, fullPath As String
    Dim k As Integer, screenshotCount As Integer
    
    ' --- Prompt for Date ---
    reportDate = InputBox("Enter the date to generate report for (YYYY-MM-DD):", "Daily Report")
    If reportDate = "" Then Exit Sub

    ' --- Prompt for Folder ---
    evidenceFolder = PickEvidenceFolder()
    If evidenceFolder = "" Then
        MsgBox "No evidence folder selected.", vbExclamation
        Exit Sub
    End If

    ' --- Validate Worksheet ---
    Set ws = ThisWorkbook.Sheets("DailyLog")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' --- Open Word Template ---
    templatePath = ThisWorkbook.Path & "\DailyReport_Template.docm"
    If Dir(templatePath) = "" Then
        MsgBox "Template not found: " & templatePath, vbCritical
        Exit Sub
    End If

    On Error GoTo WordError
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Open(templatePath)
    wordApp.Visible = True
    On Error GoTo 0

    Set contentRange = wordDoc.Content
    contentRange.Collapse Direction:=0 ' wdCollapseStart

    ' --- Header ---
    With contentRange
        .InsertAfter "🔒 Daily Wrap-Up Report" & vbCrLf
        .InsertAfter "📅 Date: " & reportDate & vbCrLf & vbCrLf
        .InsertParagraphAfter
    End With

    screenshotCount = 1

    ' --- Loop Through Findings ---
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) = reportDate Then
            With contentRange
                ' Finding Title and Severity
                .InsertAfter ws.Cells(i, 2).Value & ". " & ws.Cells(i, 3).Value & _
                    " [" & UCase(ws.Cells(i, 4).Value) & "]" & vbCrLf

                ' Host, Description, Notes
                .InsertAfter "🔹 Host: " & ws.Cells(i, 6).Value & vbCrLf
                .InsertAfter "📝 Description: " & ws.Cells(i, 5).Value & vbCrLf
                .InsertAfter "📓 Notes: " & ws.Cells(i, 10).Value & vbCrLf & vbCrLf

                ' Screenshots
                For k = 7 To 9
                    shot = Trim(ws.Cells(i, k).Value)
                    If shot <> "" Then
                        fullPath = evidenceFolder & "\" & shot
                        If Dir(fullPath) <> "" Then
                            .InsertParagraphAfter
                            .InsertAfter "🖼 Figure " & screenshotCount & ": " & shot & vbCrLf
                            .InsertParagraphAfter
                            .InlineShapes.AddPicture FileName:=fullPath, _
                                LinkToFile:=False, SaveWithDocument:=True
                            .InsertParagraphAfter
                            screenshotCount = screenshotCount + 1
                        Else
                            .InsertAfter "[⚠️ Missing Screenshot: " & shot & "]" & vbCrLf
                        End If
                        .InsertParagraphAfter
                    End If
                Next k

                .InsertParagraphAfter
                .InsertAfter String(50, "-") & vbCrLf & vbCrLf
            End With
        End If
    Next i

    MsgBox "✅ Daily report created successfully!", vbInformation
    Exit Sub

WordError:
    MsgBox "❌ Error opening Word template. Ensure Word is installed and the file exists.", vbCritical
    Exit Sub
End Sub
