# ethical-hacking-report-project
This is a temp report for the student. The idea is to automate daily findings. No screenshots or exploits will live in the repo.

# 🛡️ Automated Pentest Daily Report Builder

A fully local, VBA-driven system for generating structured daily wrap-up reports during a month-long ethical hacking exam. Built with Excel + Word automation. No internet required. No cloud integration.

---
# Start of Project
<img width="758" height="503" alt="version-1" src="https://github.com/user-attachments/assets/649a0bb4-6659-436d-9e65-1e5d6461e456" />

| A               | B                   |
| --------------- | ------------------- |
| Screenshot File | Caption             |
| `01_enum.png`   | `Nmap scan result`  |
| `02_shell.png`  | `Meterpreter shell` |
| ...             | ...                 |

# The above is working - and will evolve. The pdf macro needs testing - Push report to a pdf.
```vba
    ' === Export to PDF ===
`ImportScreenshotsFromExcel` => Is the first working procedure. Dial in on a file naming scheme for greenshot. => Push to webDav or home sharepoint server and PowerAutomate
    Dim pdfPath As String
    pdfPath = ThisDocument.Path & "\" & _
              Replace(ThisDocument.Name, ".docm", "_" & Format(Now, "yyyy-mm-dd_hhmmss") & ".pdf")

    ThisDocument.ExportAsFixedFormat OutputFileName:=pdfPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks

    MsgBox " PDF saved to: " & pdfPath, vbInformation

```
