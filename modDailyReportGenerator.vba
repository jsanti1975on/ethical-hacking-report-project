Option Explicit

' Prompt user to select an evidence folder
Function PickEvidenceFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Select the Evidence Folder for This Day"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickEvidenceFolder = ""
        Else
            PickEvidenceFolder = .SelectedItems(1)
        End If
    End With
End Function
