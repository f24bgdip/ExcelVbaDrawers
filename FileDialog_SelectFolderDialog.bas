Attribute VB_Name = "FileDialog_SelectFolderDialog"

Public Function SelectFolderDialog(Optional ByRef strTitle As String) As String
    'SelectFolderDialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = strTitle
        If .Show = -1 Then
            SelectFolderDialog = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
        Else
            SelectFolderDialog = ""
        End If
    End With
End Function
