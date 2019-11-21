Attribute VB_Name = "Cells_RowCol2Text"
Option Explicit

'
' The row and column values are converted to text address.
'
Function RowCol2Text(row As Long, col As Long) As String
    RowCol2TextAddress = Cells(row, col).Address
End Function

Sub Sample()
    MsgBox (RowCol2TextAddress(1, 1))
End Sub


