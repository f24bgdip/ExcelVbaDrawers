Attribute VB_Name = "Search_InStr"
Option Explicit

Sub Search_InStr()
    ' input
    Dim inputFileName As String
    Dim inputFn As Long

    ' Set the input file name, and
    ' set the file number with free number.
    inputFileName = "D:\dummy.txt"
    inputFn = FreeFile

    ' Read the input file as binary.
    ' Reserve buffers, and load binary to them.
    Dim buffer() As Byte

    Open inputFileName For Binary As #inputFn
        ReDim buffer(LOF(inputFn))
        Get #inputFn, , buffer
    Close #inputFn

    ' Search a string in the input file.
    Dim str As String
    Dim ret As Variant

    str = "b"
    ret = InStr(1, buffer, str, vbBinaryCompare)
    
    If (str = buffer()) Then
        Debug.Print str
    End If

End Sub
