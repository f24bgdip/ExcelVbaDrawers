Attribute VB_Name = "FileIO_ReadWriteBinary"
Option Explicit

Sub FileIO_Binary()
    ' Input
    Dim inputFileName As String
    Dim inputFn As Long

    ' Set the input file name, and
    ' set the file number with free number.
    inputFileName = "C:\example_input.jpg"
    inputFn = FreeFile

    ' Read the input file as binary.
    Dim buffer() As Byte
    
    Open inputFileName For Binary As #inputFn
        ' Reserve buffers, and load binary to them.
        ReDim buffer(LOF(inputFn))
        Get #inputFn, , buffer
    Close #inputFn


    ' Output
    Dim outputFileName As String
    Dim outputFn As Long
    
    ' Set the output file name, and
    ' set the file number with free number.
    outputFileName = "C:\example_output.jpg"
    outputFn = FreeFile

    ' Write the output file by the input file.
    Dim i As Long
    
    Open outputFileName For Binary As #outputFn
        For i = 0 To UBound(buffer) - 1
            Put #outputFn, , buffer(i)
        Next
    Close #outputFn

End Sub
