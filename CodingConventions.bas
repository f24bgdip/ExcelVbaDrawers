Attribute VB_Name = "CodingConventions"
Option Explicit

' Visual Basic Coding Conventions
' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/coding-conventions
' Visual Basic naming rules
' https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules


' 
' 
' 
Sub ProcedureCall()
    Dim str As String
    
    ' str is sent as ByRef
    str = "First"
    Call Sample2(str)
    MsgBox str
    
    ' str is sent as ByRef
    str = "First"
    Sample2 str
    MsgBox str
    
    ' str is sent as ByVal
    str = "First"
    Sample2 (str)
    MsgBox str
End Sub

Sub Sample2(ByRef str As String)
    str = "Second"
End Sub
