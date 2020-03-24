Attribute VB_Name = "CodingConventions"
Option Explicit

' Visual Basic Coding Conventions
' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/coding-conventions
' Naming Conventions
' For information about naming guidelines, see Naming Guidelines topic.
' Do not use "My" or "my" as part of a variable name. This practice creates confusion with the My objects.
' You do not have to change the names of objects in auto-generated code to make them fit the guidelines.

' Visual Basic naming rules
' https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules

' 
' 
' 
Sub ProcedureCall()
    ' Variable declaration
    Dim bok As Workbook
    Dim shtIn As Worksheet, shtOut As Worksheet
    Dim rng As Range
    Dim vnt As Variant
    Dim str As String
    
    ' Regist Exception processing
    On Error GoTo ErrorHandler

    ' Stop updationg screen
    Application.ScreenUpdating = False
    
    ' Initialization
    Set bok = Workbooks("sample.xlsx")
    Set shtIn = bok.Worksheets("sheet")
    

    ' Input
    str = "I"
    ' Transformation
    ' str is sent as ByRef
    Call Stub(str)
    ' Output
    MsgBox str
    

    ' Input
    str = "I"
    ' Transformation
    ' str is sent as ByRef
    Stub str
    ' Output
    MsgBox str
    

    ' Input
    str = "I"
    ' Transformation
    ' str is sent as ByVal
    Stub (str)
    ' Output
    MsgBox str
   
    
' Error handling
ErrorHandler:

    ' Restart updationg screen
    Application.ScreenUpdating = True

End Sub


Sub Stub(ByRef str As String)
    str = "You"
End Sub
