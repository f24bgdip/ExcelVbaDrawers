Attribute VB_Name = "Array_and_Split_with_OptionBase"
Option Explicit
Option Base 0

'
' Array function is affected by Option Base,
' but Split function is not.
Sub OptionBaseSample()
    ' Variable declaration
    Dim ArrayData As Variant
    Dim SplitData As Variant

    ArrayData = Array("0", "1", "2")
    SplitData = Split("0,1,2", ",")
    
    MsgBox "Array ÅF " & ArrayData(1) & vbCrLf _
        & "Split  ÅF " & SplitData(1)

End Sub
