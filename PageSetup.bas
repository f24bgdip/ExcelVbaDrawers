Attribute VB_Name = "PageSetup"

Sub Set_PageSetup()
    Dim ws As Variant

    For Each ws In Worksheets
        With ws.PageSetup
            .LeftHeader = "[]"
            .CenterHeader = ""
            .RightHeader = "&F"
            .LeftFooter = ""
            .CenterFooter = "&A - page &P/&N"
            .RightFooter = ""
        End With
    Next
End Sub
