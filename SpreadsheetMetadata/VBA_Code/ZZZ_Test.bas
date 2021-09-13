Attribute VB_Name = "ZZZ_Test"
Option Explicit

Sub test()

    Dim a
    Dim lo As ListObject
    
    'a = Application.Evaluate("=UNIQUE(tbl_ListObjectFormats[SheetName])")

    Set lo = ActiveSheet.ListObjects(1)


    a = WorksheetFunction.Unique(lo.ListColumns("SheetName").DataBodyRange)

End Sub




Sub test2()

    Dim a As zLIB_ListStorage
    
    a.


End Sub
