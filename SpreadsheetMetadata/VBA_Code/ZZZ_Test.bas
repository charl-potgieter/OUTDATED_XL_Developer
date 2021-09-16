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
    

    Set a = New zLIB_ListStorage
    a.CreateStorageFromPowerQuery ActiveWorkbook, "MyTest", "Table1"
    Set a = Nothing


End Sub


Sub TestCreateStorage()

    Dim a As zLIB_ListStorage

    Set a = New zLIB_ListStorage
    a.CreateStorage ActiveWorkbook, "TestStore3", Array("Header1", "Header2", "Header3", "Header4")

    Set a = Nothing

End Sub
