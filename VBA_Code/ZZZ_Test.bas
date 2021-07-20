Attribute VB_Name = "ZZZ_Test"
Option Explicit

Sub testProperties()


    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("aah").Delete
    On Error GoTo 0

    ActiveWorkbook.CustomDocumentProperties.Add "aah", False, msoPropertyTypeBoolean, True
    
    Debug.Print (ActiveWorkbook.CustomDocumentProperties("aah"))


End Sub


