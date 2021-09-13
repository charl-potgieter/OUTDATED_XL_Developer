Attribute VB_Name = "m030_General"
Option Explicit

Function RangeOfStoredData(ByVal ItemDescription As String) As Range

    Dim lo As ListObject
    Dim EvaluationFormula As String
    Dim ItemIndex As Integer
    
    Set lo = ThisWorkbook.Sheets("XL_Developer").ListObjects("tbl_Data")
    On Error Resume Next
    ItemIndex = WorksheetFunction.Match(ItemDescription, lo.ListColumns("Item").DataBodyRange, 0)
    
    If Err.Number <> 0 Then
        Set RangeOfStoredData = Nothing
    Else
        Set RangeOfStoredData = lo.ListColumns("Value").DataBodyRange.Cells(ItemIndex)
    End If

End Function



Function StoredDataValue(ByVal ItemDescription As String)

    Dim lo As ListObject
    Dim EvaluationFormula As String
    Dim ItemIndex As Integer
    
    Set lo = ThisWorkbook.Sheets("XL_Developer").ListObjects("tbl_Data")
    On Error Resume Next
    ItemIndex = WorksheetFunction.Match(ItemDescription, lo.ListColumns("Item").DataBodyRange, 0)
    
    If Err.Number <> 0 Then
        StoredDataValue = "NULL"
    Else
        StoredDataValue = lo.ListColumns("Value").DataBodyRange.Cells(ItemIndex).Value
    End If

End Function
