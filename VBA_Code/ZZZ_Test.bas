Attribute VB_Name = "ZZZ_Test"
Option Explicit

Sub testFileNames()
    
    Dim FileItems() As Scripting.File
    Dim i As Integer
    Dim FolderPath As String
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    FolderPath = "C:\Users\charl\Dropbox\Dropbox_Charl\Computer_Technical\Programming_GitHub\ExcelVbaCodeLibrary\VBA_Code"
    
    FileItemsInFolder FolderPath, False, FileItems
    
    For i = LBound(FileItems) To UBound(FileItems)
        Debug.Print (fso.GetBaseName(FileItems(i).Name))
    Next i
    
End Sub


