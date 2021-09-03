Attribute VB_Name = "EntryPoints"
Option Explicit


Sub ExportActiveWorkbookVbaCode(Optional control As IRibbonControl)

    Dim wkb As Workbook
    Dim sVbaCodePath As String

    Set wkb = ActiveWorkbook
    sVbaCodePath = wkb.Path & Application.PathSeparator & "VBA_Code"
    
    'Delete any existing (which may be outdated) code files in folder
    On Error Resume Next
    Kill sVbaCodePath & Application.PathSeparator & "*.*"
    On Error GoTo 0
    

    If Not FolderExists(sVbaCodePath) Then
        CreateFolder sVbaCodePath
    End If

    ExportVBAModules wkb, sVbaCodePath
    MsgBox ("VBA code exported")

End Sub



Sub SaveStandardCodeLibraryAndImportIntoCurrentWorkbook() '(Optional control As IRibbonControl)

    Dim sVbaCodePath As String
    Dim ModuleNames() As String
    Dim i As Long

    StandardEntry
      
    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox ("Cannot apply this action in when " & ThisWorkbook.Name & _
                " is the active workbook")
    Else
        ThisWorkbook.Save
        sVbaCodePath = ThisWorkbook.Path & Application.PathSeparator & "zLIB_VBA_Code"
        If Not FolderExists(sVbaCodePath) Then
            CreateFolder sVbaCodePath
        End If
        'Delete any previous files in this folder
        On Error Resume Next
        Kill sVbaCodePath & Application.PathSeparator & "*.*"
        On Error GoTo 0
        
        ExportVBAModules ThisWorkbook, sVbaCodePath
        
        'Delete any current code lib files in active workbook
        GetBaseFileNamesInFolder sVbaCodePath, ModuleNames
        For i = LBound(ModuleNames) To UBound(ModuleNames)
            DeleteModule ActiveWorkbook, ModuleNames(i)
        Next i
        
        ImportVBAModules ActiveWorkbook, sVbaCodePath
        MsgBox ("Code library saved and imported into active workbook")
            
    End If
    
    StandardExit

End Sub
