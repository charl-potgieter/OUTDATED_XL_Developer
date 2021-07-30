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



Sub SaveStandardCodeLibraryAndImportIntoCurrentWorkbook(Optional control As IRibbonControl)

    Dim wkbCodeSource As Workbook
    Dim sVbaCodePath As String
    Dim ModuleNames() As String
    Dim i As Long
    Const csCodeLibFileName As String = "ExcelVbaCodeLibrary.xlam"

    Select Case True
        
        Case ActiveWorkbook.Name = ThisWorkbook.Name
            MsgBox ("Cannot apply this action in when " & ThisWorkbook.Name & _
                " is the active workbook")

        Case Not WorkbookIsOpen(csCodeLibFileName)
            MsgBox ("A workbook or add-in named " & csCodeLibFileName & _
            " needs to be open as the source of code libraries.  Exiting.")
            
        Case Else
            
            Set wkbCodeSource = Workbooks(csCodeLibFileName)
            wkbCodeSource.Save
            sVbaCodePath = wkbCodeSource.Path & Application.PathSeparator & "VBA_Code"
            
            'Delete any existing (which may be outdated) code files in folder
            On Error Resume Next
            Kill sVbaCodePath & Application.PathSeparator & "*.*"
            On Error GoTo 0
            
            ExportVBAModules wkbCodeSource, sVbaCodePath
            
            'Delete any current code lib files in active workbook
            GetBaseFileNamesInFolder sVbaCodePath, ModuleNames
            For i = LBound(ModuleNames) To UBound(ModuleNames)
                DeleteModule ActiveWorkbook, ModuleNames(i)
            Next i
            
            
            ImportVBAModules ActiveWorkbook, sVbaCodePath
            MsgBox ("Code library saved and imported into active workbook")
            
    End Select

End Sub
