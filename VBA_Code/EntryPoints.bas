Attribute VB_Name = "EntryPoints"
Option Explicit

Public Type TypeMenuConfig
    PopupMenuCaptions() As String
    ButtonCaption As String
    SpreadsheetName As String
    SubName As String
End Type


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



Sub ShowPopupMenu()

    Dim MenuItems() As TypeMenuConfig
    Const PopupMenuName As String = "VbaMgrPopupMenu"
    
    ReadMenuConfig MenuItems
    GenerateMenu PopupMenuName, MenuItems
    Application.CommandBars(PopupMenuName).ShowPopup

End Sub


Sub GenerateExampleMenuConfig(Optional control As IRibbonControl)

    Dim sht As Worksheet
    
    Set sht = Application.Workbooks.Add.Sheets(1)
    ThisWorkbook.Sheets("MenuConfigExample").Range("A1").CurrentRegion.Copy
    
    sht.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    sht.Activate
    sht.Cells.EntireColumn.AutoFit
    sht.Range("A1").Select
    Application.WindowState = xlMaximized 'maximize Excel
    ActiveWindow.WindowState = xlMaximized 'maximize the workbook in Excel
    sht.Parent.Saved = True
    

End Sub


Sub ListCurrentMenuConfig(Optional control As IRibbonControl)

    Dim sht As Worksheet
    
    Set sht = Application.Workbooks.Add.Sheets(1)
    
    ThisWorkbook.Sheets("MenuConfig").Range("A1").CurrentRegion.Copy
    sht.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    sht.Activate
    sht.Cells.EntireColumn.AutoFit
    sht.Range("A1").Select
    Application.WindowState = xlMaximized 'maximize Excel
    ActiveWindow.WindowState = xlMaximized 'maximize the workbook in Excel
    sht.Parent.Saved = True

End Sub


Sub ReplaceMenuConfigWithSelection(Optional control As IRibbonControl)

    With ThisWorkbook.Sheets("MenuConfig")
        .Cells.EntireRow.Delete
        Selection.Copy
        .Range("A1").PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    ThisWorkbook.Save
    MsgBox ("Menu configuration updated")

End Sub


Sub ChangePopUpMenuKeyboardShortcut(Optional control As IRibbonControl)

    Dim ShortCutKey As String
    Dim PreviousShortCutKey As String
    
    ShortCutKey = InputBox("Change shortcut to " & vbCrLf & _
        "<ctrl> <shift> and " & vbCrLf & _
        "<Enter single key below...>")
    
    If Len(ShortCutKey) <> 1 Then
        MsgBox ("A single key is required - shortcut not updated")
    Else
        'Delete previous shortcut key
        PreviousShortCutKey = ThisWorkbook.Sheets("KeyboardShortcut").Range("KeyboardShortcutKey").Value
        Application.OnKey "^+{" & LCase(PreviousShortCutKey) & "}", ""
        
        'Implement new shortcut key
        ThisWorkbook.Sheets("KeyboardShortcut").Range("KeyboardShortcutKey").Value = LCase(ShortCutKey)
        Application.OnKey "^+{" & LCase(ShortCutKey) & "}", "ShowPopupMenu"
        ThisWorkbook.Save
        
        MsgBox ("Shortcut updated")
    End If

End Sub


Sub SaveStandardCodeLibraryAndImportIntoCurrentWorkbook()

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
