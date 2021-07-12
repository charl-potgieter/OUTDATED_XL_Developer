Attribute VB_Name = "EntryPoints"
Option Explicit

Public Type TypeKeyboardMenuDetails
    PopupCaptions() As String
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


Sub RefreshCodeLibrariesInActiveWorkbookFromGithubSource(Optional control As IRibbonControl)

    Dim sTargetDirectory As String
    Dim sTargetFileName As String
    Dim sTargetFilePathAndName As String
    Dim sModuleName As String
    Dim rngCell As Range
    Dim sUrl As String
    Dim wkb As Workbook

    Set wkb = ActiveWorkbook
    If wkb.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook.  " & _
        "Not possible to import in this workbook "
        Exit Sub
    End If
    
    
    sTargetDirectory = Environ("Temp") & Application.PathSeparator & "Vba_Libraries"
    On Error Resume Next
    Kill sTargetDirectory & Application.PathSeparator & "*.*"
    RmDir sTargetDirectory
    On Error GoTo 0
    CreateFolder sTargetDirectory
    
    
    For Each rngCell In ThisWorkbook.Sheets("StandardCodeLibraries").Range("A1").CurrentRegion
        sUrl = rngCell.Value
        sTargetFileName = Right(sUrl, (Len(sUrl) - InStrRev(sUrl, "/")))
        sModuleName = Left(sTargetFileName, InStrRev(sTargetFileName, ".") - 1)
        sTargetFilePathAndName = sTargetDirectory & Application.PathSeparator & sTargetFileName
        
        DeleteModule wkb, sModuleName
        DownloadFileFromUrl sUrl, sTargetFilePathAndName
        ConvertTextFileUnixToWindowsLineFeeds sTargetFilePathAndName
        
    Next rngCell
    
    ImportVBAModules wkb, sTargetDirectory
    MsgBox "Refresh complete"

End Sub



Sub ListGithubCodeLibraries(Optional control As IRibbonControl)

    Dim sht As Worksheet
    
    Set sht = Application.Workbooks.Add.Sheets(1)
    
    ThisWorkbook.Sheets("StandardCodeLibraries").Range("A1").CurrentRegion.Copy
    sht.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    sht.Activate
    sht.Cells.EntireColumn.AutoFit
    sht.Range("A1").Select
    Application.WindowState = xlMaximized 'maximize Excel
    ActiveWindow.WindowState = xlMaximized 'maximize the workbook in Excel
    sht.Parent.Saved = True

End Sub


Sub ReplaceGithubCodeLibrariesWithSelection(Optional control As IRibbonControl)

    With ThisWorkbook.Sheets("StandardCodeLibraries")
        .Cells.EntireRow.Delete
        Selection.Copy
        .Range("A1").PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    ThisWorkbook.Save
    MsgBox ("Code libraries updated")

End Sub


Sub ShowPopupMenu()

    Dim MenuDetails() As TypeKeyboardMenuDetails
    Const PopupCaptionMenuName As String = "VbaMgrPopupCaptionMenu"
    
    ReadMenuDetails MenuDetails
    GenerateMenu MenuDetails, PopupCaptionMenuName
    Application.CommandBars(PopupCaptionMenuName).ShowPopup

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
    
    ThisWorkbook.Sheets("MenuBuilder").Range("A1").CurrentRegion.Copy
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

    With ThisWorkbook.Sheets("MenuBuilder")
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

