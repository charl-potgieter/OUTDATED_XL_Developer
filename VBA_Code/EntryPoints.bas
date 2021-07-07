Attribute VB_Name = "EntryPoints"
Option Explicit

Public Const gcsMenuName As String = "Code manager"

Sub DisplayPopUpMenu()
Attribute DisplayPopUpMenu.VB_ProcData.VB_Invoke_Func = "C\n14"

    DeletePopUpMenu
    CreatePopUpMenu
    Application.CommandBars(gcsMenuName).ShowPopup

End Sub



Sub ExportActiveWorkbookVbaCode()

    Dim wkb As Workbook
    Dim sVbaCodePath As String

    Set wkb = ActiveWorkbook
    sVbaCodePath = wkb.Path & Application.PathSeparator & "VBA_Code"

    If Not FolderExists(sVbaCodePath) Then
        CreateFolder sVbaCodePath
    End If

    ExportVBAModules wkb, sVbaCodePath
    MsgBox ("VBA code exported")

End Sub


Sub RefreshCodeLibrariesInActiveWorkbookFromGithubSource()

    Dim sTargetDirectory As String
    Dim sTargetFileName As String
    Dim sTargetFilePathAndName As String
    Dim sModuleName As String
    Dim rngCell As Range
    Dim sUrl As String
    Dim wkb As Workbook

    Set wkb = ActiveWorkbook
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
        
    Next rngCell
    
    ImportVBAModules wkb, sTargetDirectory
    MsgBox "Import complete"

End Sub



Sub ListGithubCodeLibraries()

    Dim sht As Worksheet
    
    Set sht = Application.Workbooks.Add.Sheets(1)
    
    ThisWorkbook.Sheets("StandardCodeLibraries").Range("A1").CurrentRegion.Copy
    sht.Range("A1").PasteSpecial xlPasteValues
    sht.Range("A1").Select
    Application.CutCopyMode = False
    sht.Parent.Saved = True

End Sub


Sub ReplaceGithubCodeLibrariesWithSelection()

    With ThisWorkbook.Sheets("StandardCodeLibraries")
        .Cells.EntireRow.Delete
        Selection.Copy
        .Range("A1").PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    ThisWorkbook.Save

End Sub



