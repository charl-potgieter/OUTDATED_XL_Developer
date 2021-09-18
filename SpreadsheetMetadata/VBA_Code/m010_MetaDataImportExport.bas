Attribute VB_Name = "m010_MetaDataImportExport"
Option Explicit
Option Private Module

Sub GenerateSpreadsheetMetadata(ByVal wkb As Workbook)

'Generates selected spreadsheet data to allow the spreadsheet to be recreated
'via VBA.
'Aspects covered include:
'   - Sheet names
'   - Sheet category
'   - Sheet heading
'   - Table name
'   - Number of table columns
'   -  Listobject number format
'   -  Listobject font colour

    Dim sMetaDataRootPath As String
    Dim sTableStructurePath As String
    Dim sVbaCodePath As String
    Dim sOtherPath As String

    sMetaDataRootPath = wkb.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sTableStructurePath = sMetaDataRootPath & Application.PathSeparator & "TableStructure"
    sVbaCodePath = sMetaDataRootPath & Application.PathSeparator & "VBA_Code"
    sOtherPath = sMetaDataRootPath & Application.PathSeparator & "Other"
    

    'Create folders for storing metadata
    If Not FolderExists(sMetaDataRootPath) Then CreateFolder sMetaDataRootPath
    If Not FolderExists(sTableStructurePath) Then CreateFolder sTableStructurePath
    If Not FolderExists(sVbaCodePath) Then CreateFolder sVbaCodePath
    If Not FolderExists(sOtherPath) Then CreateFolder sOtherPath

    'Delete any old files in above folders
    On Error Resume Next
    Kill sTableStructurePath & Application.PathSeparator & "*.*"
    Kill sVbaCodePath & Application.PathSeparator & "*.*"
    Kill sOtherPath & Application.PathSeparator & "*.*"
    On Error GoTo 0

    'Generate listobject metadata
    GenerateMetadataFileListObjectFields wkb, _
        sTableStructurePath & Application.PathSeparator & "ListObjectFields.txt"
    GenerateMetadataFileListObjectValues wkb, _
        sTableStructurePath & Application.PathSeparator & "ListObjectFieldValues.txt"
    GenerateMetadataFileListObjectFormat wkb, _
        sTableStructurePath & Application.PathSeparator & "ListObjectFieldFormats.txt"

    'Export VBA code
    ExportVBAModules wkb, sVbaCodePath
    
    'Generate other info
    GenerateMetadataOther wkb, _
        sOtherPath & Application.PathSeparator & "OtherData.txt"
    


End Sub



Sub GenerateMetadataFileListObjectFields(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim sht As Worksheet
    Dim ListStorage As zLIB_ListStorage
    Dim StorageAssigned As Boolean
    Dim lo As ListObject
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|IsFormula|Formula";
    
    Set ListStorage = New zLIB_ListStorage
    
    For Each sht In wkb.Worksheets
        StorageAssigned = ListStorage.AssignStorage(wkb, sht.Name)
        If StorageAssigned Then
            Set lo = ListStorage.ListObj
            If ListStorage.IsEmpty Then ListStorage.AddBlankRow
            For i = 1 To lo.HeaderRowRange.Columns.Count
                sRowToWrite = vbCr & _
                    sht.Name & "|" & _
                    lo.Name & "|" _
                    & lo.HeaderRowRange.Cells(i) & "|" & _
                    lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula & "|"
                If lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula Then
                    sRowToWrite = sRowToWrite & lo.ListColumns(i).DataBodyRange.Cells(1).Formula
                End If
                Print #iFileNo, sRowToWrite;
            Next i
        End If
    Next sht
    
    Close #iFileNo
    Set ListStorage = Nothing

End Sub


Sub GenerateMetadataFileListObjectValues(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim j As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim ListStorage As zLIB_ListStorage
    Dim StorageAssigned As Boolean
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    
    
    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|Value";
    
    Set ListStorage = New zLIB_ListStorage
    
    'Write value row by row
    For Each sht In wkb.Worksheets
        StorageAssigned = ListStorage.AssignStorage(wkb, sht.Name)
        If StorageAssigned Then
            Set lo = ListStorage.ListObj
            If ListStorage.IsEmpty Then ListStorage.AddBlankRow
            For i = 1 To lo.DataBodyRange.Rows.Count
                For j = 1 To lo.HeaderRowRange.Columns.Count
                    If Not (lo.ListColumns(j).DataBodyRange.Cells(1).HasFormula) Then
                        sRowToWrite = vbCr & _
                            sht.Name & "|" & _
                            lo.Name & "|" & _
                            lo.ListColumns(j).Name & "|" & _
                            lo.ListColumns(j).DataBodyRange.Cells(i).Value
                            Print #iFileNo, sRowToWrite;
                    End If
                Next j
            Next i
        End If
    Next sht
    
    Close #iFileNo
    Set ListStorage = Nothing

End Sub


Sub GenerateMetadataFileListObjectFormat(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim ListStorage As zLIB_ListStorage
    Dim StorageAssigned As Boolean
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|NumberFormat|FontColour";
    
    Set ListStorage = New zLIB_ListStorage
    
    For Each sht In wkb.Worksheets
        StorageAssigned = ListStorage.AssignStorage(wkb, sht.Name)
        If StorageAssigned Then
            If ListStorage.IsEmpty Then ListStorage.AddBlankRow
            Set lo = ListStorage.ListObj
            For i = 1 To lo.HeaderRowRange.Columns.Count
                sRowToWrite = vbCr & _
                    sht.Name & "|" & _
                    lo.Name & "|" & _
                        lo.HeaderRowRange.Cells(i) & "|" & _
                        lo.ListColumns(i).DataBodyRange.Cells(1).NumberFormat & "|" & _
                        lo.ListColumns(i).DataBodyRange.Cells(1).Font.Color
                Print #iFileNo, sRowToWrite;
            Next i
        End If
                
    Next sht
    Close #iFileNo
    Set ListStorage = Nothing

End Sub


Sub GenerateMetadataOther(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim iFileNo As Integer
    Dim sRowToWrite As String
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject

    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "Item|Value";
            
    sRowToWrite = vbCr & _
        "FileName|" & fso.GetBaseName(wkb.Name)
    Print #iFileNo, sRowToWrite;
                
    Close #iFileNo

    Set fso = Nothing

End Sub


Sub FormatCoverSheet(ByVal sht As Worksheet, ByVal FileName As String)

    With sht
        .Activate
        .Move Before:=sht.Parent.Sheets(1)
        .Name = "Cover"
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Size = 16
        .Range("B2").Value = FileName
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80
    End With

End Sub


Function CreateNewWorkbookWithOneSheet() As Workbook

    Dim wkb As Workbook
    
    Set wkb = Application.Workbooks.Add
    Do While wkb.Sheets.Count > 1
        wkb.Sheets(1).Delete
    Loop

    Set CreateNewWorkbookWithOneSheet = wkb
    Set wkb = Nothing
    
End Function





