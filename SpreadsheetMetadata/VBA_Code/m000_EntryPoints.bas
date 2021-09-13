Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Sub GenerateSpreadsheetMetadataActiveWorkbook() 'Optional control As IRibbonControl

    StandardEntry
    GenerateSpreadsheetMetadata ActiveWorkbook
    MsgBox ("Metadata created")
    StandardExit

End Sub

Sub CreateSpreadsheetFromMetadata()
'Generates spreadsheet from metadata stored in text files in selected folder


    Dim sFolderPath As String
    Dim sFilePath As String
    Dim wkb As Workbook
    Dim sQueryText As String
    Dim sht As Worksheet
    Dim rng As Range
    Dim InitialSheetOnWorkbookCreation As Worksheet
    Dim i As Integer
    Dim loListObjFields As ListObject
    Dim SheetWithListObjData As Worksheet
    Dim loListObjFieldValues As ListObject
    Dim loListObjFormats As ListObject
    Dim loOtherFileMetaData As ListObject
    Dim fso As FileSystemObject
    Dim SheetNameColumn As Range
    Dim TableNameColumn As Range
    Dim TableHeaderColumn As Range
    Dim IsFormulaColumn As Range
    Dim FormulaColumn As Range
    Const StorageRefOFLastFolder As String = _
        "Last utilised folder for creating spreadsheet from metadata"
    
    StandardEntry
    
    'Get folder containing metadata
    sFolderPath = GetFolder(StoredDataValue(StorageRefOFLastFolder))
    If sFolderPath = "" Then
        Exit Sub
    End If
    Set fso = New FileSystemObject
    RangeOfStoredData(StorageRefOFLastFolder).Value = fso.GetParentFolderName(sFolderPath)
    ThisWorkbook.Save

    Set wkb = CreateNewWorkbookWithOneSheet
    Set InitialSheetOnWorkbookCreation = wkb.Sheets(1)
    FormatCoverSheet InitialSheetOnWorkbookCreation
    
    'Create temporary sheets for storing listobject properties and values
    Set loListObjFields = GenerateListObjectWorkingTableOnNewSheet(wkb, "ListObjectFields", _
        sFolderPath & Application.PathSeparator & "TableStructure")
    Set loListObjFieldValues = GenerateListObjectWorkingTableOnNewSheet(wkb, "ListObjectFieldValues", _
        sFolderPath & Application.PathSeparator & "TableStructure")
    Set loListObjFormats = GenerateListObjectWorkingTableOnNewSheet(wkb, "ListObjectFormats", _
        sFolderPath & Application.PathSeparator & "TableStructure")
    Set loListObjFormats = GenerateListObjectWorkingTableOnNewSheet(wkb, "OtherData", _
        sFolderPath & Application.PathSeparator & "Other")

    Set SheetWithListObjData = loListObjFields.Parent
    Set SheetNameColumn = SheetWithListObjData.Range("J:J")
    Set TableNameColumn = SheetWithListObjData.Range("K:K")
    Set TableHeaderColumn = SheetWithListObjData.Range("M:M")
    Set IsFormulaColumn = SheetWithListObjData.Range("N:N")
    Set FormulaColumn = SheetWithListObjData.Range("O:O")
    
    SheetNameColumn.Cells(1).Formula2 = _
        "=UNIQUE(tbl_ListObjectFields[[SheetName]:[ListObjectName]])"
    For i = 1 To SheetNameColumn.CurrentRegion.Rows.Count
        TableHeaderColumn.Cells(1).Formula2 = _
            "=FILTER(" & vbLf & _
            "    tbl_ListObjectFields[[ListObjectHeader]:[Formula]], " & vbLf & _
            "    (tbl_ListObjectFields[SheetName]=""" & SheetNameColumn.Cells(i) & """) * " & vbLf & _
            "    (tbl_ListObjectFields[ListObjectName]=""" & TableNameColumn.Cells(i) & """)" & vbLf & _
            ")"
    Next i




'    SheetNames = WorksheetFunction.Unique(loListObjFields.ListColumns("SheetName").DataBodyRange)
'    For i = LBound(SheetNames, 1) To UBound(SheetNames, 1)
'        With loListObjFields
'            SheetName = SheetNames(i, 1)
'            TableName = WorksheetFunction.Xlookup( _
'                SheetName, _
'                .ListColumns("SheetName").DataBodyRange, _
'                .ListColumns("ListObjectName").DataBodyRange)
'            TableHeaders = WorksheetFunction.Filter( _
'                .ListColumns("ListObjectHeader").DataBodyRange, _
'                Array(.ListColumns("SheetName").DataBodyRange = SheetName))
'        End With
'
'    Next i
'
'
'    'Add sheets to target workbook
'    With wkb.Sheets("ListObjectFields").ListObjects("tbl_ListObjectFields")
'        For Each rng In .ListColumns("SheetName").DataBodyRange
'            If Not SheetExists(wkb, rng.Value) Then
'                Set sht = wkb.Sheets.Add(after:=wkb.Sheets(wkb.Sheets.Count))
'                sht.Activate
'                sht.Name = rng.Value
'                ActiveWindow.DisplayGridlines = False
'                ActiveWindow.Zoom = 80
'            End If
'        Next rng
'    End With
'
'    CreateListObjectsFromMetadata wkb, loListObjFields
''    PopulateListObjectValues wkb
'    SetListObjectFormats wkb
'

'    'Delete temp sheets queries and connections
'    For i = LBound(ArrayOfSourceFiles) To UBound(ArrayOfSourceFiles)
'        wkb.Sheets(ArrayOfSourceFiles(i)).Delete
'        wkb.Queries(ArrayOfSourceFiles(i)).Delete
'        wkb.Connections(1).Delete  'Always delete 1st as index decreases as connections are deleted
'    Next i


    'Import VBA code
    ImportVBAModules wkb, sFolderPath & Application.PathSeparator & "VBA_Code"
        
    wkb.Activate
    ActiveWindow.WindowState = xlMaximized
    MsgBox ("Spreadsheet created")
        
    StandardExit
    

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
        GenerateSpreadsheetMetadata ThisWorkbook

        'Delete any current code lib files in active workbook
        sVbaCodePath = ThisWorkbook.Path & Application.PathSeparator & _
            "SpreadsheetMetadata" & Application.PathSeparator & _
            "VBA_Code"
        GetBaseFileNamesInFolder sVbaCodePath, ModuleNames
        For i = LBound(ModuleNames) To UBound(ModuleNames)
            DeleteModule ActiveWorkbook, ModuleNames(i)
        Next i
      
        ImportVBAModules ActiveWorkbook, sVbaCodePath, "zLIB"
        MsgBox ("Code library saved and imported into active workbook")

    End If

    StandardExit

End Sub
