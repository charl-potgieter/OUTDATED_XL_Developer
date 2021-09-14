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
    Dim wkb As Workbook
    Dim fso As FileSystemObject
    Dim InitialSheetOnWorkbookCreation As Worksheet
    Dim StorageListObjFields
    Dim StorageListObjFieldValues
    Dim StorageListObjFieldFormats
    Dim StorageOther
    Dim SheetNames As Variant
    Const StorageRefOFLastFolder As String = _
        "Last utilised folder for creating spreadsheet from metadata"


    StandardEntry
    
    'Get folder containing metadata
    sFolderPath = GetFolder(StoredDataValue(StorageRefOFLastFolder))
    If sFolderPath = "" Then
        Exit Sub
    End If
    
    'Save the selected folder for future use
    Set fso = New FileSystemObject
    RangeOfStoredData(StorageRefOFLastFolder).Value = fso.GetParentFolderName(sFolderPath)
    ThisWorkbook.Save

    Set wkb = CreateNewWorkbookWithOneSheet
    Set InitialSheetOnWorkbookCreation = wkb.Sheets(1)
    FormatCoverSheet InitialSheetOnWorkbookCreation

    'Assign storage for the relevant spreadsheet metadata
    Set StorageListObjFields = CreateListObjFieldStorage( _
        sFolderPath & Application.PathSeparator & "TableStructure" & _
            Application.PathSeparator & "ListObjectFields.txt", _
        wkb)
    Set StorageListObjFieldValues = CreateListObjFieldValuesStorage( _
        sFolderPath & Application.PathSeparator & "TableStructure" & _
            Application.PathSeparator & "ListObjectFieldValues.txt", _
        wkb)
    Set StorageListObjFieldFormats = CreateListObjFieldFormatsStorage( _
        sFolderPath & Application.PathSeparator & "TableStructure" & _
            Application.PathSeparator & "ListObjectFieldFormats.txt", _
        wkb)
    Set StorageOther = CreateOtherStorage( _
        sFolderPath & Application.PathSeparator & "Other" & _
            Application.PathSeparator & "OtherData.txt", _
        wkb)

    SheetNames = GetSheetNames(StorageListObjFields)
    AddSheetsToWorkbook wkb, SheetNames
    

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
    wkb.Sheets(1).Select
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
