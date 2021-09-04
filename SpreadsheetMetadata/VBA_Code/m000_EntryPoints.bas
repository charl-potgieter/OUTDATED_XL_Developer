Attribute VB_Name = "m000_EntryPoints"
Option Explicit


Sub GenerateSpreadsheetMetaData() '(Optional control As IRibbonControl)

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
    Dim sWorksheetStructurePath As String
    Dim sPowerQueriesPath As String
    Dim sVbaCodePath As String

    StandardEntry

    sMetaDataRootPath = ActiveWorkbook.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sWorksheetStructurePath = sMetaDataRootPath & Application.PathSeparator & "WorksheetStructure"
    sPowerQueriesPath = sMetaDataRootPath & Application.PathSeparator & "PowerQueries"
    sVbaCodePath = sMetaDataRootPath & Application.PathSeparator & "VBA_Code"

    'Create folders for storing metadata
    If Not FolderExists(sMetaDataRootPath) Then CreateFolder sMetaDataRootPath
    If Not FolderExists(sWorksheetStructurePath) Then CreateFolder sWorksheetStructurePath
    If Not FolderExists(sPowerQueriesPath) Then CreateFolder sPowerQueriesPath
    If Not FolderExists(sVbaCodePath) Then CreateFolder sVbaCodePath

    'Delete any old files in above folders
    On Error Resume Next
    Kill sWorksheetStructurePath & Application.PathSeparator & "*.*"
    Kill sPowerQueriesPath & Application.PathSeparator & "*.*"
    Kill sVbaCodePath & Application.PathSeparator & "*.*"
    On Error GoTo 0

    'Generate Worksheet structure metadata text files
    GenerateMetadataFileWorksheets ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "MetadataWorksheets.txt"
    GenerateMetadataFileListObjectFields ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFields.txt"
    GenerateMetadataFileListObjectValues ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFieldValues.txt"
    GenerateMetadataFileListObjectFormat ActiveWorkbook, sWorksheetStructurePath & Application.PathSeparator & "ListObjectFormat.txt"

    'Export VBA code
    ExportVBAModules ActiveWorkbook, sVbaCodePath

    'Export Power Queries
    ExportPowerQueriesToFiles sPowerQueriesPath, ActiveWorkbook

    MsgBox ("Metadata created")
    StandardExit


End Sub


Sub SaveStandardCodeLibraryAndImportIntoCurrentWorkbook(Optional control As IRibbonControl)

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

        ExportVBAModules ThisWorkbook, sVbaCodePath, , "zLIB"

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
