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
        sTableStructurePath & Application.PathSeparator & "ListObjectFormats.txt"

    'Export VBA code
    ExportVBAModules wkb, sVbaCodePath
    
    'Generate other info
    GenerateMetadataOther wkb, _
        sOtherPath & Application.PathSeparator & "OtherData.txt"
    


End Sub



Sub GenerateMetadataFileListObjectFields(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    Dim fso As Scripting.FileSystemObject

    
    Set fso = New Scripting.FileSystemObject

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|IsFormula|Formula";
    
    For Each sht In wkb.Worksheets
            
        If sht.ListObjects.Count = 1 Then
            Set lo = sht.ListObjects(1)
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

End Sub


Sub GenerateMetadataFileListObjectValues(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim j As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    
    
    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|Value";
    
    For Each sht In wkb.Worksheets
        
        If sht.ListObjects.Count = 1 Then
            Set lo = sht.ListObjects(1)
            For i = 1 To lo.HeaderRowRange.Columns.Count
                If Not (lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula) Then
                    For j = 1 To lo.DataBodyRange.Rows.Count
                        sRowToWrite = vbCr & _
                            sht.Name & "|" & _
                            lo.Name & "|" & _
                            lo.ListColumns(i).Name & "|" & _
                            lo.ListColumns(i).DataBodyRange.Cells(j).Value
                            Print #iFileNo, sRowToWrite;
                    Next j
                End If
            Next i
        End If
                
    Next sht
    Close #iFileNo

End Sub


Sub GenerateMetadataFileListObjectFormat(ByRef wkb As Workbook, ByVal sFilePathAndName As String)

    Dim i As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "SheetName|ListObjectName|ListObjectHeader|NumberFormat|FontColour";
    
    For Each sht In wkb.Worksheets
            
        If sht.ListObjects.Count = 1 Then
            Set lo = sht.ListObjects(1)
            For i = 1 To lo.HeaderRowRange.Columns.Count
            
                sRowToWrite = vbCr & _
                    sht.Name & "|" & _
                    lo.Name & "|" & _
                        lo.HeaderRowRange.Cells(i) & "|" & _
                        lo.ListColumns(i).DataBodyRange.Cells(1).NumberFormat & "|" & _
                        GetCellFontColour(lo.ListColumns(i).DataBodyRange.Cells(1))
                Print #iFileNo, sRowToWrite;
            Next i
        End If
                
    Next sht
    Close #iFileNo

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


Function GenerateListObjectWorkingTableOnNewSheet(ByVal wkb As Workbook, _
    SourceFileBaseName As String, SourcePath As String) As ListObject

    Dim sht As Worksheet
    Dim FilePath As String
    Dim QueryText As String
    
    Set sht = wkb.Sheets.Add
    sht.Name = SourceFileBaseName
    FilePath = SourcePath & Application.PathSeparator & _
        SourceFileBaseName & ".txt"
    QueryText = PipeDelimitedSourcePowerQuery(FilePath)
    Set GenerateListObjectWorkingTableOnNewSheet = CreatePowerQueryTable( _
        sht, SourceFileBaseName, QueryText, "tbl_" & SourceFileBaseName)
        

End Function


Function PipeDelimitedSourcePowerQuery(ByVal sFilePath As String) As String

    PipeDelimitedSourcePowerQuery = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & vbCr & _
        "in " & vbCr & _
        "   PromotedHeaders"

End Function


Function CreatePowerQueryTable( _
    ByVal sht As Worksheet, _
    ByVal sQueryName As String, _
    ByVal sQueryText As String, _
    ByVal sTableName As String) As ListObject

'Creates power query and loads as a table on sht
            
    Dim lo As ListObject
    
    sht.Parent.Queries.Add sQueryName, sQueryText
        
    'Output the Power Query to a worksheet table
    Set lo = sht.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & sQueryName & ";Extended Properties=""""", _
        Destination:=Range("$A$1"))
        
    lo.Name = sTableName
    
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & sQueryName & "]")
        .Refresh BackgroundQuery:=False
    End With
    
    Set CreatePowerQueryTable = lo
        
End Function


Sub FormatCoverSheet(ByVal sht As Worksheet)

    With sht
        .Activate
        .Move Before:=sht.Parent.Sheets(1)
        .Name = "Cover"
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Size = 16
        .Range("B2").Value = Evaluate("=XLOOKUP(""FileName"",tbl_OtherData[Item],tbl_OtherData[Value])")
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 80
    End With

End Sub


Sub CreateListObjectsFromMetadata(ByRef wkb As Workbook, ByVal loFieldDetails As ListObject)
'Records listobject field names and formulas in wkb based on metadata stored in
'wkb.Sheets("Temp_ListObjectFields").ListObjects("tbl_ListObjectFields")

    Dim loTargetListObj As ListObject
    Dim i As Long
    Dim j As Long
    Dim sht As Worksheet
    Dim sSheetName As String
    Dim sListObjName As String
    Dim sListObjHeader As String
    Dim bIsFormula As Boolean
    Dim sFormula As String
    
    'No ListObject field details.  Nothing to do.  Sub exited to prevent error
    'caused by referencing the list databodyrange
    If loFieldDetails.DataBodyRange Is Nothing Then
        Exit Sub
    End If
    
    
    With loFieldDetails
        For i = 1 To .DataBodyRange.Rows.Count
            sSheetName = .ListColumns("SheetName").DataBodyRange.Cells(i)
            sListObjName = .ListColumns("ListObjectName").DataBodyRange.Cells(i)
            sListObjHeader = .ListColumns("ListObjectHeader").DataBodyRange.Cells(i)
            bIsFormula = CBool(.ListColumns("isFormula").DataBodyRange.Cells(i))
            sFormula = .ListColumns("Formula").DataBodyRange.Cells(i)
            
            If Not SheetExists(wkb, sSheetName) Then
                Set sht = wkb.Sheets.Add(after:=wkb.Sheets(wkb.Sheets.Count))
                sht.Name = sSheetName
            End If
            
            'Increment j as header col counter if writing to the table name as previous iteration
            If i = 1 Then
                j = 1
            ElseIf sListObjName = .ListColumns("ListObjectName").DataBodyRange.Cells(i - 1) Then
                j = j + 1
            Else
                j = 1
            End If
            
            Set loTargetListObj = wkb.Worksheets(sSheetName).ListObjects(sListObjName)
            loTargetListObj.HeaderRowRange.Cells(j) = sListObjHeader
            
            If bIsFormula Then
                loTargetListObj.ListColumns(sListObjHeader).DataBodyRange.Formula = sFormula
            End If
            
            
        Next i
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





