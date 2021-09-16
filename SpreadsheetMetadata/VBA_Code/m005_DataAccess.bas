Attribute VB_Name = "m005_DataAccess"

Option Explicit

Function CreateListObjFieldStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As zLIB_ListStorage
    
    Dim Storage As zLIB_ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "ListObjectFields"

    Set Storage = New zLIB_ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "ListObjectFields", "ListObjectFields"
    
    Set CreateListObjFieldStorage = Storage
    
End Function



Function CreateListObjFieldValuesStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As zLIB_ListStorage
    
    Dim Storage As zLIB_ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "ListObjectFieldValues"

    Set Storage = New zLIB_ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "ListObjectFieldValues", _
        "ListObjectFieldValues"
    
    Set CreateListObjFieldValuesStorage = Storage
    
End Function



Function CreateListObjFieldFormatsStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As zLIB_ListStorage
    
    Dim Storage As zLIB_ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "ListObjectFieldFormats"

    Set Storage = New zLIB_ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "ListObjectFieldFormats", _
        "ListObjectFieldFormats"
    
    Set CreateListObjFieldFormatsStorage = Storage
    
End Function


Function CreateOtherStorage(ByVal SourceFilePath As String, _
    ByVal TargetStorage) As zLIB_ListStorage
    
    Dim Storage As zLIB_ListStorage

    CreatePipeDelimitedPowerQuery TargetStorage, _
        SourceFilePath, _
        "OtherData"

    Set Storage = New zLIB_ListStorage
    Storage.CreateStorageFromPowerQuery TargetStorage, "OtherData", "OtherData"
    
    Set CreateOtherStorage = Storage
    
End Function


Function GetSheetNames(ByVal ListgObjFieldStorage As Variant) As Variant

    Dim Storage As zLIB_ListStorage
    
    Set Storage = ListgObjFieldStorage
    GetSheetNames = Storage.ItemsInField(sFieldName:="SheetName", bUnique:=True)
    Set Storage = Nothing

End Function



Private Sub CreatePipeDelimitedPowerQuery(ByVal wkb As Workbook, _
    ByVal SourceDelimitedFilePath As String, _
    ByVal QueryName As String)

    Dim QueryString As String
    
    QueryString = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        SourceDelimitedFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & vbCr & _
        "in " & vbCr & _
        "   PromotedHeaders"
    
    wkb.Queries.Add QueryName, QueryString

End Sub



Function GetCreatorFileName(ByVal StorageOther As Variant) As String

    Dim Storage As zLIB_ListStorage
    Set Storage = StorageOther
    GetCreatorFileName = Storage.Xlookup("FileName", "[Item]", "[Value]")
    Set Storage = Nothing

End Function



Function GetListObjName(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String) As String

    Dim Storage As zLIB_ListStorage
    Set Storage = StorageListObjFields
    GetListObjName = Storage.Xlookup(SheetName, "[SheetName]", "[ListObjectName]")
    Set Storage = Nothing

End Function


Function GetListObjHeaders(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String) As Variant

    Dim Storage As zLIB_ListStorage
    Set Storage = StorageListObjFields
    
    Storage.Filter "[SheetName] = """ & SheetName & """"
    GetListObjHeaders = Storage.ItemsInField(sFieldName:="ListObjectHeader", bFiltered:=True)

End Function


Function GetHeaderHasFormula(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String, ByVal Header As String) As Boolean

    Dim Storage As zLIB_ListStorage
    
    Set Storage = StorageListObjFields
    GetHeaderHasFormula = Storage.Xlookup(SheetName & Header, _
        "[SheetName] & [ListObjectHeader]", "[IsFormula]")
    
    Set Storage = Nothing

End Function



Function GetColumnFormula(ByVal StorageListObjFields As Variant, _
    ByVal SheetName As String, ByVal Header As String) As String

    Dim Storage As zLIB_ListStorage
    
    Set Storage = StorageListObjFields
    GetColumnFormula = Storage.Xlookup(SheetName & Header, _
        "[SheetName] & [ListObjectHeader]", "[Formula]")
    
    Set Storage = Nothing

End Function

