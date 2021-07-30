Attribute VB_Name = "General"
Option Explicit
Option Private Module

Function FolderExists(ByVal sFolderPath) As Boolean
'Requires reference to Microsoft Scripting runtime
'An alternative solution exists using the DIR function but this seems to result in memory leak and folder is
'not released by VBA

    Dim fso As Scripting.FileSystemObject
    Dim FolderPath As String

    Set fso = New Scripting.FileSystemObject

    If Right(sFolderPath, 1) <> Application.PathSeparator Then
        FolderPath = FolderPath & Application.PathSeparator
    End If

    FolderExists = fso.FolderExists(sFolderPath)
    Set fso = Nothing

End Function



Sub CreateFolder(ByVal sFolderPath As String)
'   Requires reference to Microsoft Scripting runtime

    Dim fso As FileSystemObject

    If FolderExists(sFolderPath) Then
        MsgBox ("Folder already exists, new folder not created")
    Else
        Set fso = New FileSystemObject
        fso.CreateFolder sFolderPath
    End If

    Set fso = Nothing

End Sub


Sub DeleteFirstLineOfTextFile(ByVal sFilePathAndName As String)

    Dim sInput As String
    Dim sOutput() As String
    Dim i As Long
    Dim j As Long
    Dim lSizeOfOutput As Long
    Dim iFileNo As Integer
    
    iFileNo = FreeFile
    
    'Import file lines to array excluding firt line
    Open sFilePathAndName For Input As iFileNo
    i = 0
    j = 0
    Do Until EOF(iFileNo)
        j = j + 1
        Line Input #iFileNo, sInput
        If j > 1 Then
            i = i + 1
            ReDim Preserve sOutput(1 To i)
            sOutput(i) = sInput
        End If
    Loop
    Close #iFileNo
    lSizeOfOutput = i
    
    'Write array to file
    Open sFilePathAndName For Output As 1
    For i = 1 To lSizeOfOutput
        Print #iFileNo, sOutput(i)
    Next i
    Close #iFileNo
    
    
End Sub


Function WorkbookIsOpen(ByVal sWbkName As String) As Boolean
'Checks if workbook is open based on filename including extension

    WorkbookIsOpen = False
    On Error Resume Next
    WorkbookIsOpen = Len(Workbooks(sWbkName).Name) <> 0
    On Error GoTo 0

End Function



Sub FileItemsInFolder(ByVal sFolderPath As String, ByVal bRecursive As Boolean, ByRef FileItems() As Scripting.File)
'Returns an array of files (which can be used to get filename, path etc)
'(Cannot create function due to recursive nature of the code)

    
    Dim fso As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim FileItem As Scripting.File
    
    Set fso = New Scripting.FileSystemObject
    Set SourceFolder = fso.GetFolder(sFolderPath)
    
    For Each FileItem In SourceFolder.Files
    
        If Not IsArrayAllocated(FileItems) Then
            ReDim FileItems(0)
        Else
            ReDim Preserve FileItems(UBound(FileItems) + 1)
        End If
        
        Set FileItems(UBound(FileItems)) = FileItem
        
    Next FileItem
    
    If bRecursive Then
        For Each SubFolder In SourceFolder.SubFolders
            FileItemsInFolder SubFolder.Path, True, FileItems
        Next SubFolder
    End If
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set fso = Nothing
    

End Sub


Function GetBaseFileNamesInFolder(ByVal FolderPath As String, ByRef BaseFileNames() As String)
    
    Dim FileItems() As Scripting.File
    Dim i As Integer
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    
    FileItemsInFolder FolderPath, False, FileItems
    
    ReDim BaseFileNames(LBound(FileItems) To UBound(FileItems))
    For i = LBound(FileItems) To UBound(FileItems)
        BaseFileNames(i) = fso.GetBaseName(FileItems(i).Name)
    Next i
    
End Function



Public Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   CREDIT FOR BELOW = Chip Pearson
'   http://www.cpearson.com/Excel/LegaleseAndDisclaimers.aspx
'
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long
On Error Resume Next

' if Arr is not an array, return FALSE and get out.
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
N = UBound(Arr, 1)
If (Err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    If LBound(Arr) <= UBound(Arr) Then
        ' no error. array has been allocated.
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    ' error. unallocated array
    IsArrayAllocated = False
End If

End Function
