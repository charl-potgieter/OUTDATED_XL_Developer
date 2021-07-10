Attribute VB_Name = "General"
Option Explicit

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


Sub DownloadFileFromUrl(ByVal sUrl As String, ByVal sTargetFilePathAndName)

    
    Dim oStream
    
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", sUrl, False
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile sTargetFilePathAndName, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If

End Sub


Function ConvertTextFileUnixToWindowsLineFeeds _
    (sSourceFilePathAndName As String, Optional sTargetFilePathAndName As String = "") As Boolean
'Inspired by : https://newtonexcelbach.com/2015/11/10/importing-text-files-unix-format/
'*Nix operating systems utilise different line feeds in text files compared to Windows.
'This function converts to Windows Format

    Dim WholeLine As String
    Dim iFileNo As Integer

    'Get first free file number
    iFileNo = FreeFile

    If sTargetFilePathAndName = "" Then sTargetFilePathAndName = sSourceFilePathAndName
    Open sSourceFilePathAndName For Input Access Read As #iFileNo

    Line Input #iFileNo, WholeLine
    If EOF(iFileNo) Then
        WholeLine = Replace(WholeLine, vbLf, vbCrLf)
        Close #iFileNo
        Open sTargetFilePathAndName For Output Access Write As #iFileNo
        Print #iFileNo, WholeLine
        Close #iFileNo
        ConvertTextFileUnixToWindowsLineFeeds = True
    Else
        ConvertTextFileUnixToWindowsLineFeeds = False
    End If

End Function

'Sub ConvertTextFileUnixToWindowsLineFeeds(sSourceFilePathAndName As String, _
'    Optional sTargetFilePathAndName As String = "")
'''*Nix operating systems utilise different line feeds in text files compared to Windows.
'''This function converts to Windows Format
'
'    Dim sFileContents As String
'
'    If sTargetFilePathAndName = "" Then
'        sTargetFilePathAndName = sSourceFilePathAndName
'    End If
'
'    sFileContents = ReadTextFileIntoString(sSourceFilePathAndName)
'    Replace sFileContents, vbLf, vbCrLf
'    WriteStringToTextFile sFileContents, sTargetFilePathAndName
'
'End Sub
'
'
'
'Function ReadTextFileIntoString(sFilePath As String) As String
''Inspired by:
''https://analystcave.com/vba-read-file-vba/
'
'    Dim iFileNo As Integer
'
'    'Get first free file number
'    iFileNo = FreeFile
'
'    Open sFilePath For Input As #iFileNo
'    ReadTextFileIntoString = Input$(LOF(iFileNo), iFileNo)
'    Close #iFileNo
'
'End Function
'
'
'Function WriteStringToTextFile(ByVal sStr As String, ByVal sFilePath As String)
''Requires reference to Microsoft Scripting Runtime
''Writes sStr to a text file
''*** THIS WILL OVERWRITE ANY CURRENT CONTENT OF THE FILE ***
'
'    Dim fso As Object
'    Dim oFile As Object
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set oFile = fso.CreateTextFile(sFilePath)
'    oFile.Write (sStr)
'    oFile.Close
'    Set fso = Nothing
'    Set oFile = Nothing
'
'End Function
'
'
