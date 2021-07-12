Attribute VB_Name = "PopUpMenuCreator"
Option Explicit
Option Private Module



Sub ReadMenuDetails(ByRef MenuDetails() As TypeKeyboardMenuDetails)
'Each row in sheet MenuBuilder contains columns in below order
' (a) Optional number of PopupCaption menu names
' (b) ButtonCaption name
' (c) Name of spreadsheet containing the sub
' (d) Name of the sub being called


    Dim shtMenuDetails As Worksheet
    Dim iNumberOfMenuItems As Integer
    Dim iNumberOfColumnsInRow
    Dim iNumberOfPopupCaptionMenuItems As Integer
    Dim iCurrentColumnNumber As Integer
    Dim iCurrentRowNumber As Integer
    

    Set shtMenuDetails = ThisWorkbook.Worksheets("MenuBuilder")
    iNumberOfMenuItems = shtMenuDetails.Range("A1").CurrentRegion.Rows.Count
    ReDim MenuDetails(0 To iNumberOfMenuItems - 1)
    
    For iCurrentRowNumber = 1 To iNumberOfMenuItems
        iNumberOfColumnsInRow = WorksheetFunction.CountA(shtMenuDetails.Rows(iCurrentRowNumber))
        'Subtract 3 for ButtonCaption, spreadsheet and sub name included in each row
        iNumberOfPopupCaptionMenuItems = iNumberOfColumnsInRow - 3
        
        ReDim MenuDetails(iCurrentRowNumber - 1).PopupCaptions(0 To iNumberOfPopupCaptionMenuItems - 1)
        iCurrentColumnNumber = 1
        Do While iCurrentColumnNumber <= iNumberOfPopupCaptionMenuItems
            MenuDetails(iCurrentRowNumber - 1).PopupCaptions(iCurrentColumnNumber - 1) = shtMenuDetails.Cells(iCurrentRowNumber, iCurrentColumnNumber)
            iCurrentColumnNumber = iCurrentColumnNumber + 1
        Loop
        
        MenuDetails(iCurrentRowNumber - 1).ButtonCaption = shtMenuDetails.Cells(iCurrentRowNumber, iCurrentColumnNumber)
        iCurrentColumnNumber = iCurrentColumnNumber + 1
        MenuDetails(iCurrentRowNumber - 1).SpreadsheetName = shtMenuDetails.Cells(iCurrentRowNumber, iCurrentColumnNumber)
        iCurrentColumnNumber = iCurrentColumnNumber + 1
        MenuDetails(iCurrentRowNumber - 1).SubName = shtMenuDetails.Cells(iCurrentRowNumber, iCurrentColumnNumber)
        
    Next iCurrentRowNumber

End Sub



Sub GenerateMenu(ByRef MenuDetails() As TypeKeyboardMenuDetails, ByVal PopupCaptionMenuName As String)

    Dim cb As CommandBar
    Dim MenuCategory As CommandBarPopup
    Dim MenuCategoryParent ' As CommandBarPopup
    Dim MenuItem As CommandBarControl
    Dim CurrentMenuIndex As Integer
    Dim CurrentPopIndex As Integer
    Dim sPopUpItemName
    
    
    On Error Resume Next
    Application.CommandBars(PopupCaptionMenuName).Delete
    On Error GoTo 0
    
    Set cb = Application.CommandBars.Add(Name:=PopupCaptionMenuName, Position:=msoBarPopup, _
                                     MenuBar:=False, Temporary:=True)

    For CurrentMenuIndex = LBound(MenuDetails) To UBound(MenuDetails)
        
        For CurrentPopIndex = LBound(MenuDetails(CurrentMenuIndex).PopupCaptions) To UBound(MenuDetails(CurrentMenuIndex).PopupCaptions)
            
            sPopUpItemName = GetCurentPopUpItemName(MenuDetails(CurrentMenuIndex), CurrentPopIndex)
           
            Select Case True
            
                Case CommandBarPopUpExists(sPopUpItemName)
                'Do Nothing
            
                Case CurrentPopIndex = LBound(MenuDetails(CurrentMenuIndex).PopupCaptions)
                'First layer - add to commandbar
                    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
                    MenuCategory.Caption = MenuDetails(CurrentMenuIndex).PopupCaptions(CurrentPopIndex)
                    MenuCategory.CommandBar.Name = sPopUpItemName
                Case Else
                    Set MenuCategoryParent = GetMenuCategoryPopupParent(cb, sPopUpItemName)
                    Set MenuCategory = MenuCategoryParent.Controls.Add(Type:=msoControlPopup)
                    MenuCategory.Caption = MenuDetails(CurrentMenuIndex).PopupCaptions(CurrentPopIndex)
                    MenuCategory.CommandBar.Name = sPopUpItemName
            End Select
           
            
        Next CurrentPopIndex
        
        Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
        MenuItem.Caption = MenuDetails(CurrentMenuIndex).ButtonCaption
        MenuItem.OnAction = "'" & MenuDetails(CurrentMenuIndex).SpreadsheetName & _
            "'!" & MenuDetails(CurrentMenuIndex).SubName
        
    Next CurrentMenuIndex

End Sub

Function GetCurentPopUpItemName(ByRef MenuDetail As TypeKeyboardMenuDetails, _
    ByVal CurrentPopIndex As Integer) As String
'Returns name as a concatenation of previous and current popup itms at current menu level
'This is done to ensure unique popup menu names

    Dim I As Integer
    Const UniqueNamePrefix As String = "VbaCodeManager|" 'To Ensure no clash with existing commandbars
    
    For I = LBound(MenuDetail.PopupCaptions) To CurrentPopIndex
        If I = LBound(MenuDetail.PopupCaptions) Then
            GetCurentPopUpItemName = UniqueNamePrefix & MenuDetail.PopupCaptions(I)
        Else
            GetCurentPopUpItemName = GetCurentPopUpItemName & "|" & MenuDetail.PopupCaptions(I)
        End If
    Next I

End Function


Function GetMenuCategoryPopupParent(ByVal cb As CommandBar, ByVal sMenuCategoryName) As Object
'In this menu the popup names are set as Prefix|Level1Name|Level2Name|... etc

    Dim sParentName As String
    Dim iPositionOfLastDelimiter As Integer
    
    iPositionOfLastDelimiter = InStrRev(sMenuCategoryName, "|")
    sParentName = Left(sMenuCategoryName, iPositionOfLastDelimiter - 1)
    Set GetMenuCategoryPopupParent = Application.CommandBars(sParentName)

End Function


Function CommandBarPopUpExists(ByVal sPopUpName As String) As Boolean

    Dim sTest As String
    
    On Error Resume Next
    sTest = Application.CommandBars(sPopUpName).Name
    CommandBarPopUpExists = (Err.Number = 0)
    On Error GoTo 0

End Function




