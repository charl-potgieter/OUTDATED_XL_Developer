Attribute VB_Name = "PopUpMenuCreator"
Option Explicit
Option Private Module



Sub ReadMenuConfig(ByRef MenuItems() As TypeMenuConfig)
'Each row in sheet MenuBuilder contains columns in below order
' (a) A number of PopupCaption menu names (at least one)
' (b) ButtonCaption name
' (c) Name of spreadsheet containing the sub
' (d) Name of the sub being called

    Dim shtMenuConfig As Worksheet
    Dim NumberOfMenuItems As Integer
    Dim NumberOfPopupMenuItems As Integer
    Dim IndexOfCurrentMenuItem As Integer
    Dim IndexOfCurrentPopupMenu As Integer
    Dim CurrentMenuItem As TypeMenuConfig
    Dim CellToRead As Range

    Set shtMenuConfig = ThisWorkbook.Worksheets("MenuConfig")
    NumberOfMenuItems = shtMenuConfig.Range("A1").CurrentRegion.Rows.Count
    ReDim MenuItems(0 To NumberOfMenuItems - 1)
    
    For IndexOfCurrentMenuItem = 0 To (NumberOfMenuItems - 1)
        
        Set CellToRead = shtMenuConfig.Cells(IndexOfCurrentMenuItem + 1, 1)
        
        'Subtract 3 for ButtonCaption, spreadsheet and sub name included in each row
        NumberOfPopupMenuItems = WorksheetFunction.CountA(shtMenuConfig.Rows(IndexOfCurrentMenuItem + 1)) - 3
        
        ReDim CurrentMenuItem.PopupMenuCaptions(0 To NumberOfPopupMenuItems - 1)
        IndexOfCurrentPopupMenu = 0
        Do While (IndexOfCurrentPopupMenu + 1) <= NumberOfPopupMenuItems
            CurrentMenuItem.PopupMenuCaptions(IndexOfCurrentPopupMenu) = CellToRead.Value
            IndexOfCurrentPopupMenu = IndexOfCurrentPopupMenu + 1
            Set CellToRead = CellToRead.Offset(0, 1)
        Loop
        
        CurrentMenuItem.ButtonCaption = CellToRead.Value
        
        Set CellToRead = CellToRead.Offset(0, 1)
        CurrentMenuItem.SpreadsheetName = CellToRead.Value
        
        Set CellToRead = CellToRead.Offset(0, 1)
        CurrentMenuItem.SubName = CellToRead.Value
        
        MenuItems(IndexOfCurrentMenuItem) = CurrentMenuItem
        
    Next IndexOfCurrentMenuItem

End Sub



Sub GenerateMenu(ByVal PopupCaptionMenuName As String, ByRef MenuConfig() As TypeMenuConfig)

    Dim cb As CommandBar
    Dim MenuCategory As CommandBarPopup
    Dim MenuCategoryParent As Object
    Dim LastPopUpMenuInCurrentMenuItem
    Dim Button As CommandBarControl
    Dim CurrentMenuItem As TypeMenuConfig
    Dim IndexOfCurrentMenuItem As Integer
    Dim IndexOfCurrentPopupMenu As Integer
    Dim sPopUpItemName
    
    
    On Error Resume Next
    Application.CommandBars(PopupCaptionMenuName).Delete
    On Error GoTo 0
    
    Set cb = Application.CommandBars.Add(Name:=PopupCaptionMenuName, Position:=msoBarPopup, _
                                     MenuBar:=False, Temporary:=True)

    For IndexOfCurrentMenuItem = LBound(MenuConfig) To UBound(MenuConfig)
        
        CurrentMenuItem = MenuConfig(IndexOfCurrentMenuItem)
        
        For IndexOfCurrentPopupMenu = LBound(CurrentMenuItem.PopupMenuCaptions) To UBound(CurrentMenuItem.PopupMenuCaptions)
            
            sPopUpItemName = GetCurentPopUpItemName(PopupCaptionMenuName, _
                CurrentMenuItem, IndexOfCurrentPopupMenu)
               
            Select Case True
            
                Case CommandBarPopUpExists(sPopUpItemName)
                'Do Nothing
            
                Case IndexOfCurrentPopupMenu = LBound(CurrentMenuItem.PopupMenuCaptions)
                'First layer - add to commandbar
                    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
                    MenuCategory.Caption = CurrentMenuItem.PopupMenuCaptions(IndexOfCurrentPopupMenu)
                    MenuCategory.CommandBar.Name = sPopUpItemName
                Case Else
                    Set MenuCategoryParent = GetMenuCategoryPopupParent(sPopUpItemName)
                    Set MenuCategory = MenuCategoryParent.Controls.Add(Type:=msoControlPopup)
                    MenuCategory.Caption = CurrentMenuItem.PopupMenuCaptions(IndexOfCurrentPopupMenu)
                    MenuCategory.CommandBar.Name = sPopUpItemName
            End Select
           
            
        Next IndexOfCurrentPopupMenu
        
        Set LastPopUpMenuInCurrentMenuItem = GetLastPopUpMenu(PopupCaptionMenuName, MenuConfig(IndexOfCurrentMenuItem))
        Set Button = MenuCategory.Controls.Add(Type:=msoControlButton)
        Button.Caption = MenuConfig(IndexOfCurrentMenuItem).ButtonCaption
        Button.OnAction = "'" & MenuConfig(IndexOfCurrentMenuItem).SpreadsheetName & _
            "'!" & MenuConfig(IndexOfCurrentMenuItem).SubName
        
    Next IndexOfCurrentMenuItem

End Sub

Function GetCurentPopUpItemName(ByVal PopupCaptionMenuName As String, _
    ByRef MenuConfig As TypeMenuConfig, ByVal CurrentPopIndex As Integer) As String
'Returns name as a concatenation of previous and current popup itms at current menu level
'This is done to ensure unique popup menu names

    Dim i As Integer
    
    For i = LBound(MenuConfig.PopupMenuCaptions) To CurrentPopIndex
        If i = LBound(MenuConfig.PopupMenuCaptions) Then
            GetCurentPopUpItemName = PopupCaptionMenuName & "|" & MenuConfig.PopupMenuCaptions(i)
        Else
            GetCurentPopUpItemName = GetCurentPopUpItemName & "|" & MenuConfig.PopupMenuCaptions(i)
        End If
    Next i

End Function


Function GetMenuCategoryPopupParent(ByVal sMenuCategoryName) As Object
'In this menu the popup names are set as Prefix|Level1Name|Level2Name|... etc

    Dim sParentName As String
    Dim iPositionOfLastDelimiter As Integer
    
    iPositionOfLastDelimiter = InStrRev(sMenuCategoryName, "|")
    sParentName = Left(sMenuCategoryName, iPositionOfLastDelimiter - 1)
    Set GetMenuCategoryPopupParent = Application.CommandBars(sParentName)

End Function

Function GetLastPopUpMenu(ByVal PopupCaptionMenuName As String, MenuDetail As TypeMenuConfig) As Object
'In this menu the popup names are set as Prefix|Level1Name|Level2Name|... etc

    Dim sLastPopupMenuName As String
    Dim i As Integer
    
    sLastPopupMenuName = PopupCaptionMenuName
    For i = UBound(MenuDetail.PopupMenuCaptions) To LBound(MenuDetail.PopupMenuCaptions)
        sLastPopupMenuName = sLastPopupMenuName & "|" & MenuDetail.PopupMenuCaptions(i)
    Next i

    Set GetLastPopUpMenu = Application.CommandBars(sLastPopupMenuName)

End Function


Function CommandBarPopUpExists(ByVal sPopUpName As String) As Boolean

    Dim sTest As String
    
    On Error Resume Next
    sTest = Application.CommandBars(sPopUpName).Name
    CommandBarPopUpExists = (Err.Number = 0)
    On Error GoTo 0

End Function




