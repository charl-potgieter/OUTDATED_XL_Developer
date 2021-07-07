Attribute VB_Name = "MenuGenerator"
Option Explicit
Option Private Module


Sub DeletePopUpMenu()
'Delete PopUp menu if it exists
    
    On Error Resume Next
    Application.CommandBars(gcsMenuName).Delete
    On Error GoTo 0
    
End Sub



Sub CreatePopUpMenu()

    Dim cb As CommandBar
    Dim MenuCategory As CommandBarPopup
    Dim MenuSubcategory As CommandBarPopup
    Dim MenuItem As CommandBarControl
    
    Set cb = Application.CommandBars.Add(Name:=gcsMenuName, Position:=msoBarPopup, _
                                     MenuBar:=False, Temporary:=True)
    
    
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Code Manager"

    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export active workbook VBA code (overwrites existing)"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportActiveWorkbookVbaCode"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Refresh standard code libraries in active workbook from github source"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "RefreshCodeLibrariesInActiveWorkbookFromGithubSource"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "List Github code libraries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ListCodeLibraries"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Replace Github code libraries with selection"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ReplaceGithubCodeLibrariesWithSelection"
    
    
    
    
End Sub



