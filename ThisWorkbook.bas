' ThisWorkbook module for Excel-DirectoryBuilder project
Option Explicit

' Workbook_Open event: Adds a custom menu item to the cell context menu when the workbook is opened.
Private Sub Workbook_Open()
    AddToCellContextMenu
End Sub

' Workbook_BeforeClose event: Removes the custom menu item from the cell context menu before the workbook is closed.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RemoveFromCellContextMenu
End Sub

' Removes the custom "Create Folders" option from the cell context menu to avoid duplicates and clean up on workbook close.
Private Sub RemoveFromCellContextMenu()
    On Error Resume Next ' Ignore errors in case the control does not exist
    Application.CommandBars("Cell").Controls("Create Folders").Delete ' Attempt to delete the custom menu item
    On Error GoTo 0 ' Turn back on normal error handling
End Sub

' Adds a custom "Create Folders" option to the cell context menu for easy access to the functionality.
Public Sub AddToCellContextMenu()
    Dim ContextMenu As CommandBar
    Dim MenuItem As CommandBarButton
    Set ContextMenu = Application.CommandBars("Cell") ' Reference to the cell context menu
    
    On Error Resume Next ' Ignore errors to handle case where the button already exists
    Application.CommandBars("Cell").Controls("Create Folders").Delete ' Ensure no duplicate buttons
    On Error GoTo 0 ' Resume normal error handling
    
    ' Create a new button in the context menu for creating folders
    Set MenuItem = ContextMenu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Create Folders" ' Text displayed in the context menu
        .OnAction = "CreateFolders" ' Name of the macro to be called when the menu item is clicked
        .FaceId = 485 ' Icon image for the menu item, 485 is a folder icon
        .BeginGroup = True ' Adds a separator before this menu item to group it visually
    End With
End Sub
