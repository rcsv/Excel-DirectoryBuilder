' ThisWorkbook
Option Explicit

Private Sub Workbook_Open()
    AddToCellContextMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RemoveFromCellContextMenu
End Sub

Private Sub RemoveFromCellContextMenu()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Create Folders").Delete
    On Error GoTo 0
End Sub

Public Sub AddToCellContextMenu()
    Dim ContextMenu As CommandBar
    Dim MenuItem As CommandBarButton
    Set ContextMenu = Application.CommandBars("Cell")
    
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Create Folders").Delete
    On Error GoTo 0
    
    Set MenuItem = ContextMenu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Create Folders"
        .OnAction = "CreateFolders" ' macro name
        .FaceId = 485
        .BeginGroup = True ' separate
    End With
End Sub
