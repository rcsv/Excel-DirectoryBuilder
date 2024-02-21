Option Explicit

'#Region "Private Fields"
    ' Constants definition: Name of the worksheet to use, cell for base path, starting row and column for scanning
    Private Const WORKSHEET_NAME As String = "DirDigger"
    Private Const CELL_BASEPATH As String = "C2"
    Private Const ROW_SCANSTART As Integer = 5
    Private Const COL_SCANSTART As Integer = 2
'#End Region

' [EntryPoint]
' Main entry point: Starts the folder creation process
Public Sub CreateFolders()
    Dim baseDir As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WORKSHEET_NAME) ' Set the DirDigger sheet
    
    ' Start notification
    Call MsgBox("Starting folder creation.", vbOKOnly, "Directory Digger")
    
    ' Retrieve the base directory path
    baseDir = ws.Range(CELL_BASEPATH).Value
    ' Check if the base directory is valid
    If baseDir = "" Or Dir(baseDir, vbDirectory) = "" Then
        MsgBox "The base directory is invalid or does not exist.", vbExclamation
        Exit Sub
    End If
        
    ' Execute the directory creation process
    On Error GoTo ErrorHandler
    Call DirectoryDigger(ws, ROW_SCANSTART, COL_SCANSTART, baseDir)
    
    ' Completion notification
    Call MsgBox("All folders have been created.", vbOKOnly, "Directory Digger")
    
    Exit Sub
    
ErrorHandler:
    ' Error handling: Display error message
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' Private function to recursively create directories
Private Function DirectoryDigger(ws As Worksheet, ByVal current_row As Integer, ByVal current_col As Integer, ByVal current_base As String) As Integer
    Dim newFolderPath As String
    On Error Resume Next ' Error handling to avoid interruption
    
    ' Read folder paths from specified cells and create folders
    While ws.Cells(current_row, current_col).Value <> ""
        newFolderPath = current_base & "\" & ws.Cells(current_row, current_col).Value
        
        ' Create the folder if it does not exist
        If Dir(newFolderPath, vbDirectory) = "" Then
            MkDir newFolderPath
            ' Check for errors during folder creation
            If Err.Number <> 0 Then
                MsgBox "Could not create folder: " & newFolderPath, vbExclamation
                Exit Function
            End If
        End If
        
        ' Move to the next row
        current_row = current_row + 1
        
        ' Process for subfolders
        If ws.Cells(current_row, current_col + 1).Value <> "" Then
            ' Move to the next column (subfolder)
            current_row = DirectoryDigger(ws, current_row, current_col + 1, newFolderPath)
        End If
    Wend
    
    ' Return the last row
    DirectoryDigger = current_row
    On Error GoTo 0 ' Reset error handling
    
End Function

' Opens the specified path's folder in Explorer
Public Sub OpenFolderPath()
    Dim baseDir As String
    baseDir = ws.Range(CELL_BASEPATH).Value
    
    ' Open the folder in Explorer if the path is valid
    If baseDir <> "" And Dir(baseDir, vbDirectory) <> "" Then
        Shell "explorer.exe " & Chr(34) & baseDir & Chr(34), vbNormalFocus
    Else
        MsgBox "The folder path is invalid.", vbExclamation
    End If
End Sub
