Option Explicit

'#Region "Private Fields"
    Private Const WORKSHEET_NAME As String = "DirDigger"
    Private Const CELL_BASEPATH As String = "C2"
    Private Const ROW_SCANSTART As Integer = 5
    Private Const COL_SCANSTART As Integer = 2
'#End Region

' [EntryPoint]
Public Sub CreateFolders()
    Dim baseDir As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WORKSHEET_NAME)
    
    Call MsgBox("フォルダの作成を開始します。", vbOKOnly, "Directory Digger")
    
    baseDir = ActiveSheet.Range(CELL_BASEPATH).Value
    If baseDir = "" Or Dir(baseDir, vbDirectory) = "" Then
        MsgBox "ベースディレクトリが無効、または存在しません。", vbExclamation
        Exit Sub
    End If
        
    On Error GoTo ErrorHandler
    Call DirectoryDigger(ws, ROW_SCANSTART, COL_SCANSTART, baseDir)
    
    Set ws = Nothing
    Call MsgBox("全てのフォルダ作成が完了いたしました。", vbOKOnly, "Directory Digger")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました： " & Err.Description, vbCritical
End Sub

Private Function DirectoryDigger(ws As Worksheet, ByVal current_row As Integer, ByVal current_col As Integer, ByVal current_base As String) As Integer
    Dim newFolderPath As String
    On Error Resume Next
    
    While ws.Cells(current_row, current_col).Value <> ""
        newFolderPath = current_base & "\" & ws.Cells(current_row, current_col).Value
        
        If Dir(newFolderPath, vbDirectory) = "" Then
            MkDir newFolderPath
            If Err.Number <> 0 Then
                MsgBox "フォルダを作成できません: " & newFolderPath, vbExclamation
                Exit Function
            End If
        End If
        
        current_row = current_row + 1
        
        If ws.Cells(current_row, current_col + 1).Value <> "" Then
            ' move next column for making subfolder
            current_row = DirectoryDigger(ws, current_row, current_col + 1, newFolderPath)
        End If

    Wend
    
    DirectoryDigger = current_row
    On Error GoTo 0
    
End Function

Public Sub OpenFolderPath()
    Dim baseDir As String
    baseDir = ActiveSheet.Range(CELL_BASEPATH).Value
    
    If baseDir <> "" And Dir(baseDir, vbDirectory) <> "" Then
        Shell "explorer.exe " & Chr(34) & baseDir & Chr(34), vbNormalFocus
    Else
        MsgBox "有効なフォルダパスではありません。", vbExclamation
    End If
End Sub
