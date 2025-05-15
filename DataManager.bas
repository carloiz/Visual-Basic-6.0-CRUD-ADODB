Attribute VB_Name = "DataManager"

Public Sub LoadToGrid(ByVal sql As String, ByRef targetGrid As Object)
    Dim rs As ADODB.Recordset
    Dim col As Integer
    Dim row As Integer

    On Error GoTo ErrHandler

    ' Check database connection
    If conn Is Nothing Then
        MsgBox "Database connection is not initialized.", vbCritical
        Exit Sub
    End If

    If conn.State = adStateClosed Then
        conn.Open
    End If

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    ' Clear existing grid content
    targetGrid.Clear

    targetGrid.Rows = 1
    targetGrid.Cols = rs.Fields.Count
    targetGrid.FixedRows = 0
    targetGrid.FixedCols = 0

    ' Populate header names
    For col = 0 To rs.Fields.Count - 1
        targetGrid.TextMatrix(0, col) = rs.Fields(col).Name
    Next col

    ' Auto-fit column width to fill the grid
    Dim totalWidth As Long
    totalWidth = targetGrid.Width - 88 ' Adjust margin if needed

    Dim colWidth As Long
    colWidth = totalWidth \ rs.Fields.Count

    For col = 0 To rs.Fields.Count - 1
        targetGrid.colWidth(col) = colWidth
    Next col

    ' Populate rows
    row = 1
    Do Until rs.EOF
        targetGrid.AddItem ""
        For col = 0 To rs.Fields.Count - 1
            Dim val As String
            If Not IsNull(rs.Fields(col).value) Then
                val = CStr(rs.Fields(col).value)
            Else
                val = ""
            End If
            targetGrid.TextMatrix(row, col) = val
        Next col
        row = row + 1
        rs.MoveNext
    Loop

    ' Set styles
    With targetGrid
        .BackColorBkg = &HFFFFFF            ' White background
        .BackColorFixed = &H404040          ' Dark gray header background
        .ForeColorFixed = &HFFFFFF          ' White header text
        .ForeColor = &H0                    ' Black text
        .GridLines = flexGridFlat
        .RowHeightMin = 300
    End With

    rs.Close
    Set rs = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error loading data: " & Err.Description, vbCritical
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

