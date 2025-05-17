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
    targetGrid.Cols = rs.fields.Count
    targetGrid.FixedRows = 0
    targetGrid.FixedCols = 0

    ' Populate header names
    For col = 0 To rs.fields.Count - 1
        targetGrid.TextMatrix(0, col) = rs.fields(col).Name
    Next col

    ' Auto-fit column width to fill the grid
    Dim totalWidth As Long
    totalWidth = targetGrid.Width - 88 ' Adjust margin if needed

    Dim colWidth As Long
    colWidth = totalWidth \ rs.fields.Count

    For col = 0 To rs.fields.Count - 1
        targetGrid.colWidth(col) = colWidth
    Next col

    ' Populate rows
    row = 1
    Do Until rs.EOF
        targetGrid.AddItem ""
        For col = 0 To rs.fields.Count - 1
            Dim val As String
            If Not IsNull(rs.fields(col).value) Then
                val = CStr(rs.fields(col).value)
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


Public Sub InsertToTable(ByVal tableName As String, ByVal dto As Object)
    Dim cmd As New ADODB.Command
    Dim tliApp As New TLI.TLIApplication
    Dim typeInfo As TLI.InterfaceInfo
    Dim member As TLI.MemberInfo
    Dim fieldList As String
    Dim fields() As String
    Dim i As Integer

    ' Get the type information
    Set typeInfo = tliApp.InterfaceInfoFromObject(dto)

    ' Build field list from property names
    For Each member In typeInfo.Members
        If member.InvokeKind = INVOKE_PROPERTYGET Then
            fieldList = fieldList & member.Name & ","
        End If
    Next

    If Len(fieldList) = 0 Then
        MsgBox "No properties found in DTO.", vbCritical
        Exit Sub
    End If

    ' Remove trailing comma
    fieldList = Left(fieldList, Len(fieldList) - 1)
    fields = Split(fieldList, ",")

    ' Open connection if needed
    If conn Is Nothing Then
        MsgBox "Database connection is not initialized.", vbCritical
        Exit Sub
    End If

    If conn.State = adStateClosed Then
        conn.Open
    End If

    ' Setup command
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText

        ' Build SQL
        Dim sql As String
        sql = "INSERT INTO " & tableName & " (" & Join(fields, ",") & ") VALUES ("
        For i = 0 To UBound(fields)
            sql = sql & "?"
            If i < UBound(fields) Then sql = sql & ","
        Next i
        sql = sql & ")"
        .CommandText = sql

        ' Add parameters
        For i = 0 To UBound(fields)
            Dim field As String
            Dim val As Variant
            field = Trim(fields(i))
            val = CallByName(dto, field, VbGet)

            Select Case VarType(val)
                Case vbBoolean: .Parameters.Append .CreateParameter(, adBoolean, adParamInput, , val)
                Case vbDate: .Parameters.Append .CreateParameter(, adDate, adParamInput, , val)
                Case vbCurrency: .Parameters.Append .CreateParameter(, adCurrency, adParamInput, , val)
                Case vbDouble, vbSingle: .Parameters.Append .CreateParameter(, adDouble, adParamInput, , val)
                Case vbInteger, vbLong: .Parameters.Append .CreateParameter(, adInteger, adParamInput, , val)
                Case Else: .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 255, CStr(val))
            End Select
        Next i

        .Execute
    End With

    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub

