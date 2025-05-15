Attribute VB_Name = "Connection"
' Module: modDatabase
Option Explicit
Public conn As ADODB.Connection


Private Function NpgSql() As String
    NpgSql = "Driver={PostgreSQL Unicode};Server=localhost;Port=8000;Database=JCBBase;Uid=postgres;Pwd=1234;"
End Function

Private Function Odbc() As String
   Odbc = "Driver={MySql ODBC 8.0 ANSI Driver};server=localhost;database=jcbbase;uid=root;password=carlo;port=3306;"
End Function

' Tawagin ito sa main form load para buksan ang connection
Public Sub OpenConnection()
    Set conn = New ADODB.Connection
    conn.ConnectionString = NpgSql()
    conn.Open
End Sub
