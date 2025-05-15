Attribute VB_Name = "Connection"
' Module: modDatabase
Option Explicit
Public conn As ADODB.Connection

Dim Odbc As String

Dim NpgSql As String


' Tawagin ito sa main form load para buksan ang connection
Public Sub OpenConnection()

    Odbc = "Driver={MySql ODBC 8.0 ANSI Driver};server=localhost;database=jcbbase;uid=root;password=carlo;port=3306;"
    NpgSql = "Driver={PostgreSQL Unicode};Server=localhost;Port=8000;Database=JCBBase;Uid=postgres;Pwd=1234;"
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = NpgSql
    conn.Open
End Sub
