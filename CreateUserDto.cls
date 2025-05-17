VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreateUserDto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=== clsUserCreateDto ===
Private m_UserNumber As String
Private m_Username As String
Private m_Password As String
Private m_UserLevel As String
Private m_Status As Boolean
Private m_IsSessionActive As Boolean
Private m_CurrentToken As String
Private m_RecordDate As Date

' -- UserNumber
Public Property Get UserNumber() As String
    UserNumber = m_UserNumber
End Property

Public Property Let UserNumber(ByVal value As String)
    m_UserNumber = value
End Property

' -- Username
Public Property Get Username() As String
    Username = m_Username
End Property

Public Property Let Username(ByVal value As String)
    m_Username = value
End Property

' -- Password
Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let Password(ByVal value As String)
    m_Password = value
End Property

' -- UserLevel
Public Property Get UserLevel() As String
    UserLevel = m_UserLevel
End Property

Public Property Let UserLevel(ByVal value As String)
    m_UserLevel = value
End Property

' -- Status
Public Property Get Status() As Boolean
    Status = m_Status
End Property

Public Property Let Status(ByVal value As Boolean)
    m_Status = value
End Property

' -- IsSessionActive
Public Property Get IsSessionActive() As Boolean
    IsSessionActive = m_IsSessionActive
End Property

Public Property Let IsSessionActive(ByVal value As Boolean)
    m_IsSessionActive = value
End Property

' -- CurrentToken
Public Property Get CurrentToken() As String
    CurrentToken = m_CurrentToken
End Property

Public Property Let CurrentToken(ByVal value As String)
    m_CurrentToken = value
End Property

' -- RecordDate
Public Property Get RecordDate() As Date
    RecordDate = m_RecordDate
End Property

Public Property Let RecordDate(ByVal value As Date)
    m_RecordDate = value
End Property

