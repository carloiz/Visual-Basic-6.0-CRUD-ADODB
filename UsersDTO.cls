VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsersDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: UsersDto

Private pUserNumber As String
Private pUsername As String
Private pUserLevel As String
Private pStatus As String
Private pIsSessionActive As Boolean
Private pRecordDate As Date

' === UserNumber ===
Public Property Get UserNumber() As String
    UserNumber = pUserNumber
End Property

Public Property Let UserNumber(ByVal value As String)
    pUserNumber = value
End Property

' === Username ===
Public Property Get Username() As String
    Username = pUsername
End Property

Public Property Let Username(ByVal value As String)
    pUsername = value
End Property

' === UserLevel ===
Public Property Get UserLevel() As String
    UserLevel = pUserLevel
End Property

Public Property Let UserLevel(ByVal value As String)
    pUserLevel = value
End Property

' === Status ===
Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Let Status(ByVal value As String)
    pStatus = value
End Property

' === IsSessionActive ===
Public Property Get IsSessionActive() As Boolean
    IsSessionActive = pIsSessionActive
End Property

Public Property Let IsSessionActive(ByVal value As Boolean)
    pIsSessionActive = value
End Property

' === RecordDate ===
Public Property Get RecordDate() As Date
    RecordDate = pRecordDate
End Property

Public Property Let RecordDate(ByVal value As Date)
    pRecordDate = value
End Property

