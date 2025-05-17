VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form UserListFrm 
   BackColor       =   &H80000003&
   Caption         =   "User Management"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   13680
   Begin VB.CommandButton AddBtn 
      Caption         =   "Add New"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "UserListFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addBtn_Click()
    With UserMngmntFrm
        .Top = Top + (Height - .Height) \ 2
        .Left = Left + (Width - .Width) \ 2
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
    ' I-load ang database connection sa simula
    Call OpenConnection
    
    ' Sample Query
    Dim sql As String
    sql = "SELECT UserNumber, Username, UserLevel, Status, IsSessionActive, RecordDate FROM Users"

    Call LoadToGrid(sql, MSFlexGrid1)
End Sub

