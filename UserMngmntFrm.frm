VERSION 5.00
Begin VB.Form UserMngmntFrm 
   Caption         =   "User Management"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbRole 
      Height          =   315
      ItemData        =   "UserMngmntFrm.frx":0000
      Left            =   2280
      List            =   "UserMngmntFrm.frx":0007
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox rePasswordTxt 
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox passwordTxt 
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox usernameTxt 
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton addBtn 
      Caption         =   "Save"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Role"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Retype Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "UserMngmntFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addBtn_Click()
    ' I-load ang database connection sa simula
    Call OpenConnection
    
    Dim user As New CreateUserDto
    user.UserNumber = "U000005"
    user.Username = usernameTxt(0).Text
    user.Password = passwordTxt(1).Text
    user.UserLevel = cbRole.Text
    user.Status = True
    user.IsSessionActive = False
    user.CurrentToken = ""
    user.RecordDate = Now

    Call InsertToTable("Users", user)

End Sub
