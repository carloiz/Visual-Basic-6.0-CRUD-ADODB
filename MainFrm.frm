VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainFrm 
   BackColor       =   &H8000000C&
   Caption         =   "JCB System"
   ClientHeight    =   5400
   ClientLeft      =   6495
   ClientTop       =   3495
   ClientWidth     =   10530
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton Command2 
         Caption         =   "Users"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Load UsersListFrm
    UsersListFrm.Show
End Sub
