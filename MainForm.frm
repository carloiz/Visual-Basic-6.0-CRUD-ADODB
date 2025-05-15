VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   0  'None
   Caption         =   "JCB System"
   ClientHeight    =   7530
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab Frame1 
      Height          =   4215
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "MainForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "MainForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "MainForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   555
         Left            =   2760
         TabIndex        =   4
         Top             =   0
         Width           =   3495
      End
      Begin VB.CommandButton usersBtn 
         Caption         =   "Users"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton exitBtn 
         Caption         =   "X"
         Height          =   615
         Left            =   13800
         TabIndex        =   1
         Top             =   0
         Width           =   615
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    LoadFormInTab "ProductListFrm"
End Sub

Private Sub Form_Load()
    Frame1.Visible = False
End Sub



Private Sub exitBtn_Click()
    Unload Me
End Sub

Private Sub usersBtn_Click()
    ' Load UsersListFrm dynamically inside the selected tab of SSTab1
    LoadFormInTab "UsersListFrm"
End Sub

Private Sub LoadFormInTab(FormName As String)
    ' Unload any form currently loaded in the container for the selected tab
    On Error Resume Next
    Unload UsersListFrm
    Unload ProductListFrm ' Optional: If you're going to load another form later
    On Error GoTo 0

    ' Dynamically load the form into the container
    LoadForm FormName

    ' Position the form dynamically inside the selected tab
    If FormName = "UsersListFrm" Then
        With UsersListFrm
            .BorderStyle = 0
            .Caption = "" ' Remove title

            ' Position the form inside the container on SSTab1 (e.g., Frame1 for Tab(0))
            .Top = Frame1.Top
            .Left = Frame1.Left
            .Width = Frame1.Width
            .Height = Frame1.Height

            ' Show the form modeless (non-blocking)
            .Show vbModeless
        End With
    ElseIf FormName = "ProductListFrm" Then
        With ProductListFrm
            .BorderStyle = 0
            .Caption = "" ' Remove title

            ' Position the form inside the container on SSTab1 (e.g., Frame1 for Tab(0))
            .Top = Frame1.Top
            .Left = Frame1.Left
            .Width = Frame1.Width
            .Height = Frame1.Height

            ' Show the form modeless (non-blocking)
            .Show vbModeless
        End With
    End If
End Sub

Private Sub LoadForm(FormName As String)
    ' Dynamically load the form
    Select Case FormName
        Case "UsersListFrm"
            Load UsersListFrm
        Case "ProductListFrm"
            Load ProductListFrm
        ' Add more cases if needed
        Case Else
            MsgBox "Form not found: " & FormName
    End Select
End Sub

Private Sub SSTab1_TabClick(ByVal Index As Integer)
    ' Unload any previously loaded form when tab is switched
    On Error Resume Next
    Unload UsersListFrm
    Unload ProductListFrm
    On Error GoTo 0
    
    ' Load form based on the selected tab index
    Select Case Index
        Case 0 ' Tab(0) - Users List
            LoadFormInTab "UsersListFrm"
        Case 1 ' Tab(1) - Product List (Example)
            LoadFormInTab "ProductListFrm"
        ' Add more cases if you have more tabs/forms
    End Select
End Sub


