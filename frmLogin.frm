VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1590
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   3870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   939.424
   ScaleMode       =   0  'User
   ScaleWidth      =   3633.72
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   600
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00FF0000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "$"
      TabIndex        =   0
      Top             =   600
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmLogin.frx":0442
      TabIndex        =   1
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESS CODE REQUIRED !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtpassword = "data" Then
       
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
               frmCallerId2.Show
  
  Unload Me
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtpassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Command1_Click()
FAB.Show
FAB.SetFocus

End Sub

Private Sub Form_Load()
Timer1_Timer
End Sub

Private Sub Timer1_Timer()
Label3.Visible = Not (Label3.Visible)
End Sub
