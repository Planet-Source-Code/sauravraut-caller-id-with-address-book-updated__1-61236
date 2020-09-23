VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding data"
   ClientHeight    =   5055
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Picture         =   "Form2.frx":082D
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Picture         =   "Form2.frx":0A5C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtemail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   17
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtother 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txthp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   14
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txthome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtbirth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtnick 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MaxLength       =   20
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox txtcompany 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00800000&
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Company Telephone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   "Hp/PG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Caption         =   "Home telephone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "Last date of order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Customer   ph no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu munfile 
      Caption         =   "&File"
      Begin VB.Menu munsave 
         Caption         =   "save"
      End
      Begin VB.Menu munshow 
         Caption         =   "show information"
      End
      Begin VB.Menu munblank 
         Caption         =   "-"
      End
      Begin VB.Menu munclose 
         Caption         =   "&close"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim Temp As String

 
  Option Compare Text

 


Private Sub cmdOK_Click()
Set MyData = OpenDatabase(App.Path + "\db1.mdb")
Set MyRecord = MyData.OpenRecordset("book1")
If txtname.Text = "" Then
    If MsgBox("Name Field is needed." & (Chr(10)) & "Please enter your name.", vbInformation, "Enter Value") = vbOK Then
        txtname.SetFocus
        Exit Sub
    End If
End If




If txtbirth.Text <> "" Then
         If IsDate(txtbirth.Text) Then
            DOB = CDate(txtbirth.Text)

        Else
            MsgBox "The birth field is not a valid date" & (Chr(10)) & " pls change or left is blank.", vbOKOnly, "date of birth"
            txtbirth.Text = Empty
            
            Exit Sub
            
            End If
    End If
With MyRecord
        .AddNew
        
        !Name = Trim(txtname.Text)
        !Address = Trim(txtaddress.Text)
        !birth = Trim(txtbirth.Text)
        !nick = Trim(txtnick.Text)
        !Email = Trim(txtemail.Text)
        !phone = Trim(txthome.Text)
        !Hp = Trim(txthp.Text)
        !others = Trim(txtother.Text)
        !company = Trim(txtcompany.Text)
        .Update
        
    End With
  Form1.List1.AddItem txtname.Text
  Form1.List2.AddItem txtcompany.Text
   
   iCount = Form1.List1.ListCount
    For j = 0 To iCount - 2
    For i = 0 To iCount - 2
With Form1.List1
    If .List(i) > .List(i + 1) Then
            Temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = Temp
     End If
     End With
  
     
      Next i
        Next j
    
     
   
    Close
   Unload Form1
   Form1.Show
    Unload Me
 Form1.SetFocus
End Sub



Private Sub Command1_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub Command2_Click()
msg = MsgBox("Are you sure you want to clear", vbYesNo + vbQuestion, "clear data")
If msg = vbYes Then
    txtname.Text = ""
    txtaddress.Text = ""
    txtnick.Text = ""
    txtbirth.Text = ""
    txtemail.Text = ""
    txthome.Text = ""
    txthp.Text = ""
    txtother.Text = ""
    txtcompany.Text = ""
    Else
    Exit Sub
    
    
    End If
   Command2.Enabled = False
End Sub

Private Sub Form_Load()
    txtname.Text = ""
    txtaddress.Text = ""
    txtnick.Text = ""
    txtbirth.Text = ""
    txtemail.Text = ""
    txthome.Text = ""
    txthp.Text = ""
    txtother.Text = ""
    txtcompany.Text = ""
    Command2.Enabled = False
    
    
End Sub

Private Sub munclose_Click()
Call Command1_Click

End Sub

Private Sub munsave_Click()
Call cmdOK_Click

End Sub

Private Sub munsearch_Click()
Form3.Show
Form2.Hide

End Sub

Private Sub munshow_Click()
Form1.Show
Form2.Hide
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
Command2.Enabled = True
End Sub

Private Sub txtbirth_KeyPress(KeyAscii As Integer)
Command2.Enabled = True
End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
Command2.Enabled = True
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
Command2.Enabled = True
End Sub

Private Sub txthome_KeyPress(KeyAscii As Integer)
'If (KeyAscii < (33) Or KeyAscii > (58)) And Not KeyAscii = 8 Then KeyAscii = 0
Command2.Enabled = True
End Sub

'Private Sub txthp_KeyPress(KeyAscii As Integer)
'If (KeyAscii < (33) Or KeyAscii > (58)) And Not KeyAscii = 8 Then KeyAscii = 0
'Command2.Enabled = True
'End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Command2.Enabled = True

End Sub

Private Sub txtnick_KeyPress(KeyAscii As Integer)
Command2.Enabled = True
End Sub

Private Sub txtother_KeyPress(KeyAscii As Integer)
'If (KeyAscii < (33) Or KeyAscii > (58)) And Not KeyAscii = 8 Then KeyAscii = 0
Command2.Enabled = True
End Sub
