VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KITCHEN HOUSE CUSTOMER RECORD."
   ClientHeight    =   6390
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7755
   FillColor       =   &H0080FFFF&
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Search by name ! "
      Height          =   495
      Left            =   3000
      TabIndex        =   28
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   6360
      TabIndex        =   27
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Picture"
      Height          =   615
      Left            =   5400
      TabIndex        =   26
      Top             =   2160
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   65
      ImageHeight     =   55
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0422
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":081E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdpic 
      Caption         =   "Insert picture"
      Height          =   615
      Left            =   4320
      TabIndex        =   25
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton command3 
      Caption         =   "Save update"
      Height          =   855
      Left            =   600
      Picture         =   "Form1.frx":0D12
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OKAY"
      Height          =   855
      Left            =   6120
      TabIndex        =   23
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtpic 
      Height          =   285
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Phone no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Picture         =   "Form1.frx":0DC6
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtcompany 
      Height          =   405
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   20
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtnick 
      Height          =   375
      Left            =   4920
      MaxLength       =   15
      TabIndex        =   18
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtother 
      Height          =   375
      Left            =   960
      MaxLength       =   20
      TabIndex        =   16
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Picture         =   "Form1.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Picture         =   "Form1.frx":131F
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "Form1.frx":173F
      Left            =   6360
      List            =   "Form1.frx":1741
      TabIndex        =   12
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtMobile 
      Height          =   375
      Left            =   4200
      MaxLength       =   20
      TabIndex        =   11
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox txtPhone 
      Height          =   375
      Left            =   960
      MaxLength       =   15
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   4200
      MaxLength       =   25
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox txtBirth 
      Height          =   405
      Left            =   960
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtAddress 
      Height          =   1575
      Left            =   240
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   1800
      MaxLength       =   18
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   4080
      Picture         =   "Form1.frx":1743
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Customer Name."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label42 
      BackColor       =   &H00800000&
      Caption         =   "Company Tel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   "HP/PG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Caption         =   "Home Tel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3360
      MousePointer    =   10  'Up Arrow
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label32 
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Customer Ph no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Menu munfile 
      Caption         =   "&File"
      Begin VB.Menu munsave 
         Caption         =   "&save edit"
      End
      Begin VB.Menu munblank 
         Caption         =   "-"
      End
      Begin VB.Menu munclose 
         Caption         =   "Exit Program"
      End
   End
   Begin VB.Menu munadd 
      Caption         =   "&Add"
      Begin VB.Menu munData 
         Caption         =   "&Add Data"
      End
   End
   Begin VB.Menu munabt 
      Caption         =   "&About"
      Begin VB.Menu muninfor 
         Caption         =   "&information"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim DOB
Public MyData As Database
Public MyRecord As Recordset
Dim SQL As String
Public strfind As String
'Public mySQL As String
 


Private Sub cmdAdd_Click()

Load Form2
Form2.Show

    
    Form2.txtname.Text = ""
    Form2.txtaddress.Text = ""
    Form2.txtnick.Text = ""
    Form2.txtbirth.Text = ""
    Form2.txtemail.Text = ""
    Form2.txthp.Text = ""
    Form2.txtother.Text = ""
    Form2.txtname.SetFocus
    Form2.txtcompany = ""
    Form2.txthome.Text = ""
End Sub




Private Sub cmdEdit_Click()


Dim DOB

      
Set MyData = OpenDatabase(App.Path + "\db1.mdb")
Set MyRecord = MyData.OpenRecordset("book1")

    Do Until MyRecord.EOF
        If List1.Text = MyRecord!Name Then
            List1.RemoveItem (List1.ListIndex)
            List1.AddItem txtname.Text
            With MyRecord
                .Edit
                !Name = Trim(txtname.Text)
                !Address = Trim(txtaddress.Text)
               !birth = Trim(txtbirth.Text)
                !nick = Trim(txtnick.Text)
                !Email = Trim(txtemail.Text)
                !phone = Trim(txtPhone.Text)
                !Hp = Trim(txtMobile.Text)
                !others = Trim(txtother.Text)
                !company = Trim(txtcompany.Text)
                !pic = Trim(txtpic.Text)
 

               
             .Update
                
    End With
   
        End If
       
        MyRecord.MoveNext
    Loop
     
    
If txtbirth.Text <> "" Then
         
        If IsDate(txtbirth.Text) Then
            DOB = CDate(txtbirth.Text)

        Else
            MsgBox "The date of birth field is not a valid date,pls change or left is blank.", vbOKOnly, "date of birth"
            
            
        End If
End If
    
    
    
Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim Temp As String
 iCount = List1.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With List1
         If .List(i) > .List(i + 1) Then
            Temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = Temp
     End If
     End With
  
     
      Next i
Next j
  
   
   Command3.Enabled = False
cmdpic.Enabled = False

cmdRemove.Enabled = False
munsave.Enabled = False

    
End Sub





Private Sub cmdpic_Click()
Set MyData = OpenDatabase(App.Path + "\db1.mdb")
        Set MyRecord = MyData.OpenRecordset("book1")
On Error GoTo DialogError
With CommonDialog1
        .CancelError = True
        .Filter = "JPG File (*.jpg)|*.jpg|Bitmap File (*.bmp)|*.bmp|GIF File(*.gif)|*.gif|All Files(*.*)|*.*"
        .FilterIndex = 2
        .DialogTitle = "Select a Picture File"
        .ShowOpen
   Image1.Picture = LoadPicture(.FileName)
   txtpic.Text = .FileName
   
   End With
   
DialogError:
End Sub

Private Sub cmdRemove_Click()
Dim msg
msg = MsgBox("Do you want to remove this name", vbYesNo, "Delete")
If msg = vbYes Then

Set MyData = OpenDatabase(App.Path + "\db1.mdb")
SQL = "SELECT * FROM Book1"
Set MyRecord = MyData.OpenRecordset(SQL)
    Do Until MyRecord.EOF
        If List1.Text = MyRecord!Name Then
            MyRecord.Delete
            List1.RemoveItem (List1.ListIndex)
        End If
    MyRecord.MoveNext
   
    Loop


                txtname.Text = ""
                txtaddress.Text = ""
                txtnick.Text = ""
                txtbirth.Text = ""
                txtemail.Text = ""
                txtPhone.Text = ""
                txtMobile.Text = ""
                txtother.Text = ""
                 txtcompany.Text = ""
                 txtpic.Text = ""
                Image1.Picture = LoadPicture("")
                
                
                
Command3.Enabled = False
cmdpic.Enabled = False

cmdRemove.Enabled = False
munsave.Enabled = False

Else
    Exit Sub
    End If
    

Set MyData = OpenDatabase(App.Path + "\db1.mdb")
SQL = "SELECT * FROM Book1"
Set MyRecord = MyData.OpenRecordset(SQL)
    Do Until MyRecord.EOF
        If List2.Text = MyRecord!company Then
            MyRecord.Delete
            List2.RemoveItem (List2.ListIndex)
        End If
    MyRecord.MoveNext
   
    Loop


                txtname.Text = ""
                txtaddress.Text = ""
                txtnick.Text = ""
                txtbirth.Text = ""
                txtemail.Text = ""
                txtPhone.Text = ""
                txtMobile.Text = ""
                txtother.Text = ""
                 txtcompany.Text = ""
                 txtpic.Text = ""
                Image1.Picture = LoadPicture("")
                
                
                
Command3.Enabled = False
cmdpic.Enabled = False

cmdRemove.Enabled = False
munsave.Enabled = False


    Exit Sub
    
    
           
           
End Sub








Private Sub Command1_Click()

Dim Search As String
Dim where
Dim i
Dim AtLeastone
AtLeastone = False

Search = InputBox("Enter Name to be Found", " Search Phone no")
If Search = "" Then
MsgBox "NO PHONE NUMBER FOUND", vbCritical
Exit Sub
Else

For i = 0 To List1.ListCount - 1
List1.Selected(i) = False 'Not to hignlight item

where = InStr(List1.List(i), Search)
If where Then
List1.Selected(i) = True
AtLeastone = True
End If
Next
If Not AtLeastone Then
MsgBox "Not Name Found"
End If
End If

End Sub
       
       'If (strfind) <> "" Then
       ' List1.Text = strfind
    ' End If
            

   
    
        '    If strfind <> List1.Text Then
       ' MsgBox "The Name you type is NOT in the List"
       ' End If

Private Sub Command3_Click()
Dim DOB
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
                MsgBox "The birth field is not a valid date" & (Chr(10)) & " pls change or left is blank.", vbInformation, "date of birth "
                txtbirth.Text = Empty
            Set MyData = OpenDatabase(App.Path + "\db1.mdb")
            Set MyRecord = MyData.OpenRecordset("book1")
                Do Until MyRecord.EOF
                If List1.Text = MyRecord!Name Then
                     List1.RemoveItem (List1.ListIndex)
                     List1.AddItem txtname.Text
                   List2.RemoveItem (List2.ListIndex)
                   List2.AddItem txtcompany.Text
                    With MyRecord
                    .Edit
                    !Name = Trim(txtname.Text)
                    !Address = Trim(txtaddress.Text)
                  ' !birth = Trim(txtbirth.Text)
                    !nick = Trim(txtnick.Text)
                    !Email = Trim(txtemail.Text)
                    !phone = Trim(txtPhone.Text)
                    !Hp = Trim(txtMobile.Text)
                    !others = Trim(txtother.Text)
                    !company = Trim(txtcompany.Text)
                    !pic = Trim(txtpic.Text)
                    .Update
                     
                     End With
                        List2.AddItem txtcompany.Text
   
              
                                
                                End If
                     MyRecord.MoveNext
            Loop
        
        
        
        
        
 Dim iiCount As Integer
 Dim ii As Integer
 Dim jj As Integer
 Dim tempp As String
 iiCount = List1.ListCount
 For jj = 0 To iiCount - 2
   For ii = 0 To iiCount - 2
     With List1
         If .List(ii) > .List(ii + 1) Then
            tempp = .List(ii + 1)
            .List(ii + 1) = .List(ii)
            .List(ii) = tempp
     End If
     End With
  
     
      Next ii
Next jj
Command3.Enabled = False
cmdpic.Enabled = False


munsave.Enabled = False

cmdRemove.Enabled = False

              Exit Sub
                 End If
  End If
  
  
  
  
     Set MyData = OpenDatabase(App.Path + "\db1.mdb")
        Set MyRecord = MyData.OpenRecordset("book1")

    Do Until MyRecord.EOF
        If List1.Text = MyRecord!Name Then
            List1.RemoveItem (List1.ListIndex)
            List1.AddItem txtname.Text
            With MyRecord
                .Edit
                !Name = Trim(txtname.Text)
                !Address = Trim(txtaddress.Text)
                !birth = Trim(txtbirth.Text)
                !nick = Trim(txtnick.Text)
                !Email = Trim(txtemail.Text)
                !phone = Trim(txtPhone.Text)
                !Hp = Trim(txtMobile.Text)
                !others = Trim(txtother.Text)
                !company = Trim(txtcompany.Text)
                !pic = Trim(txtpic.Text)
                
                
                
                
                
                .Update
                
    End With
        End If
        MyRecord.MoveNext
    Loop
    
     
    

'End If
    
    
    
Dim iCount1 As Integer
 Dim x As Integer
 Dim z As Integer
 Dim temp1 As String
 iCount1 = List1.ListCount
 For z = 0 To iCount1 - 2
   For x = 0 To iCount1 - 2
     With List1
         If .List(x) > .List(x + 1) Then
            temp1 = .List(x + 1)
            .List(x + 1) = .List(x)
            .List(x) = temp1
     End If
     End With
  
     
      Next x
Next z
Command3.Enabled = False
cmdpic.Enabled = False
munsave.Enabled = False



cmdRemove.Enabled = False
Form1.Refresh
Label2.Refresh

Exit Sub

End Sub


Private Sub Command2_Click()
Close
Unload Form1
Unload Form2

End Sub

Private Sub Command4_Click()
MsgBox "Are you sure?", vbYesNo, "picture"
If vbYes Then

Image1.Picture = LoadPicture("")
txtpic.Text = ""

Else
Exit Sub
End If
End Sub

Private Sub Command5_Click()
Dim Search As String
Dim where
Dim i
Dim AtLeastone
AtLeastone = False

Search = InputBox("Enter Name to be Found", " Search names")
If Search = "" Then
MsgBox "NO PHONE NUMBER FOUND", vbCritical
Exit Sub
Else

For i = 0 To List2.ListCount - 1
List2.Selected(i) = False 'Not to hignlight item

where = InStr(List2.List(i), Search)
If where Then
List2.Selected(i) = True
AtLeastone = True
End If
Next
If Not AtLeastone Then
MsgBox "Not Name Found"
End If
End If

End Sub

Private Sub Form_Load()
Command3.Enabled = False
cmdpic.Enabled = False
munsave.Enabled = False
cmdRemove.Enabled = False


Set MyData = OpenDatabase(App.Path + "\db1.mdb")
Set MyRecord = MyData.OpenRecordset("book1")
    'MyRecord.MoveFirst
    Do Until MyRecord.EOF
        List1.AddItem MyRecord.Fields("name")
        'List2.AddItem MyRecord.Fields("company")
        MyRecord.MoveNext
    
    
    Loop
    
   
   
   
Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim Temp As String
 iCount = List1.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With List1
         If .List(i) > .List(i + 1) Then
            Temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = Temp
     End If
     End With
  
   
      Next i
        Next j
    Close
    Set MyData = OpenDatabase(App.Path + "\db1.mdb")
Set MyRecord = MyData.OpenRecordset("book1")
  '  MyRecord.MoveFirst
    Do Until MyRecord.EOF
        'List1.AddItem MyRecord.Fields("name")
        List2.AddItem MyRecord.Fields("company")
        MyRecord.MoveNext
    
    
    Loop
Dim iCount1 As Integer
 Dim i1 As Integer
 Dim j1 As Integer
 Dim temp1 As String
 iCount1 = List2.ListCount
 For j1 = 0 To iCount1 - 2
   For i1 = 0 To iCount1 - 2
     With List2
         If .List(i1) > .List(i1 + 1) Then
            temp1 = .List(i1 + 1)
            .List(i1 + 1) = .List(i1)
            .List(i1) = temp1
     End If
     End With
  
   
      Next i1
        Next j1
 
 Close
    
End Sub


Private Sub Label5_Click()
Dim Myemail As Long
    Myemail = Shell("start mailto:" + txtemail.Text, 0)
End Sub

Private Sub List1_Click()

 
 
 
On Error Resume Next
        Set MyData = OpenDatabase(App.Path + "\db1.mdb")
        Set MyRecord = MyData.OpenRecordset("book1")
        MyRecord.MoveFirst
        
    Do Until MyRecord.EOF
            If List1.Text = MyRecord!Name Then
                 txtname.Text = MyRecord!Name
            
                txtaddress.Text = MyRecord!Address
                txtnick.Text = MyRecord!nick
                txtbirth.Text = MyRecord!birth
                txtemail.Text = MyRecord!Email
                txtPhone.Text = MyRecord!phone
                txtMobile.Text = MyRecord!Hp
                txtother.Text = MyRecord!others
              txtcompany.Text = MyRecord!company
                txtpic.Text = MyRecord!pic
           If MyRecord!pic <> "" Then
          Image1.Picture = LoadPicture(MyRecord!pic)
            Else
            Image1.Picture = LoadPicture
            
            
                End If
                
             
            
            End If
            
        
           MyRecord.MoveNext
        
    
          MyRecord.Update
           
          
          
            
    Loop



Command3.Enabled = True

cmdpic.Enabled = True
munsave.Enabled = True


cmdRemove.Enabled = True

End Sub






Private Sub List2_Click()
    
On Error Resume Next
        Set MyData = OpenDatabase(App.Path + "\db1.mdb")
        Set MyRecord = MyData.OpenRecordset("book1")
         
       
        
    Do Until MyRecord.EOF
            If List2.Text = MyRecord!company Then
                 txtname.Text = MyRecord!Name
            
                txtaddress.Text = MyRecord!Address
                txtnick.Text = MyRecord!nick
                txtbirth.Text = MyRecord!birth
                txtemail.Text = MyRecord!Email
                txtPhone.Text = MyRecord!phone
                txtMobile.Text = MyRecord!Hp
                txtother.Text = MyRecord!others
              txtcompany.Text = MyRecord!company
                txtpic.Text = MyRecord!pic
           If MyRecord!pic <> "" Then
          Image1.Picture = LoadPicture(MyRecord!pic)
            Else
            Image1.Picture = LoadPicture
            
            
            
            End If
              
            End If
            
        
           MyRecord.MoveNext
        
    
          MyRecord.Update
           
          
          
            
   Loop



Command3.Enabled = True

cmdpic.Enabled = True
munsave.Enabled = True


cmdRemove.Enabled = True

End Sub

Private Sub munclose_Click()



Unload Form1


End Sub

Private Sub munData_Click()
Call cmdAdd_Click


    
   
    
End Sub

Private Sub munpicture_Click()
Set MyData = OpenDatabase(App.Path + "\db1.mdb")
        Set MyRecord = MyData.OpenRecordset("book1")
On Error GoTo DialogError
With CommonDialog1
        .CancelError = True
        .Filter = "All Files(*.*)|*.*|JPG File (*.jpg)|*.jpg|Bitmap File (*.bmp)|*.bmp|GIF File(*.gif)|*.gif"
        .FilterIndex = 1
        .DialogTitle = "Select a Picture File"
        .ShowOpen
   Image1.Picture = LoadPicture(.FileName)
   txtpic.Text = .FileName
   
   End With
   
   
   
DialogError:
End Sub

Private Sub muninfor_Click()
FAB.Show
FAB.SetFocus
End Sub

Private Sub munsave_Click()
Call Command3_Click



    
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

Case 1
    Call Command1_Click
Case 2
    Call cmdRemove_Click

Case 3
 Call cmdAdd_Click

Case 4
    Call Command3_Click
    End Select
    
 
    
End Sub



Private Sub txtMobile_KeyPress(KeyAscii As Integer)
'If (KeyAscii < (33) Or KeyAscii > (58)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtother_KeyPress(KeyAscii As Integer)
'If (KeyAscii < (33) Or KeyAscii > (58)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtphone_KeyPress(KeyAscii As Integer)
'If (KeyAscii < (33) Or KeyAscii > (58)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
