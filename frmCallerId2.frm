VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{B0E4B491-7E7B-11D0-8E50-444553540000}#4.0#0"; "ICONTRAY.OCX"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form frmCallerId2 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caller ID "
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFF00&
   Icon            =   "frmCallerId2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1800
      Width           =   4275
   End
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "&Say It!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSpeed 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   8520
      TabIndex        =   14
      Text            =   "150"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   4560
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4560
      Top             =   1080
   End
   Begin VB.CommandButton cmdsys 
      Caption         =   "HIDE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      MaskColor       =   &H00FFC0FF&
      Picture         =   "frmCallerId2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      MaskColor       =   &H00FFC0FF&
      Picture         =   "frmCallerId2.frx":170C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search by Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Address      Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Manual Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command123 
      Caption         =   "Get Ph no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cboPort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCallerId2.frx":1B4E
      Left            =   4200
      List            =   "frmCallerId2.frx":1B5E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdReleaseModem 
      Caption         =   "Release Modem"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdRecieveEvents 
      Caption         =   "Recieve Events"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "frmCallerId2.frx":1B6E
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      InputLen        =   1
   End
   Begin IconTrayOCX.IconTray IconTray1 
      Left            =   2280
      Top             =   3480
      _ExtentX        =   688
      _ExtentY        =   741
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS sp 
      Height          =   495
      Left            =   3240
      OleObjectBlob   =   "frmCallerId2.frx":1FB0
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MAIN CONTROL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Address book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      TabIndex        =   20
      Top             =   0
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "All events"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   21
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H008080FF&
      Caption         =   "Recent calls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4440
      TabIndex        =   22
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label say 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   8520
      Picture         =   "frmCallerId2.frx":2008
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Caption         =   "      Last Call:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   -240
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Caption         =   "Comm Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCallerId2"
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
'read ini files
Private Declare Function GetPrivateProfileString Lib "kernel32" _
          Alias "GetPrivateProfileStringA" ( _
          ByVal lpApplicationName As String, _
          ByVal lpKeyName As Any, _
          ByVal lpDefault As String, _
          ByVal lpReturnedString As String, _
          ByVal nSize As Long, _
          ByVal lpFileName As String) As Long
'write ini files

Private Declare Function WritePrivateProfileString Lib "kernel32" _
          Alias "WritePrivateProfileStringA" ( _
          ByVal lpApplicationName As String, _
          ByVal lpKeyName As Any, _
          ByVal lpString As Any, _
          ByVal lpFileName As String) As Long
          
  


'-------STOP WAVE SOUND-------
Public Function WriteINIString( _
          ByVal strSection As String, _
          ByVal strKeyName As String, _
          ByVal strValue As String, _
          ByVal strFile As String) As Long
  Dim lngStatus As Long
  Dim sNetworkUser As String

  lngStatus& = WritePrivateProfileString( _
            strSection, _
            strKeyName, _
            strValue, _
            strFile)
  WriteINIString& = (lngStatus& <> 0)
End Function

Public Function GetINIString( _
          ByVal strSection As String, _
          ByVal strKeyName As String, _
          ByVal strFile As String, _
          Optional ByVal strDefault As String = "") As String
  Dim strBuffer         As String * 256, lngSize As Long
  Dim sNetworkUser      As String

  lngSize& = GetPrivateProfileString( _
            strSection$, _
            strKeyName$, _
            strDefault$, _
            strBuffer$, _
            CLng(256), _
            strFile$)
  GetINIString$ = Left$(strBuffer$, lngSize&)
End Function


Public Sub logCall(ByVal sInfo As String, sFileName As String)
  Dim iFreeFile As Long
  Dim sData2Save As String
  
  sData2Save = sInfo
  'sData2Save = Replace(sData2Save, vbNewLine, " ")
  sData2Save = funcEscapeText(sData2Save)
  
  iFreeFile = FreeFile
  Open sFileName For Append As #iFreeFile
  Print #iFreeFile, sData2Save
  Close #iFreeFile
  txtLog.Text = txtLog.Text & sData2Save & vbNewLine
  'scroll to end
  txtLog.SelLength = 0
  txtLog.SelStart = Len(txtLog.Text)
End Sub

Public Sub logNameAndNumber(ByVal sInfo As String, sFileName As String)
  Dim iFreeFile As Long
  Dim sData2Save As String
  
  sData2Save = sInfo
  sData2Save = funcEscapeText(sData2Save)
  
  iFreeFile = FreeFile
  Open sFileName For Append As #iFreeFile
  Print #iFreeFile, sData2Save
  Close #iFreeFile
End Sub

Private Function funcEscapeText(sInput As String) As String
  Dim sDataToDisplay As String
  Dim iEachLetter As Long
  
  sDataToDisplay = sInput
  For iEachLetter = 0 To 31
    sDataToDisplay = Replace(sDataToDisplay, Chr(iEachLetter), "{\" & iEachLetter & "}")
  Next iEachLetter
  
  For iEachLetter = 127 To 144
    sDataToDisplay = Replace(sDataToDisplay, Chr(iEachLetter), "{\" & iEachLetter & "}")
  Next iEachLetter
  
  For iEachLetter = 147 To 255
    sDataToDisplay = Replace(sDataToDisplay, Chr(iEachLetter), "{\" & iEachLetter & "}")
  Next iEachLetter
  funcEscapeText = sDataToDisplay
End Function

Private Sub cboPort_Click()
  WriteINIString "settings", "port", cboPort.Text, App.Path & "\settings.ini"
End Sub

Private Sub cmdRecieveEvents_Click()
  On Error GoTo Connect_Click_Err
  MSComm1.CommPort = Val(cboPort.Text)
  If Not MSComm1.PortOpen Then                              ' Open the comm port if not already open
      MSComm1.PortOpen = True
  End If

  If Not MSComm1.PortOpen Then                              ' if there is a problem opening the port
      MsgBox "Cannot open comm port " & MSComm1.CommPort    ' display an error first
      Exit Sub                                                   ' bail out of the program
  End If

  ' Initialize communications and update app UI
  MSComm1.RThreshold = 1                                    ' Generate a receive event on every character received
  MSComm1.InputLen = 1                                      ' Read the receive buffer 1 char at a time
  
  ' Make sure that you send the correct Modem Command
  MSComm1.Output = "AT+CID=1" & vbCr                       ' Send command to put Identifier in event mode and receive serial number
  cmdRecieveEvents.Enabled = False
  cmdReleaseModem.Enabled = True

  Exit Sub

Connect_Click_Err:

  If Err.Number = 8005 Then
    cmdReleaseModem_Click
    MsgBox "Unable to connect to modem, port already open.", vbCritical, "Error"
  End If
End Sub

Private Sub cmdReleaseModem_Click()
  MSComm1.PortOpen = False
  cmdRecieveEvents.Enabled = True
  cmdReleaseModem.Enabled = False
End Sub

Private Sub cmdSpeak_Click()
sp.Speak say.Caption
    sp.Speed = txtSpeed.Text
   ' Sspeak = True
End Sub

Private Sub Command5_Click()
sp.Speak "welcome to kitchen house customer caller idd system"
End Sub

Private Sub txtSpeed_LostFocus()
    If txtSpeed.Text < 50 Then
        MsgBox "Speed is too low."
        txtSpeed.Text = "150"
    End If
    If txtSpeed.Text > 250 Then
        MsgBox "Speed is too high."
        txtSpeed.Text = "150"
    End If
End Sub

Private Sub Command1_Click()
Dim Search As String
Dim where
Dim i
Dim AtLeastone
AtLeastone = False
Search = InputBox("Enter Phone number to be Found!", " MANUAL PHONE NO SEARCH")
If Search = "" Then
MsgBox "YOU DIDN,T ENTER ANY PHONE NUMBER ?", vbCritical
Exit Sub
Else

For i = 0 To Form1.List1.ListCount - 1
Form1.List1.Selected(i) = False 'Not to hignlight item

where = InStr(Form1.List1.List(i), Search)
If where Then
Form1.List1.Selected(i) = True
AtLeastone = True
Form1.Show
End If
Next
If Not AtLeastone Then
Form1.Hide
MsgBox "Not Name Found"
End If
End If
End Sub

Private Sub Command123_Click()
Dim Search As String
Dim where
Dim i
Dim AtLeastone
AtLeastone = False

Search = lblName.Caption
If Search = "" Then
MsgBox "NO PHONE NUMBER FOUND", vbCritical
Exit Sub
Else

For i = 0 To Form1.List1.ListCount - 1
Form1.List1.Selected(i) = True 'Not to hignlight item

where = InStr(Form1.List1.List(i), Search)
If where Then
Form1.List1.Selected(i) = True
AtLeastone = False
Form2.txtname.Text = lblName.Caption
Form2.Show
Form1.Show
Form1.SetFocus
End If
Next
If Not AtLeastone Then
Form2.txtname.Text = lblName.Caption
Form2.Show
Form1.Show
Form1.SetFocus
'Form2.Hide
End If
End If


End Sub

Private Sub Command2_Click()
Form1.Show

End Sub

Private Sub Command3_Click()
Dim Search As String
Dim where
Dim i
Dim AtLeastone
AtLeastone = False

Search = InputBox("Enter Name to be Found", " Search by Names")
If Search = "" Then
MsgBox "NO NAME FOUND", vbCritical
Exit Sub
Else

For i = 0 To Form1.List2.ListCount - 1
Form1.List2.Selected(i) = False 'Not to hignlight item

where = InStr(Form1.List2.List(i), Search)
If where Then
Form1.List2.Selected(i) = True
AtLeastone = True
Form1.Show
End If
Next
If Not AtLeastone Then
MsgBox "Not Name Found"
End If
End If

End Sub


Private Sub Command4_Click()
Unload Form1
Unload Form2
Unload Me
End Sub

Private Sub Form_Activate()
frmCallerId2.Caption = "KITCHEN HOUSE AUTO CUSTOMER NAME CALLER ID SYSTEM " & Time$
  Timer1.Enabled = True
  
  'scroll to end
  txtLog.SelLength = 0
  txtLog.SelStart = Len(txtLog.Text)
 ' cmdRecieveEvents_Click
End Sub

'gets the text between two other strings
'Examples:
'  Debug.Print funcParseStringFromString2String("<start>data</end>", "<start>", "</end>")
'  '>>>data
'  Debug.Print funcParseStringFromString2String("<start>data</end>", "<Start>", "</end>", True)
'  '>>>data
'  Debug.Print funcParseStringFromString2String("<start>data</end>", "", "</end>")
'  '>>><start>data
'  Debug.Print funcParseStringFromString2String("<start>data</end>", "<stArt>", "", True)
'  '>>>data</end>
'  Debug.Print funcParseStringFromString2String("<start>data</end>", "<staRt>", "</eNd>")
'  '>>>(nothing was returned)
'  Debug.Print funcParseStringFromString2String("<start>data</end>", "", "")
'  >>><start>data</end>
Function funcParseStringFromString2String(sSourceString, sString1 As String, sString2 As String, Optional fCaseCaseInsensitive As Boolean = False) As String
  Dim sOutput As String
  Dim iLocationOfString1 As Long
  Dim iLocationOfString2 As Long
  Dim iCompareStyle As Long
  
  If fCaseCaseInsensitive Then
    iCompareStyle = vbTextCompare
  Else
    iCompareStyle = vbBinaryCompare
  End If
  
  sOutput = sSourceString
  iLocationOfString1 = InStr(1, sOutput, sString1, iCompareStyle)
  iLocationOfString2 = InStr(1, sOutput, sString2, iCompareStyle)
  If iLocationOfString1 = 0 And iLocationOfString2 = 0 Then
    'nothing found
    sOutput = ""
  Else
    If Len(sString1) = 0 And Len(sString2) = 0 Then
      'do nothing
    ElseIf Len(sString1) = 0 Then
      If iLocationOfString2 <> 0 Then
        sOutput = Mid(sOutput, 1, iLocationOfString2 - 1)
      End If
    ElseIf Len(sString2) = 0 Then
      sOutput = Mid(sOutput, iLocationOfString1 + Len(sString1))
    Else
      'cut off begining
      If iLocationOfString1 <> 0 Then
        sOutput = Mid(sOutput, iLocationOfString1 + Len(sString1))
      End If
      'take off the end part
      iLocationOfString2 = InStr(1, sOutput, sString2, iCompareStyle)
      If iLocationOfString2 <> 0 Then
        sOutput = Mid(sOutput, 1, iLocationOfString2 - 1)
      End If
    End If
  End If
  funcParseStringFromString2String = sOutput
End Function

Private Sub Form_Load()
  Call Command5_Click
  Dim iFreeFile As Long
  Dim sLogFileContent As String
  Dim sPort As String
  
  On Error Resume Next
  iFreeFile = FreeFile
  Open App.Path & "\call_data_log.txt" For Input As #iFreeFile
  sLogFileContent = StrConv(InputB(LOF(iFreeFile), iFreeFile), vbUnicode)
  Reset
  Me.txtLog.Text = sLogFileContent
  
  sPort = GetINIString("settings", "port", App.Path & "\settings.ini", "")
  If sPort <> "" Then
    cboPort.Text = sPort
  End If
 cmdRecieveEvents_Click
End Sub

Private Sub lblLastCallTime_Click()

End Sub

Private Sub Image1_Click()
FAB.Show
FAB.SetFocus

End Sub

Private Sub MSComm1_OnComm()
  Dim sNewCharacter               As String * 1                   'temporary storage for received comm port data
  Dim sDataChunk As String
  Static sSingleMessage As String
  Dim sNameAndNumber As String
  Dim iEachChar As Long
  Dim sCurrentChar As String
  Dim sNameString As String
  
  Select Case MSComm1.CommEvent
    Case OnCommConstants.comEvReceive                                       ' Received RThreshold # of chars.
      Call logCall(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & _
                "Received event.", _
                App.Path & "\call_data_log.txt")
      Do
        sNewCharacter = MSComm1.Input                         'read 1 character .Inputlen = 1
        sDataChunk = sDataChunk & sNewCharacter
      Loop While MSComm1.InBufferCount                      'Loop until all characters in receive buffer are processed
      For iEachChar = 1 To Len(sDataChunk)
        sCurrentChar = Mid(sDataChunk, iEachChar, 1)
        If sCurrentChar = Chr(10) Then
          'ignore
        ElseIf sCurrentChar = Chr(13) Then
          'end of a single message
          Beep
          If InStr(sSingleMessage, "NMBR = ") = 1 Then
            sNameString = sSingleMessage
            sNameString = Replace(sNameString, "NMBR = ", "")
            sNameString = Trim(sNameString)
            lblName.Caption = Format(sNameString, " ##########")
           say.Caption = Format(sNameString, " # # # # # # # # # #")
           List1.AddItem lblName.Caption ' = Format(sNameString, " # # # # # # # # # #")
            
            Call IconTray1_LeftDblClick
            Call cmdSpeak_Click
           Call Command123_Click
          ElseIf InStr(sSingleMessage, "NAME = ") = 1 Then
            sNameString = sSingleMessage
            sNameString = Replace(sNameString, "NAME = ", "")
            sNameString = Trim(sNameString)
            lblName.Caption = lblName.Caption & " " & sNameString
            Call logNameAndNumber(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & lblName.Caption, App.Path & "\numbers.txt")
           ' lblLastCallTime.Caption = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM")
          End If
          If Trim(sSingleMessage) <> "" Then
            logCall Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & sSingleMessage, App.Path & "\call_data_log.txt"
          End If
          sSingleMessage = ""
        Else
          sSingleMessage = sSingleMessage & sCurrentChar
        End If
      Next iEachChar
    Case OnCommConstants.comEvCD                                       ' Received RThreshold # of chars.
      Call logCall(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & _
                "Change in carrier detect line.", _
                App.Path & "\call_data_log.txt")
    Case OnCommConstants.comEvCTS                                       ' Received RThreshold # of chars.
    Call logCall(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & _
               "Change in clear-to-send line.", _
                App.Path & "\call_data_log.txt")
    Case OnCommConstants.comEvDSR                                       ' Received RThreshold # of chars.
      Call logCall(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & _
               "Change in data set ready line.", _
              App.Path & "\call_data_log.txt")
   Case OnCommConstants.comEvEOF                                       ' Received RThreshold # of chars.
    Call logCall(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & _
                "End of file.", _
               App.Path & "\call_data_log.txt")
    Case OnCommConstants.comEvRing                                       ' Received RThreshold # of chars.
     Call logCall(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & _
                "Ring detected.", _
              App.Path & "\call_data_log.txt")
   Case OnCommConstants.comEvSend                                       ' Received RThreshold # of chars.
      Call logCall(Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss AM/PM") & " " & _
                "Send event.", _
                App.Path & "\call_data_log.txt")
  End Select
End Sub


Private Sub cmdSys_Click()
IconTray1.Icon = Form2.Icon
IconTray1.ToolTip = "Waiting for call..."
IconTray1.Add
cmdsys.Enabled = False
Me.Visible = False
End Sub
Private Sub IconTray1_LeftDblClick()
Me.Show
IconTray1.Remove
cmdsys.Enabled = True
End Sub
Private Sub mnuShow_Click()
Me.Show
IconTray1.Remove
cmdsys.Enabled = True
End Sub

'Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)
'SystemTray1.Action = sys_delete
'SystemTray1.Action = sys_modify
'Me.Visible = True
'cmdsys.Enabled = True
'End Sub



Private Sub Timer1_Timer()
frmCallerId2.Caption = "KITCHEN HOUSE AUTO CUSTOMER NAME CALLER ID SYSTEM " & Time$
End Sub
