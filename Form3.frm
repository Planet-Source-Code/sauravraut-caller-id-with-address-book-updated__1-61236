VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton playwavesound 
      Caption         =   "Command1"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10
    '----------PLAY WAVE SOUND--------


Private Sub playwavesound_Click()
    soundfile$ = "c:\ring.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    HaHa = sndPlaySound(soundfile$, wFlags%)
End Sub
'-------STOP WAVE SOUND-------

