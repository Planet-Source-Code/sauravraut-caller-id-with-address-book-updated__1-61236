VERSION 5.00
Begin VB.Form FAB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Box"
   ClientHeight    =   4485
   ClientLeft      =   1020
   ClientTop       =   1425
   ClientWidth     =   6240
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer FSR_Check 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1140
   End
   Begin VB.CommandButton CommandOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   3960
      Width           =   800
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00B0D8E8&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   17
      Top             =   120
      Width           =   735
      Begin VB.PictureBox IconPicture 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "about.frx":000C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Caption         =   "denimboy01@yahoo.com"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label132 
      Caption         =   $"about.frx":044E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   23
      Top             =   1560
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin VB.Label OptLabel 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   45
      Index           =   9
      Left            =   900
      TabIndex        =   22
      Top             =   2940
      Width           =   4395
   End
   Begin VB.Label OptLabel 
      Caption         =   "%  in Use"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   3000
      TabIndex        =   21
      Top             =   4090
      Width           =   2055
   End
   Begin VB.Label OptLabel 
      Caption         =   "System Memory Load:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   960
      TabIndex        =   20
      Top             =   4090
      Width           =   1995
   End
   Begin VB.Label OptLabel 
      Caption         =   "Free Virtual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   19
      Top             =   3835
      Width           =   2055
   End
   Begin VB.Label OptLabel 
      Caption         =   "Virtual Memory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   960
      TabIndex        =   18
      Top             =   3835
      Width           =   1995
   End
   Begin VB.Label OptLabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   5
      Left            =   840
      TabIndex        =   16
      Top             =   2160
      Width           =   4395
   End
   Begin VB.Label OptLabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   2
      Left            =   840
      TabIndex        =   15
      Top             =   1440
      Width           =   4395
   End
   Begin VB.Label OptLabel 
      Caption         =   "Free Paging"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   14
      Top             =   3580
      Width           =   2055
   End
   Begin VB.Label OptLabel 
      Caption         =   "Paging Memory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   960
      TabIndex        =   13
      Top             =   3580
      Width           =   1995
   End
   Begin VB.Label OptLabel 
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   12
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label OptLabel 
      Caption         =   "Windows Type"
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
      Index           =   6
      Left            =   960
      TabIndex        =   11
      Top             =   2220
      Width           =   2535
   End
   Begin VB.Label OptLabel 
      Caption         =   "Free Physical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   10
      Top             =   3325
      Width           =   2055
   End
   Begin VB.Label OptLabel 
      Caption         =   "Physical Memory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   960
      TabIndex        =   9
      Top             =   3325
      Width           =   1995
   End
   Begin VB.Label OptLabel 
      Caption         =   "CPU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   8
      Top             =   3070
      Width           =   2055
   End
   Begin VB.Label OptLabel 
      Caption         =   "CPU Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   960
      TabIndex        =   7
      Top             =   3070
      Width           =   1995
   End
   Begin VB.Label OptLabel 
      Caption         =   "Build:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label NameLabel 
      Caption         =   "             KITCHEN HOUSE CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5235
   End
   Begin VB.Label OptLabel 
      Caption         =   "                                   041-520163"
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
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2955
   End
   Begin VB.Label OptLabel 
      Caption         =   "                                   SAURAV RAUT"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   4275
   End
   Begin VB.Label CoprLabel 
      Caption         =   " Copyright 2005     by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   540
      Width           =   4275
   End
End
Attribute VB_Name = "FAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' *******************************************************************
'
' Some code and files: 1996 by Gregory H. Bragg, SofTecH Development
'                      1995 by David Warren, MMC Software
' Some of the Registry code is from the VB4 Setup Kit, SETUP1 files.
'
' Originally published by PC Magazine. Ported over to
' 32 Bit VB4 by Gregory H. Bragg starting March 6, 1996
'
' Original: November 8, 1993  By Neil J. Rubenking
' Revised:  March 21, 1996    By Gregory H. Bragg
'
' To use the generic About Box defined in this file, your VBP file
' must also include the module ABOUTBOX.BAS. Just call the function
' DisplayAboutBox, passing parameters specific to your program.
' DO NOT load the form FAB prior to calling DisplayAboutBox!
'
' *******************************************************************
'
Option Explicit

Private Sub CommandOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim lResult As Long
    
    ' First let's centre the Icon picture inside the 3D raised container
    Dim tmp As Integer
    tmp = (SSPanel1.Width - IconPicture.Width) \ 2
    IconPicture.Left = tmp
    tmp = (SSPanel1.Height - IconPicture.Height) \ 2
    IconPicture.Top = tmp

    ' If both user and company are excluded, eliminate the second separator
    If ((Excl And AB_NO_USER) = &H1) And ((Excl And AB_NO_COMPANY) = &H2) Then
        EliminateLabel 2
    Else
        ' initialize some variables since we want either
        ' the user name or the company name or both...
        Dim szUserInfo As String, szSubkey As String
        Dim hKey As Long
        Dim Success As Boolean
        szSubkey = "SOFTWARE\Microsoft\Windows\CurrentVersion"
    End If
        
    ' Get the User name from the Registry, if wanted
    If Excl And AB_NO_USER Then
        EliminateLabel 3
    Else
        If (OSRegOpenKeyEx(HKEY_LOCAL_MACHINE, szSubkey, 0&, KEY_QUERY_VALUE, hKey)) = ERROR_SUCCESS Then
            Success = RegQueryStringValue(hKey, "RegisteredOwner", szUserInfo)
            Success = RegCloseKey(hKey)
           
        End If
    End If
    
    ' Get the Company name from the Registry, if wanted
    If Excl And AB_NO_COMPANY Then
        EliminateLabel 4
    Else
        If (OSRegOpenKeyEx(HKEY_LOCAL_MACHINE, szSubkey, 0&, KEY_QUERY_VALUE, hKey)) = ERROR_SUCCESS Then
            Success = RegQueryStringValue(hKey, "RegisteredOrganization", szUserInfo)
            Success = RegCloseKey(hKey)
            
        End If
    End If
    
    ' Show Windows version, if wanted
    If Excl And AB_NO_WIN_VERSION Then
        EliminateLabel 6
        EliminateLabel 7
        OptLabel(8).Visible = False
    Else
        Dim OSVer As OSVERSIONINFO
        OSVer.dwOSVersionInfoSize = Len(OSVer)
        lResult = GetVersionEx(OSVer)
        If lResult Then
            Select Case OSVer.dwPlatformId
                Case VER_PLATFORM_WIN32s
                    'NOTE: VB4/32 apps won't run on Win32s so this will never happen!
                    OptLabel(6).Caption = "Win32s Subsystem on Windows 3.xx"
                Case VER_PLATFORM_WIN32_WINDOWS
                    'NOTE: This value applies for all 32-bit non-NT Windows
                    '      versions, not necessarily just Windows 95
                    OptLabel(6).Caption = "Microsoft Windows 95"
                Case VER_PLATFORM_WIN32_NT
                    OptLabel(6).Caption = "Microsoft Windows NT"
            End Select
        End If
        ' Show Windows version number, if wanted
        If Excl And AB_NO_VERSION_NUMBER Then
            EliminateLabel 7
        Else
            OptLabel(7).Caption = "Version:  " & Format$(OSVer.dwMajorVersion) _
                                  & "." & Format$(OSVer.dwMinorVersion, "00")
        End If
        ' Show Windows build number, if wanted
        If Excl And AB_NO_BUILD_NUMBER Then
            OptLabel(8).Visible = False
        Else
            OptLabel(8).Caption = "Build:  " & Format$(OSVer.dwBuildNumber Mod 65536)
        End If
    End If
  
    ' Show CPU Type, if wanted
    If Excl And AB_NO_CPU Then
        EliminateLabel 10
        OptLabel(11).Visible = False
    Else
        Dim SysInfo As SYSTEM_INFO
        Dim CPU_Name As String
        Call GetSystemInfo(SysInfo)
        Select Case SysInfo.dwProcessorType
            Case PROCESSOR_INTEL_386
                CPU_Name = "Intel 386"
            Case PROCESSOR_INTEL_486
                CPU_Name = "Intel 486"
            Case PROCESSOR_INTEL_PENTIUM
                CPU_Name = "Pentium"
            Case PROCESSOR_MIPS_R2000
                CPU_Name = "Mips R2000"
            Case PROCESSOR_MIPS_R3000
                CPU_Name = "Mips R3000"
            Case PROCESSOR_MIPS_R4000
                CPU_Name = "Mips R4000"
            Case PROCESSOR_ALPHA_21064
                CPU_Name = "Alpha 21064"
            Case Else ' default if not defined...
                CPU_Name = Format$(SysInfo.dwProcessorType)
        End Select
        OptLabel(11).Caption = Format$(SysInfo.dwNumberOfProcessors) _
                               & "  " & CPU_Name & "  Processor"
    End If
    
    ' Let's enable the Timer control and call GlobalMemoryStatus()
    ' only if we are going to display the available memory status...
    If (((Excl And AB_NO_PHYSICAL) = &H40) And ((Excl And AB_NO_PAGING) = &H80) _
        And ((Excl And AB_NO_VIRTUAL) = &H100) And ((Excl And AB_NO_MEMLOAD) = &H200)) _
        = False Then
        FSR_Check.Enabled = True    'enable the Timer control
    Else
        If Excl And AB_NO_CPU Then  'eliminate the third separator
            EliminateLabel 9
        End If
    End If
    
    ' Show Physical Memory, if wanted
    If Excl And AB_NO_PHYSICAL Then
        EliminateLabel 12
        OptLabel(13).Visible = False
    End If
  
    ' Show Paging Memory, if wanted
    If Excl And AB_NO_PAGING Then
        EliminateLabel 14
        OptLabel(15).Visible = False
    End If

    ' Show Virtual Memory, if wanted
    If Excl And AB_NO_VIRTUAL Then
        EliminateLabel 16
        OptLabel(17).Visible = False
    End If

    ' Show Memory Load, if wanted
    If Excl And AB_NO_MEMLOAD Then
        EliminateLabel 18
        OptLabel(19).Visible = False
    End If
    
End Sub

'
' Let's check the memory status every Timer interval since the
' information returned is volatile, and there is no guarantee
' that two sequential calls to this function will return the
' same information...
'
Private Sub FSR_Check_Timer()
    
    Dim MemStat As MEMORYSTATUS
    Dim MemData As Long
    MemStat.dwLength = Len(MemStat)
    Call GlobalMemoryStatus(MemStat)
    
    ' Show Physical Memory, if wanted
    If (Excl And AB_NO_PHYSICAL) = False Then
        MemData = MemStat.dwAvailPhys
        If MemData <= 1024 Then
            OptLabel(13) = Format$(MemData) & "  Bytes Free"
        Else
            OptLabel(13) = Format$(MemData \ 1024, "###,###,###") & "  KB Free"
        End If
    End If
  
    ' Show Paging Memory, if wanted
    If (Excl And AB_NO_PAGING) = False Then
        MemData = MemStat.dwAvailPageFile
        If MemData <= 1024 Then
            OptLabel(15) = Format$(MemData) & "  Bytes Free"
        Else
            OptLabel(15) = Format$(MemData \ 1024, "###,###,###") & "  KB Free"
        End If
    End If

    ' Show Virtual Memory, if wanted
    If (Excl And AB_NO_VIRTUAL) = False Then
        MemData = MemStat.dwAvailVirtual
        If MemData <= 1024 Then
            OptLabel(17) = Format$(MemData) & "  Bytes Free"
        Else
            OptLabel(17) = Format$(MemData \ 1024, "###,###,###") & "  KB Free"
        End If
    End If

    ' Show Memory Load, if wanted
    If (Excl And AB_NO_MEMLOAD) = False Then
        MemData = MemStat.dwMemoryLoad
        OptLabel(19) = Format$(MemData) & " %  in Use"
    End If

End Sub


