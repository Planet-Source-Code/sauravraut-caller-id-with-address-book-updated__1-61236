Attribute VB_Name = "AboutBox_Main"
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
' Original: October 6, 1993  By Neil J. Rubenking
' Revised:  March 17, 1996   By Gregory H. Bragg
'
' Any program that includes this file must also include ABOUTBOX.FRM
'
' *******************************************************************
'
Option Explicit
'
' The AB_NO_xxxx constants are used to exclude informational lines
' from the About Box display.  You pass one or more of them, combined
' using OR, as the fourth last parameter to DisplayAboutBox.
Public Const AB_NO_USER = &H1
Public Const AB_NO_COMPANY = &H2
Public Const AB_NO_WIN_VERSION = &H4
Public Const AB_NO_VERSION_NUMBER = &H8
Public Const AB_NO_BUILD_NUMBER = &H10
Public Const AB_NO_CPU = &H20
Public Const AB_NO_PHYSICAL = &H40
Public Const AB_NO_PAGING = &H80
Public Const AB_NO_VIRTUAL = &H100
Public Const AB_NO_MEMLOAD = &H200

' Public variable holds bit flags for excluded items
Public Excl As Integer

' GetSystemMetrics returns the size (in pixels) of various on-screen
' items.  There are many more SM_xxxx constants besides those defined
' below.  The About Box uses the sizes to set its position on screen.
Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
Public Const SM_CYCAPTION = &H4
Public Const SM_CYMENU = &HF
Public Const SM_CXSIZE = &H1F

' O/S Version Info structure
' Used to get the operating system version and platform information
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type
' dwPlatformId defines for OSVERSIONINFO structure...
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
' and related Win API call...
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' System Info structure
' Used to get the amount and type of CPU information
Type SYSTEM_INFO
    dwOemID                     As Long
    dwPageSize                  As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask       As Long
    dwNumberOfProcessors        As Long
    dwProcessorType             As Long
    dwAllocationGranularity     As Long
    dwReserved                  As Long
End Type
' dwProcessorType defines for SYSTEM_INFO structure...
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R2000 = 2000
Public Const PROCESSOR_MIPS_R3000 = 3000
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064
' and related Win API call...
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

' Memory Status Info structure
' Used to get various system memory information
Type MEMORYSTATUS
    dwLength        As Long  ' sizeof(MEMORYSTATUS)
    dwMemoryLoad    As Long  ' percent of memory in use (between 1 and 100)
    dwTotalPhys     As Long  ' bytes of physical memory
    dwAvailPhys     As Long  ' free physical memory bytes
    dwTotalPageFile As Long  ' bytes of paging file
    dwAvailPageFile As Long  ' free bytes of paging file
    dwTotalVirtual  As Long  ' user bytes of address space
    dwAvailVirtual  As Long  ' free user bytes
End Type
' and related Win API call...
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

' Registry manipulation API's for getting the User or Company name
Declare Function OSRegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1


Public Sub DisplayAboutBox(F As Form, ByVal ProgName As String, _
                           ByVal Version, ByVal CoprDate, _
                           ByVal CoprName As String, _
                           ByVal Ex1 As String, ByVal Ex2 As String, _
                           ByVal Exclude As Integer, ByVal Center As Boolean, _
                           ByVal Fore As Long, ByVal Back As Long)
                    
' Your program simply calls this function to display an about box.
' F         - the main form of the calling program, used to get an
'             icon for display and to position the about box.
' ProgName  - program name, for caption and first line
' Version   - version number, displayed as 0.00
' CoprDate  - copyright year
' CoprName  - copyright holder's name
' Ex1       - extra data line 1 (optional)
' Ex2       - extra data line 2 (optional)
' Exclude   - used to exclude info from the about box.  AB_NO_xxxx
'             constants are bit-flags for this parameter.  e.g. to
'             exclude displaying User info and Company info,
'             pass AB_NO_USER OR AB_NO_COMPANY
' Center    - if TRUE, About box is centered on screen; if FALSE, About
'             box is displayed offset from calling window.
' Fore,Back - foreground and background colors for box; 0 to use default

  Screen.MousePointer = 11 'hourglass
  Excl = Exclude
  Load FAB
  Dim N As Integer
  If Fore Then
    FAB.ForeColor = Fore
    FAB.CoprLabel.ForeColor = Fore
    FAB.NameLabel.ForeColor = Fore
    FAB.SSPanel1.ForeColor = Fore
    For N = 0 To 19
      FAB.OptLabel(N).ForeColor = Fore
    Next N
  End If
  If Back Then
    FAB.BackColor = Back
    FAB.CommandOK.BackColor = Back
    FAB.CoprLabel.BackColor = Back
    FAB.IconPicture.BackColor = Back
    FAB.NameLabel.BackColor = Back
    FAB.SSPanel1.BackColor = Back
    For N = 0 To 19
      FAB.OptLabel(N).BackColor = Back
    Next N
  End If
  If Center Then
    FAB.Left = (Screen.Width - FAB.Width) \ 2
    FAB.Top = (Screen.Height - FAB.Height) \ 2
  Else
    ' Place the About box over the calling window, offset downward
    ' and to the right
    Dim tmp As Integer ' variable to keep lines of code from becoming TOO long
    tmp = GetSystemMetrics(SM_CXSIZE)
    FAB.Left = F.Left + tmp * Screen.TwipsPerPixelX
    tmp = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU)
    FAB.Top = F.Top + tmp * Screen.TwipsPerPixelY
    ' If about box now extends off the screen, move it back ON
    If FAB.Left + FAB.Width > Screen.Width Then
      FAB.Left = Screen.Width - (FAB.Width + 50)
    End If
    If FAB.Top + FAB.Height > Screen.Height Then
      FAB.Top = Screen.Height - (FAB.Height + 50)
    End If
  End If
  FAB.IconPicture.Picture = F.Icon
  FAB.Caption = "About " & ProgName
  Dim Temp As String ' variable to keep lines of code from becoming TOO long
  Temp = ProgName & ", Version " & Format$(Version, "0.00")
  FAB.NameLabel.Caption = Temp
  Temp = "Copyright © " & CoprDate & " by " & CoprName
  FAB.CoprLabel.Caption = Temp
  If Ex1 = "" Then
    EliminateLabel 0
  Else
    FAB.OptLabel(0).Caption = Ex1
  End If
  If Ex2 = "" Then
    EliminateLabel 1
  Else
    FAB.OptLabel(1).Caption = Ex2
  End If
  Screen.MousePointer = 0 'default
  FAB.Show vbModal
  
End Sub

Public Sub EliminateLabel(ByVal Which As Long)

    ' If one of the informational labels in the about box is not wanted,
    ' make it invisible and move all the other labels up to fill in the
    ' space.  Then shrink the form as well.
  
    FAB.OptLabel(Which).Visible = False
    Dim N As Integer, H As Integer
    H = FAB.OptLabel(0).Height
    For N = Which + 1 To 19
        FAB.OptLabel(N).Top = FAB.OptLabel(N).Top - H
    Next N
    FAB.Height = FAB.Height - H
       
    ' We must also reposition the command button so that
    ' it isn't hidden by the shrunken FAB form...
    FAB.CommandOK.Top = FAB.Height - (FAB.CommandOK.Height + H + 200)
    FAB.CommandOK.Left = FAB.Width - (FAB.CommandOK.Width + 200)
  
End Sub


'
' FUNCTION: RegQueryStringValue
'
' Retrieves the string data for a named
' (strValueName = name) or unnamed (strValueName = "")
' value within a registry key.  If the named value
' exists, but its data is not a string, this function
' fails.
'
' Returns: True on success, else False.
'
' On success, strData is set to the string data value
'
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean

    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    RegQueryStringValue = False
    On Error GoTo 0
    ' Get length/data type
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = True
                strData = StripTerminator(strBuf)
            End If
        End If
    End If
    
End Function


'
'-----------------------------------------------------------
' FUNCTION: RegCloseKey
'
' Closes an open registry key.
'
' Returns: True on success, else False.
'-----------------------------------------------------------
'
Function RegCloseKey(ByVal hKey As Long) As Boolean

    Dim lResult As Long
    On Error GoTo 0
    lResult = OSRegCloseKey(hKey)
    RegCloseKey = (lResult = ERROR_SUCCESS)
    
End Function


'
'-----------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'-----------------------------------------------------------
'
Function StripTerminator(ByVal strString As String) As String
    
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
    
End Function



