VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmabout 
   Caption         =   "About Microcare Call Centre Service Management System Window"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form14"
   MDIChild        =   -1  'True
   Picture         =   "frmabout.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   1080
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   5040
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Height          =   690
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdSysInfo 
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3625
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Dim WindowsInfo As String

Private Sub cmdOK_Click()
Unload frmabout
End Sub

Private Sub cmdSysInfo_Click()
Call StartSysInfo
End Sub

Private Sub Command1_Click()
Unload frmabout
frmSplash1.Show
End Sub
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    Call Shell(SysInfoPath, vbNormalFocus)
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                   ' Loop Counter
    Dim rc As Long                                   ' Return Code
    Dim hKey As Long              ' Handle To An Open Registry Key
    Dim hDepth As Long
    Dim KeyValType As Long          ' Data Type Of A Registry Key
    Dim tmpVal As String 'Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long        ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
 ' Open Registry Key
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
' Handle Error...
    tmpVal = String$(1024, 0)
' Allocate Variable Space
    KeyValSize = 1024                       ' Mark Variable Size
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)
    ' Get/Create Key Value
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)
' Null Found, Extract From String
    Else
' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)
 ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType
 ' Search Data Types...
    Case REG_SZ
' String Registry Key Data Type
        KeyVal = tmpVal
' Copy String Value
    Case REG_DWORD
' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1
' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)
' Convert Double Word To String
    End Select
    
    GetKeyValue = True
' Return Success
    rc = RegCloseKey(hKey)
' Close Registry Key
    Exit Function
' Exit
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""
' Set Return Val To Empty String
    GetKeyValue = False                       ' Return Failure
    rc = RegCloseKey(hKey)                  ' Close Registry Key
End Function

Private Sub Timer1_Timer()
'Static i As Integer
'If i = 0 Then
'frmabout.Image_ChangeComputerLogo.Picture = LoadPicture("C:\APMS\Multimedia\Graphics\Pictures\computerwww.bmp")
'i = 1
'ElseIf i = 1 Then
'frmabout.Image_ChangeComputerLogo.Picture = LoadPicture("C:\APMS\Multimedia\Graphics\Pictures\computerB.bmp")
'i = 2
'ElseIf i = 2 Then
'frmabout.Image_ChangeComputerLogo.Picture = LoadPicture("")
'i = 0
'End If
End Sub

Private Sub Form_Load()
Call GetWindowsInfo
'Call osinfo
'Label_osinfo.Caption = Trim(WindowsInfo) '& " " & Trim(osinfo_Str)
'frmabout.OLE1.Action = 7
End Sub

Public Sub GetWindowsInfo()
    Dim MsgEnd As String
    Select Case SysInfo1.OSPlatform
      Case 0
         MsgEnd = "Running on Microsoft Windows compatible system"
      Case 1
         MsgEnd = "Running on Microsoft Windows 9x [Version " & CStr(SysInfo1.OSVersion) & "." & CStr(SysInfo1.OSBuild) & "]"
      Case 2
         MsgEnd = "Running on Microsoft Windows NT [Version " & CStr(SysInfo1.OSVersion) & "." & CStr(SysInfo1.OSBuild) & "]"
    End Select
    WindowsInfo = MsgEnd
End Sub
 


