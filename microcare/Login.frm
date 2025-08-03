VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{11CAE91C-7BBC-4790-9F95-0D4BE7F23ACE}#1.0#0"; "XPControl.ocx"
Begin VB.Form Login 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text_Password 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4370
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2700
      Width           =   1600
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   0
      Picture         =   "Login.frx":0000
      ScaleHeight     =   4185
      ScaleWidth      =   6750
      TabIndex        =   0
      Top             =   0
      Width           =   6750
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   240
         Top             =   720
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=MSDAORA.1;User ID=ccsms;Data Source=vOracle;Persist Security Info=False"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=ccsms;Data Source=vOracle;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "ccsms"
         Password        =   "zzz"
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo_Users 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "Login.frx":527E
         Left            =   4370
         List            =   "Login.frx":5294
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3020
         Width           =   2130
      End
      Begin VB.TextBox Text_UserName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4370
         TabIndex        =   5
         Top             =   2400
         Width           =   1600
      End
      Begin XPControl.XPButton Command_Close 
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   3360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Close"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin XPControl.XPButton Login_Button 
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Login  >>"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         Picture         =   "Login.frx":52D9
         ScaleHeight     =   345
         ScaleWidth      =   4185
         TabIndex        =   2
         Top             =   3720
         Width           =   4215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         Picture         =   "Login.frx":9FC7
         ScaleHeight     =   465
         ScaleWidth      =   4185
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Close_Click()
'    Login.Command_Close.BorderStyle = 1
    resp = MsgBox("Are you sure to close the Login Window?", vbQuestion + vbYesNo, "Message")
    If resp = vbYes Then
        Unload Me
    Else
        'Login.Command_Close.BorderStyle = 0
    End If
End Sub

Private Sub Login_Button_Click()
'Login.Login_Button.BorderStyle = 1
Login.Refresh
If Trim(Login.Text_UserName.Text) <> "" And Trim(Login.Text_Password.Text) <> "" And Trim(Login.Combo_Users.Text) <> "" Then
    Target = "UserProfile Like '" & Trim(Login.Combo_Users.Text) & "' And UserName Like '" & Trim(Login.Text_UserName.Text) & "' AND Password Like '" & Trim(Login.Text_Password.Text) & "'"
    Login.Adodc1.Refresh
    Login.Refresh
    Login.Adodc1.Recordset.Filter = Target
    If Login.Adodc1.Recordset.EOF = True Then
        MsgBox "Access Denied...!", vbInformation, "Message"
        Login.Text_Password.Text = ""
        Login.Text_UserName.Text = ""
    Else
        UserProfile = Trim(Login.Combo_Users.Text)
        Unload Login
        MDIForm1.Show
    End If
Else
    MsgBox "Incomplete Entry...!", vbExclamation, "Message"
End If
'Login.Login_Button.BorderStyle = 0
End Sub

