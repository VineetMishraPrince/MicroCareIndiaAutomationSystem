VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "Employee Salary Details"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   8640
      ScaleHeight     =   3825
      ScaleWidth      =   2745
      TabIndex        =   12
      Top             =   1560
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   8280
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   5520
   End
   Begin VB.CommandButton Command_Close 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command_Update 
      BackColor       =   &H00FF8080&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5700
      Width           =   1215
   End
   Begin VB.TextBox Text_OT_Night 
      Height          =   390
      Left            =   5880
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   6675
      Width           =   1095
   End
   Begin VB.TextBox Text_OT_day 
      Height          =   390
      Left            =   2760
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   6675
      Width           =   1095
   End
   Begin VB.TextBox Text_BS_Night 
      Height          =   390
      Left            =   5880
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   4650
      Width           =   1095
   End
   Begin VB.TextBox Text_BS_Day 
      Height          =   390
      Left            =   2760
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   4650
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3120
      Width           =   5655
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2700
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2700
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   400
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2100
      Width           =   5655
   End
   Begin VB.ComboBox DataCombo_StaffID 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label_Date 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9720
      TabIndex        =   14
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label_Time 
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   9720
      TabIndex        =   13
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Close_Click()
Unload Me
End Sub

Private Sub Command_Update_Click()
If IsNumeric(Form5.Text_BS_Day.Text) = True And IsNumeric(Form5.Text_BS_Night.Text) = True And IsNumeric(Form5.Text_OT_day.Text) = True And IsNumeric(Form5.Text_OT_Night.Text) = True Then
    v1 = Form5.Text_BS_Day.Text
    v2 = Form5.Text_BS_Night.Text
    v3 = Form5.Text_OT_day.Text
    v4 = Form5.Text_OT_Night.Text
    xxx = "Employee_ID = '" & Trim(Form5.DataCombo_StaffID.Text) & "'"
    Form5.Adodc1.Refresh
    Form5.Adodc1.Recordset.Find xxx
    Form5.Adodc1.Recordset.Fields(18).Value = v1
    Form5.Adodc1.Recordset.Fields(19).Value = v2
    Form5.Adodc1.Recordset.Fields(20).Value = v3
    Form5.Adodc1.Recordset.Fields(21).Value = v4
    Form5.Adodc1.Recordset.UpdateBatch
    Form5.Adodc1.Refresh
    Form5.Adodc1.Recordset.Find xxx
    MsgBox "The Salary Detail is successfully updated for this Employee.", vbInformation, "Message"
Else
    MsgBox "Invaild Entry...! Check each values.", vbCritical, "Message"
End If
End Sub

Private Sub DataCombo_StaffID_KeyUp(KeyCode As Integer, Shift As Integer)
xxx = "Employee_ID = '" & Trim(Form5.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form5.DataCombo_StaffID.Text)
Form5.Adodc1.Refresh
Form5.Adodc1.Recordset.Find xxx
If Form5.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
End Sub

Private Sub DataCombo_StaffID_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
xxx = "Employee_ID = '" & Trim(Form5.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form5.DataCombo_StaffID.Text)
Form5.Adodc1.Refresh
Form5.Adodc1.Recordset.Find xxx
If Form5.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Timer1_Timer()
Form5.Label_DateOfPurchase.Caption = Format(Date, "dd/MM/yyyy")
Form5.Label_TimeOfPurchase.Caption = Format(Now, "h:mm:ss AM/PM")
End Sub
 

