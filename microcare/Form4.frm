VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Employee Duty Allocation"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3600
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
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
      UserName        =   "vOracle"
      Password        =   "zzz"
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1680
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
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
      UserName        =   "vOracle"
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
      Left            =   600
      Top             =   8160
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      TabIndex        =   17
      Top             =   6890
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   650
      Left            =   1320
      TabIndex        =   16
      Top             =   6360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1138
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   8640
      ScaleHeight     =   3705
      ScaleWidth      =   2625
      TabIndex        =   15
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update &Duty Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   6120
      TabIndex        =   14
      Top             =   7110
      Width           =   1935
   End
   Begin VB.CommandButton Command_Update_DutyTiming 
      Caption         =   "&Update Duty Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   6120
      TabIndex        =   13
      Top             =   5010
      Width           =   1935
   End
   Begin VB.TextBox Text_GeneralDutyTiming 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   425
      Left            =   1920
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   5000
      Width           =   3870
   End
   Begin VB.ComboBox Combo_AMPM2 
      Height          =   315
      Left            =   4920
      TabIndex        =   11
      Text            =   "Combo5"
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox Combo_Time2 
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Text            =   "Combo4"
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox Combo_AMPM1 
      Height          =   315
      Left            =   2760
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox Combo_Time1 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   400
      Left            =   2760
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3480
      Width           =   5655
   End
   Begin VB.TextBox Text3 
      Height          =   400
      Left            =   6600
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   400
      Left            =   2760
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   400
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2400
      Width           =   5655
   End
   Begin VB.ComboBox DataCombo_StaffID 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label_DateOfPurchase 
      BackColor       =   &H00000000&
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
      Left            =   9735
      TabIndex        =   19
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label_TimeOfPurchase 
      BackColor       =   &H00000000&
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
      TabIndex        =   18
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   4090
      TabIndex        =   9
      Top             =   4580
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   1930
      TabIndex        =   6
      Top             =   4580
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim wdc As Integer
srj = Trim(Form4.DataCombo_StaffID.Text)
If Form4.Adodc2.Recordset.RecordCount > 0 Then
    Form4.Adodc2.Recordset.MoveNext
    If Form4.Adodc2.Recordset.EOF = True Then
        GoTo srj_bhy
    End If
End If
srj_bhy:
Form4.Adodc2.Refresh
Form4.Adodc2.Refresh
s = "Employee_ID ='" & Trim(srj) & "'"
Form4.Adodc2.Recordset.Find s
wdc = 0
If Val(Form4.Adodc2.Recordset.Fields(2).OriginalValue) = -1 Then
    wdc = wdc + 1
End If
If Val(Form4.Adodc2.Recordset.Fields(3).OriginalValue) = -1 Then
    wdc = wdc + 1
End If
If Val(Form4.Adodc2.Recordset.Fields(4).OriginalValue) = -1 Then
    wdc = wdc + 1
End If
If Val(Form4.Adodc2.Recordset.Fields(5).OriginalValue) = -1 Then
    wdc = wdc + 1
End If
If Val(Form4.Adodc2.Recordset.Fields(6).OriginalValue) = -1 Then
    wdc = wdc + 1
End If
If Val(Form4.Adodc2.Recordset.Fields(7).OriginalValue) = -1 Then
    wdc = wdc + 1
End If
If Val(Form4.Adodc2.Recordset.Fields(8).OriginalValue) = -1 Then
    wdc = wdc + 1
End If
Form4.Adodc2.Recordset.Fields(1).Value = wdc
Form4.Adodc2.Recordset.UpdateBatch
Form4.Adodc2.Refresh
Form4.Adodc2.Recordset.Find s
Form4.DataGrid1.Refresh
xxx = "Employee_ID = '" & Trim(Form4.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form4.DataCombo_StaffID.Text)
Form4.Adodc1.Refresh
Form4.Adodc1.Recordset.Find xxx
If Form4.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
Form4.Text_GeneralDutyTiming.Text = Trim(Form4.Text_GeneralDutyTiming.Text)

Form4.Adodc2.Refresh
Form4.Adodc2.Recordset.Filter = xxx
If Form4.Adodc2.Recordset.EOF = True Then
    Form4.Adodc2.Recordset.AddNew
    Form4.DataGrid1.Columns("Employee_ID").Value = yyy
End If
Form4.DataCombo_StaffID.SetFocus
End Sub

Private Sub Command_Close_Click()
Unload Form4
End Sub

Private Sub Command_Update_DutyTiming_Click()
If Trim(Form4.Combo_Time1.Text) <> "" And Trim(Form4.Combo_AMPM1.Text) <> "" And Trim(Form4.Combo_Time2.Text) <> "" And Trim(Form4.Combo_AMPM2.Text) <> "" Then
    Form4.Text_GeneralDutyTiming.Text = Trim(Form4.Combo_Time1.Text) & " " & Trim(Form4.Combo_AMPM1.Text) & " - " & Trim(Form4.Combo_Time2.Text) & " " & Trim(Form4.Combo_AMPM2.Text)
    Form4.Combo_Time1.Refresh
    Form4.Combo_AMPM1.Refresh
    Form4.Combo_Time2.Refresh
    Form4.Combo_AMPM2.Refresh
    GeneralDutyTiming = Trim(Form4.Text_GeneralDutyTiming.Text)
    xxx = "Employee_ID = '" & Trim(Form4.DataCombo_StaffID.Text) & "'"
    Form4.Adodc1.Refresh
    Form4.Adodc1.Recordset.Find xxx
    Form4.Adodc1.Recordset.Fields(17).Value = GeneralDutyTiming
    Form4.Adodc1.Recordset.UpdateBatch
    Form4.Adodc1.Refresh
    Form4.Adodc1.Recordset.Find xxx
    Form4.Text_GeneralDutyTiming.Text = Trim(Form4.Text_GeneralDutyTiming.Text)
    MsgBox "The Duty Timing is successfully updated.", vbInformation, "Message"
Else
    MsgBox "Incomlete Duty Timing...! Set the time values.", vbInformation, "Message"
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DataCombo_StaffID_KeyUp(KeyCode As Integer, Shift As Integer)
xxx = "Employee_ID = '" & Trim(Form4.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form4.DataCombo_StaffID.Text)
Form4.Adodc1.Refresh
Form4.Adodc1.Recordset.Find xxx
If Form4.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
Form4.Text_GeneralDutyTiming.Text = Trim(Form4.Text_GeneralDutyTiming.Text)
Form4.Adodc2.Refresh
Form4.Adodc2.Recordset.Filter = xxx
If Form4.Adodc2.Recordset.EOF = True Then
    Form4.Adodc2.Recordset.AddNew
    Form4.DataGrid1.Columns("Employee_ID").Value = yyy
End If
End Sub

Private Sub DataCombo_StaffID_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
xxx = "Employee_ID = '" & Trim(Form4.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form4.DataCombo_StaffID.Text)
Form4.Adodc1.Refresh
Form4.Adodc1.Recordset.Find xxx
If Form4.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
Form4.Text_GeneralDutyTiming.Text = Trim(Form4.Text_GeneralDutyTiming.Text)
Form4.Adodc2.Refresh
Form4.Adodc2.Recordset.Filter = xxx
If Form4.Adodc2.Recordset.EOF = True Then
    Form4.Adodc2.Recordset.AddNew
    Form4.DataGrid1.Columns("Employee_ID").Value = yyy
End If
End Sub

Private Sub Form_Load()
Dim x As Integer
For x = 1 To 9
Form4.Combo_Time1.AddItem "0" & x
Form4.Combo_Time2.AddItem "0" & x
Next x
For x = 10 To 12
Form4.Combo_Time1.AddItem x
Form4.Combo_Time2.AddItem x
Next x
Form4.Combo_AMPM1.AddItem "AM"
Form4.Combo_AMPM1.AddItem "PM"
Form4.Combo_AMPM2.AddItem "AM"
Form4.Combo_AMPM2.AddItem "PM"
Form4.Text_GeneralDutyTiming.Text = Trim(Form4.Text_GeneralDutyTiming.Text)
xxx = "Employee_ID = '" & Trim(Form4.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form4.DataCombo_StaffID.Text)
Form4.Adodc1.Refresh
Form4.Adodc1.Recordset.Find xxx
If Form4.Adodc1.Recordset.EOF = True Then
    Exit Sub
Else
    Form4.Adodc2.Refresh
    Form4.Adodc2.Recordset.Filter = xxx
    If Form4.Adodc2.Recordset.EOF = True Then
        Form4.Adodc2.Recordset.AddNew
        Form4.DataGrid1.Columns("Employee_ID").Value = yyy
    End If
End If
Form4.Text_GeneralDutyTiming.Text = Trim(Form4.Text_GeneralDutyTiming.Text)
End Sub

Private Sub Timer1_Timer()
Form4.Label_DateOfPurchase.Caption = Format(Date, "dd/MM/yyyy")
Form4.Label_TimeOfPurchase.Caption = Format(Now, "h:mm:ss AM/PM")
End Sub
 

