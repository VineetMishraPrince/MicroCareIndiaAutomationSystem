VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Attendance"
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   1575
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Present"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   6480
      Width           =   975
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Leave"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   6000
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Absent"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2895
      Left            =   720
      TabIndex        =   10
      Top             =   4920
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Regular Attandance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Overtime Attendance"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   3200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1085
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
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Night"
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Day"
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   8640
      ScaleHeight     =   3825
      ScaleWidth      =   2865
      TabIndex        =   4
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2000
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   8880
      TabIndex        =   5
      Top             =   480
      Width           =   2760
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ST, STx, Set_ST_STx
Dim ET, ETx, Set_ET_ETx
Dim xST, xET
Private Sub Combo_AMPM1_Change()
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Combo_AMPM1_Validate(Cancel As Boolean)
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Combo_AMPM2_Change()
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Combo_AMPM2_Validate(Cancel As Boolean)
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Combo_Time1_Change()
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Combo_Time1_Validate(Cancel As Boolean)
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Combo_Time2_Change()
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Combo_Time2_Validate(Cancel As Boolean)
Form6.Text_TotalOvertimeHours.Text = ""
End Sub

Private Sub Command_Close_Click()
Unload Me
End Sub

Private Sub Command_OvertimeAttendance_Click()
If Trim(Form6.Text_TotalOvertimeHours.Text) <> "" Then
Form6.Adodc4.Refresh
Form6.Adodc4.Recordset.AddNew
Form6.Adodc4.Recordset.Fields(0).Value = Trim(Form6.DataCombo_StaffID.Text)
Form6.Adodc4.Recordset.Fields(1).Value = Trim(Form6.Combo_Time1.Text) & " " & Trim(Form6.Combo_AMPM1.Text) & " " & Trim(Form6.Combo_Time2.Text) & " " & Trim(Form6.Combo_AMPM2.Text)
If Val(Form6.Option_Shift_Day.Value) = 1 Then
    Form6.Adodc4.Recordset.Fields(2).Value = "Day"
Else
    Form6.Adodc4.Recordset.Fields(2).Value = "Night"
End If
Form6.Adodc4.Recordset.Fields(3).Value = Trim(Form6.Text_TotalOvertimeHours.Text)
Form6.Adodc4.Recordset.Fields(4).Value = CDate(Date)
Form6.Adodc4.Recordset.Fields(5).Value = Format(Now, "MMM")
Form6.Adodc4.Recordset.UpdateBatch
Form6.Adodc4.Refresh
MsgBox "The Overtime attendance is made successfully.", vbInformation, "Message"
Else
MsgBox "Incomplete Entry...!", vbInformation, "Message"
End If
End Sub

Private Sub Command_RegularAttendance_Click()
Form6.Adodc3.Refresh
z = "Employee_ID ='" & Trim(Form6.DataCombo_StaffID.Text) & "' And attendance_date = '" & Format(Trim(CDate(Date)), "dd-MMM-yyyy") & "'"
Form6.Adodc3.Recordset.Filter = z
If Form6.Adodc3.Recordset.EOF = True Then
Form6.Adodc3.Recordset.AddNew
Form6.Adodc3.Recordset.Fields(0).Value = Trim(Form6.DataCombo_StaffID.Text)
Form6.Adodc3.Recordset.Fields(1).Value = Trim(Form6.Text_GeneralDutyTiming.Text)
If Val(Form6.Option_Shift_Day.Value) = 1 Then
    Form6.Adodc3.Recordset.Fields(2).Value = "Day"
Else
    Form6.Adodc3.Recordset.Fields(2).Value = "Night"
End If
If Val(Form6.Option_Leave.Value) = 1 Or Val(Form6.Option_Present.Value) = 1 Then
    Form6.Adodc3.Recordset.Fields(3).Value = "P"
Else
    Form6.Adodc3.Recordset.Fields(3).Value = "A"
End If
Form6.Adodc3.Recordset.Fields(4).Value = CDate(Date)
Form6.Adodc3.Recordset.Fields(5).Value = Format(Now, "MMM")
Form6.Adodc3.Recordset.UpdateBatch
Form6.Adodc3.Refresh
MsgBox "The Regular attendance is made successfully.", vbInformation, "Message"
Else
rep = MsgBox("The regular attendance is already made. Do you want to do it again for this day?", vbYesNo, "Message")
If rep = vbYes Then
Form6.Adodc3.Recordset.Fields(0).Value = Trim(Form6.DataCombo_StaffID.Text)
Form6.Adodc3.Recordset.Fields(1).Value = Trim(Form6.Text_GeneralDutyTiming.Text)
If Val(Form6.Option_Shift_Day.Value) = 1 Then
    Form6.Adodc3.Recordset.Fields(2).Value = "Day"
Else
    Form6.Adodc3.Recordset.Fields(2).Value = "Night"
End If
If Val(Form6.Option_Leave.Value) = 1 Or Val(Form6.Option_Present.Value) = 1 Then
    Form6.Adodc3.Recordset.Fields(3).Value = "P"
Else
    Form6.Adodc3.Recordset.Fields(3).Value = "A"
End If
Form6.Adodc3.Recordset.Fields(4).Value = CDate(Date)
Form6.Adodc3.Recordset.Fields(5).Value = Format(Now, "MMM")
Form6.Adodc3.Recordset.UpdateBatch
Form6.Adodc3.Refresh
MsgBox "The Regular attendance is made successfully again.", vbInformation, "Message"
End If
End If
End Sub

Private Sub CommandButton1_Click()
If Trim(Form6.Combo_AMPM1.Text) <> "" And Trim(Form6.Combo_AMPM2.Text) <> "" And Trim(Form6.Combo_Time1.Text) <> "" And Trim(Form6.Combo_Time2.Text) <> "" Then
    Call Check_ProperOvertime
Else
    MsgBox "Incomplete Entry...!", vbCritical, "Message"
End If
End Sub

Private Sub DataCombo_StaffID_KeyUp(KeyCode As Integer, Shift As Integer)
xxx = "Employee_ID = '" & Trim(Form6.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form6.DataCombo_StaffID.Text)
Form6.Adodc1.Refresh
Form6.Adodc1.Recordset.Find xxx
If Form6.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
Form6.Adodc2.Refresh
Form6.Adodc2.Recordset.Filter = xxx
Call Grep_GeneralDutyTime
Call Set_GeneralDutyTime
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
Call IsItWorkingDay
Form6.DataCombo_StaffID.SetFocus
End Sub

Private Sub DataCombo_StaffID_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
xxx = "Employee_ID = '" & Trim(Form6.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form6.DataCombo_StaffID.Text)
Form6.Adodc1.Refresh
Form6.Adodc1.Recordset.Find xxx
If Form6.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
Form6.Adodc2.Refresh
Form6.Adodc2.Recordset.Filter = xxx
Call Grep_GeneralDutyTime
Call Set_GeneralDutyTime
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
Call IsItWorkingDay
End Sub

Private Sub DataCombo_StaffID_Validate(Cancel As Boolean)
Form6.DataCombo_StaffID.SetFocus
End Sub

Private Sub DataGrid1_Click()
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
End Sub

Private Sub Form_Load()
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
Form6.Option_Shift_Day.Value = True
Form6.Option_Present.Value = True
If Form6.TabStrip1.SelectedItem.Index = 1 Then
    Form6.Frame1.Visible = True
    Form6.Frame2.Visible = False
Else
    Form6.Frame1.Visible = False
    Form6.Frame2.Visible = True
End If
Dim x As Integer
For x = 1 To 9
Form6.Combo_Time1.AddItem "0" & x
Form6.Combo_Time2.AddItem "0" & x
Next x
For x = 10 To 12
Form6.Combo_Time1.AddItem x
Form6.Combo_Time2.AddItem x
Next x
Form6.Combo_AMPM1.AddItem "AM"
Form6.Combo_AMPM1.AddItem "PM"
Form6.Combo_AMPM2.AddItem "AM"
Form6.Combo_AMPM2.AddItem "PM"
Call Grep_GeneralDutyTime
Call Set_GeneralDutyTime
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
xxx = "Employee_ID = '" & Trim(Form6.DataCombo_StaffID.Text) & "'"
yyy = Trim(Form6.DataCombo_StaffID.Text)
Form6.Adodc1.Refresh
Form6.Adodc1.Recordset.Find xxx
If Form6.Adodc1.Recordset.EOF = True Then
    Exit Sub
End If
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
Form6.Adodc2.Refresh
Form6.Adodc2.Recordset.Filter = xxx
Call Grep_GeneralDutyTime
Call Set_GeneralDutyTime
Form6.Text_GeneralDutyTiming.Text = Trim(Form6.Text_GeneralDutyTiming.Text)
Call IsItWorkingDay
End Sub

Private Sub TabStrip1_Click()
If Form6.TabStrip1.SelectedItem.Index = 1 Then
    Form6.Frame1.Visible = True
    Form6.Frame2.Visible = False
Else
    Form6.Frame1.Visible = False
    Form6.Frame2.Visible = True
End If
End Sub

Private Sub Text_TotalOvertimeHours_Change()
If Val(Form6.Text_TotalOvertimeHours.Text) > 0 Then
    Form6.Command_OvertimeAttendance.Visible = True
Else
    Form6.Command_OvertimeAttendance.Visible = False
End If
End Sub

Private Sub Text_TotalOvertimeHours_Validate(Cancel As Boolean)
If Val(Form6.Text_TotalOvertimeHours.Text) > 0 Then
    Form6.Command_OvertimeAttendance.Visible = True
Else
    Form6.Command_OvertimeAttendance.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
Form6.Label_DateOfPurchase.Caption = Format(Date, "dddd, dd/MM/yyyy")
Form6.Label_TimeOfPurchase.Caption = Format(Now, "h:mm:ss AM/PM")
End Sub

Public Sub Grep_GeneralDutyTime()
s = Trim(Form6.Text_GeneralDutyTiming.Text)
ST = Mid(s, 1, 2)
STx = Mid(s, 4, 2)
ET = Mid(s, 9, 2)
ETx = Mid(s, 12, 2)
End Sub

Public Sub Set_GeneralDutyTime()
If UCase(STx) = "AM" Then
    xST = ST
Else
    xST = Val(ST) + 12
End If

If UCase(ETx) = "AM" Then
    xET = ET
Else
    xET = Val(ET) + 12
End If
End Sub

Public Sub Check_ProperOvertime()
S1 = Trim(Form6.Combo_Time1.Text)
SM = Trim(Form6.Combo_AMPM1.Text)
E1 = Trim(Form6.Combo_Time2.Text)
EM = Trim(Form6.Combo_AMPM2.Text)
'MsgBox S1 & " " & SM & " - " & E1 & " " & EM
If UCase(SM) = "AM" Then
    XS = S1
Else
    XS = Val(S1) + 12
End If
If UCase(EM) = "AM" Then
    XE = E1
Else
    XE = Val(E1) + 12
End If
                                                    
If Val(Form6.Text_TotalOvertimeHours.Text) > 0 Then
    Form6.Command_OvertimeAttendance.Visible = True
Else
    Form6.Command_OvertimeAttendance.Visible = False
End If

Call IsItWorkingDay
    'MsgBox "1: " & Form6.Command_RegularAttendance.Locked

If Form6.Command_RegularAttendance.Locked = True Then
    'MsgBox "X"
    'MsgBox "X: " & Form6.Command_RegularAttendance.Locked
    GoTo NoGeneralDutyDayCase_OnlyOvertime
End If                                                   'MsgBox XS & " - " & XE
                                                    'MsgBox Val(xST) & " > " & Val(XS) & " And " & Val(xST) & " >= " & Val(XE) & " AND " & Val(xET) & " <= " & Val(XS) & " And " & Val(xET) & " < " & Val(XE)
If Val(xST) > Val(XS) And Val(xST) >= Val(XE) And Val(xET) <= Val(XS) And Val(xET) < Val(XE) Then
    GoTo TotallyMAD
Else
    GoTo Kanke
End If
                                                    'MsgBox Val(XS) & " > " & Val(xST) & " And Not " & Val(XE) & " < " & Val(xET)
If Val(XS) > Val(xST) And Val(XE) < Val(xET) Then
                                                    'MsgBox Val(XS) & " >= " & Val(xET) & " And " & Val(XS) & " <= " & Val(xST) & " And " & Val(XE) & " >= " & Val(xET) & " And " & Val(XE) & " <= " & Val(xST)
    If Val(XS) >= Val(xET) And Val(XS) <= Val(xST) And Val(XE) >= Val(xET) And Val(XE) <= Val(xST) Then
        MsgBox "The Overtime is not allocated in the General Duty time.", vbInformation, "Message"
    Else
TotallyMAD:
NoGeneralDutyDayCase_OnlyOvertime:
        If SM = "PM" And EM = "PM" Then
            Form6.Text_TotalOvertimeHours.Text = Val(XE) - Val(XS)
        End If
        If SM = "PM" And EM = "AM" Then
            Form6.Text_TotalOvertimeHours.Text = (24 - Val(XS)) + Val(XE)
        End If
        If SM = "AM" And EM = "AM" Then
            Form6.Text_TotalOvertimeHours.Text = Val(XE) - Val(XS)
        End If
        If SM = "AM" And EM = "PM" Then
            Form6.Text_TotalOvertimeHours.Text = Val(XE) - Val(XS)
        End If
    End If
Else
Kanke:
        MsgBox "The Overtime can't lies between the General Duty time.", vbInformation, "Message"
End If
End Sub

Private Sub IsItWorkingDay()
Dim wd_SUN, wd_MON, wd_TUE, wd_WED, wd_THU, wd_FRI, wd_SAT As Integer
CurrentWeekDay = Trim(UCase(Format(Now, "ddd")))
Form6.Command_RegularAttendance.Locked = True
'MsgBox "1: " & Form6.Command_RegularAttendance.Visible
'MsgBox CurrentWeekDay
wd_SUN = Trim(Form6.DataGrid1.Columns(2).Value)
wd_MON = Trim(Form6.DataGrid1.Columns(3).Value)
wd_TUE = Trim(Form6.DataGrid1.Columns(4).Value)
wd_WED = Trim(Form6.DataGrid1.Columns(5).Value)
wd_THU = Trim(Form6.DataGrid1.Columns(6).Value)
wd_FRI = Trim(Form6.DataGrid1.Columns(7).Value)
wd_SAT = Trim(Form6.DataGrid1.Columns(8).Value)
'MsgBox wd_SUN & wd_MON & wd_TUE & wd_WED & wd_THU & wd_FRI & wd_SAT
If Trim(CurrentWeekDay) = "SUN" Then
    If Val(wd_SUN) = -1 Then
        Form6.Command_RegularAttendance.Locked = False
        Exit Sub
    End If
End If
If Trim(CurrentWeekDay) = "MON" Then
    If Val(wd_MON) = -1 Then
        Form6.Command_RegularAttendance.Locked = False
        Exit Sub
    End If
End If
If Trim(CurrentWeekDay) = "TUE" Then
    If Val(wd_TUE) = -1 Then
        Form6.Command_RegularAttendance.Locked = False
        Exit Sub
    End If
End If
If Trim(CurrentWeekDay) = "WED" Then
    If Val(wd_WED) = -1 Then
        Form6.Command_RegularAttendance.Locked = False
        Exit Sub
    End If
End If
If Trim(CurrentWeekDay) = "THU" Then
    If Val(wd_THU) = -1 Then
        Form6.Command_RegularAttendance.Locked = False
        Exit Sub
    End If
End If
If Trim(CurrentWeekDay) = "FRI" Then
    If Val(wd_FRI) = -1 Then
        Form6.Command_RegularAttendance.Locked = False
        Exit Sub
    End If
End If
If Trim(CurrentWeekDay) = "SAT" Then
    If Val(wd_SAT) = -1 Then
        Form6.Command_RegularAttendance.Locked = False
        Exit Sub
    End If
End If
End Sub
 

