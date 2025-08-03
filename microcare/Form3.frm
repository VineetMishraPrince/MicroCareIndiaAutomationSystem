VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Payroll Windows for Staff"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3120
      Top             =   8160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc rsx 
      Height          =   375
      Left            =   960
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin VB.TextBox Text10 
      Height          =   400
      Left            =   5400
      TabIndex        =   13
      Text            =   "Text10"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   400
      Left            =   2280
      TabIndex        =   12
      Text            =   "Text9"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   400
      Left            =   5400
      TabIndex        =   11
      Text            =   "Text8"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   400
      Left            =   2280
      TabIndex        =   10
      Text            =   "Text7"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   7340
      Width           =   1575
   End
   Begin VB.CommandButton Command_Payment 
      Caption         =   "&Payment"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   7340
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   420
      Left            =   9840
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   6325
      Width           =   1720
   End
   Begin VB.TextBox Text_NPA 
      Height          =   400
      Left            =   9840
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   5350
      Width           =   1700
   End
   Begin VB.TextBox TextOA 
      Height          =   420
      Left            =   9840
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   4423
      Width           =   1725
   End
   Begin VB.TextBox Text_PA 
      Height          =   400
      Left            =   9840
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text_Balance 
      Height          =   400
      Left            =   9840
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2260
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   2760
      Left            =   3840
      TabIndex        =   2
      Top             =   2100
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4470
      Width           =   2415
   End
   Begin VB.ListBox DataList_StaffName 
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   1900
      Width           =   3015
   End
   Begin VB.Label Label_DateOfPurchase 
      BackColor       =   &H80000012&
      Caption         =   "Label2"
      Height          =   375
      Left            =   10080
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label_TimeOfPurchase 
      BackColor       =   &H80000012&
      Caption         =   "Label3"
      Height          =   375
      Left            =   10080
      TabIndex        =   16
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label Label_DOLP 
      BackColor       =   &H80000012&
      Caption         =   "Label2"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoOfGeneralDayShift, NoOfGeneralNightShift, TotalDayOvertimeHours, TotalNightOvertimeHours As Integer
Dim PaymentForGeneralDayShift, PaymentForGeneralNightShift, PaymentForTotalDayOvertimeHours, PaymentForTotalNightOvertimeHours As Double

Private Sub Command_Cancel_Click()
Unload Me
End Sub

Private Sub Command_Payment_Click()
If IsNumeric(Form3.Text_PaidAmount.Text) = True And IsNumeric(Form3.TextOA.Text) = True Then
ss = Trim(Form3.Text_StaffIDNo.Caption)
s = "Employee_ID = '" & ss & "'"
v1 = Trim(ss)
v2 = Val(NoOfGeneralDayShift)
v3 = Val(NoOfGeneralNightShift)
v4 = Val(TotalDayOvertimeHours)
v5 = Val(TotalNightOvertimeHours)
v6 = Val(PaymentForGeneralDayShift)
v7 = Val(PaymentForGeneralNightShift)
v8 = Val(PaymentForTotalDayOvertimeHours)
v9 = Val(PaymentForTotalNightOvertimeHours)
v10 = Trim(Form3.TextOA.Text)
v11 = Trim(Form3.Text_PA.Text)
v12 = Trim(Form3.Text_NPA.Text)
v13 = Trim(Form3.Text_PaidAmount.Text)
v14 = Val(Form3.Text_NPA.Text) - Val(Form3.Text_PaidAmount.Text)
v15 = CDate(Date)
'MsgBox v1 & v2 & v3 & v4 & v5 & v6 & v7 & v8 & v9 & v10 & v11 & v12 & v13 & v14 & v15
Form3.rsx.Refresh
Form3.rsx.Recordset.Find s
If Form3.rsx.Recordset.EOF = True Then
    Form3.rsx.Recordset.AddNew
    Form3.rsx.Recordset.Fields(0).Value = v1
    Form3.rsx.Recordset.Fields(1).Value = v2
    Form3.rsx.Recordset.Fields(2).Value = v3
    Form3.rsx.Recordset.Fields(3).Value = v4
    Form3.rsx.Recordset.Fields(4).Value = v5
    Form3.rsx.Recordset.Fields(5).Value = v6
    Form3.rsx.Recordset.Fields(6).Value = v7
    Form3.rsx.Recordset.Fields(7).Value = v8
    Form3.rsx.Recordset.Fields(8).Value = v9
    Form3.rsx.Recordset.Fields(9).Value = v10
    Form3.rsx.Recordset.Fields(10).Value = v11
    Form3.rsx.Recordset.Fields(11).Value = v12
    Form3.rsx.Recordset.Fields(12).Value = v13
    Form3.rsx.Recordset.Fields(13).Value = v14
    Form3.rsx.Recordset.Fields(14).Value = v15
    Form3.rsx.Recordset.UpdateBatch
    Form3.rsx.Refresh
    MsgBox "The Payment is successfully updated.", vbInformation, "Message"
Else
    Form3.rsx.Recordset.Fields(0).Value = v1
    Form3.rsx.Recordset.Fields(1).Value = v2
    Form3.rsx.Recordset.Fields(2).Value = v3
    Form3.rsx.Recordset.Fields(3).Value = v4
    Form3.rsx.Recordset.Fields(4).Value = v5
    Form3.rsx.Recordset.Fields(5).Value = v6
    Form3.rsx.Recordset.Fields(6).Value = v7
    Form3.rsx.Recordset.Fields(7).Value = v8
    Form3.rsx.Recordset.Fields(8).Value = v9
    Form3.rsx.Recordset.Fields(9).Value = v10
    Form3.rsx.Recordset.Fields(10).Value = v11
    Form3.rsx.Recordset.Fields(11).Value = v12
    Form3.rsx.Recordset.Fields(12).Value = v13
    Form3.rsx.Recordset.Fields(13).Value = v14
    Form3.rsx.Recordset.Fields(14).Value = v15
    Form3.rsx.Recordset.UpdateBatch
    Form3.rsx.Refresh
    MsgBox "The Payment is successfully updated.", vbInformation, "Message"
End If
Else
    MsgBox "Unable to update due to Invaild Entry...! Check the values.", vbExclamation, "Message"
End If
End Sub

Private Sub DataList_StaffName_KeyUp(KeyCode As Integer, Shift As Integer)
Call Clinic_All_Clear
xa = "Employee_Name like '" & Trim(Form3.DataList_StaffName.Text) & "'"
Form3.Adodc1.Refresh
Form3.Adodc1.Recordset.Find xa
Call Calc_Latest_Payment_OR_Joining_Date
If CDate(Form3.Label_DOLP.Caption) = CDate(Date) Then
    Form3.Command_Payment.Enabled = False
    If Form3.Text_Balance.Text > 0 Then
        Form3.Command_Payment.Enabled = True
        Form3.Text_NPA.Text = (Val(Form3.Text_PA.Text) + Val(Form3.TextOA.Text)) + Val(Form3.Text_Balance.Text)
        Call PaymentSummary
    End If
    Exit Sub
Else
    Form3.Command_Payment.Enabled = True
End If
Call Calc_Payroll
Call PaymentSummary
End Sub

Private Sub DataList_StaffName_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Clinic_All_Clear
xa = "Employee_Name like '" & Trim(Form3.DataList_StaffName.Text) & "'"
Form3.Adodc1.Refresh
Form3.Adodc1.Recordset.Find xa
Call Calc_Latest_Payment_OR_Joining_Date
If CDate(Form3.Label_DOLP.Caption) = CDate(Date) Then
    Form3.Command_Payment.Enabled = False
    If Form3.Text_Balance.Text > 0 Then
        Form3.Command_Payment.Enabled = True
        Form3.Text_NPA.Text = (Val(Form3.Text_PA.Text) + Val(Form3.TextOA.Text)) + Val(Form3.Text_Balance.Text)
        Call PaymentSummary
    End If
    Exit Sub
Else
    Form3.Command_Payment.Enabled = True
End If
Call Calc_Payroll
Call PaymentSummary
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Call Clinic_All_Clear
xa = "Employee_Name like '" & Trim(Form3.DataList_StaffName.Text) & "'"
Form3.Adodc1.Refresh
Form3.Adodc1.Recordset.Find xa
Call Calc_Latest_Payment_OR_Joining_Date
If CDate(Form3.Label_DOLP.Caption) = CDate(Date) Then
    Form3.Command_Payment.Enabled = False
    If Form3.Text_Balance.Text > 0 Then
        Form3.Command_Payment.Enabled = True
        Form3.Text_NPA.Text = (Val(Form3.Text_PA.Text) + Val(Form3.TextOA.Text)) + Val(Form3.Text_Balance.Text)
        Call PaymentSummary
    End If
    Exit Sub
Else
    Form3.Command_Payment.Enabled = True
End If
Call Calc_Payroll
Call PaymentSummary
End Sub

Private Sub Text_PA_Change()
Form3.Text_NPA.Text = (Val(Form3.Text_PA.Text) + Val(Form3.TextOA.Text)) + Val(Form3.Text_Balance.Text)
End Sub


Private Sub TextOA_Change()
Form3.Text_NPA.Text = (Val(Form3.Text_PA.Text) + Val(Form3.TextOA.Text)) + Val(Form3.Text_Balance.Text)
End Sub


Private Sub Timer1_Timer()
Form3.Label_DateOfPurchase.Caption = Format(Date, "dd/MM/yyyy")
Form3.Label_TimeOfPurchase.Caption = Format(Now, "h:mm:ss AM/PM")
End Sub

Public Sub Calc_Latest_Payment_OR_Joining_Date()
Form3.Text_Balance.Text = "0"
xa = "Employee_ID = '" & Trim(Form3.Text_StaffIDNo.Caption) & "'"
Form3.rsx.Refresh
Form3.rsx.Recordset.Find xa
If Form3.rsx.Recordset.EOF = True Then
    Form3.Label_DOLP.Caption = Format(Trim(Form3.Adodc1.Recordset.Fields(16).Value), "dd-MMM-yyyy")
    Form3.Text_Balance.Text = "0"
    Exit Sub
Else
    Form3.Text_Balance.Text = Val(Form3.rsx.Recordset.Fields(13).Value)
    Form3.Label_DOLP.Caption = Format(Trim(Form3.rsx.Recordset.Fields(14).Value), "dd-MMM-yyyy")
End If
If CDate(Form3.Label_DOLP.Caption) = CDate(Date) Then
    Form3.Command_Payment.Enabled = False
    If Val(Form3.Text_Balance.Text) > 0 Then
        Form3.Command_Payment.Enabled = True
        Form3.Text_NPA.Text = (Val(Form3.Text_PA.Text) + Val(Form3.TextOA.Text)) + Val(Form3.Text_Balance.Text)
    End If
    Exit Sub
Else
    Form3.Command_Payment.Enabled = True
End If
End Sub

Public Sub Calc_Payroll()
eID = Trim(Form3.Text_StaffIDNo.Caption)
Date1 = CDate(Form3.Label_DOLP.Caption)
Date2 = CDate(Date)
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
conn.Open "Provider=MSDAORA.1;User ID=ccsms;PASSWORD=zzz;Persist Security Info=False"
Set cmd.ActiveConnection = conn
cmd.CommandText = "SELECT Count(*) from Regular_ATTENDENCE where Employee_ID = '" & Trim(eID) & "' and Shift = 'Day' and attendance_status = 'P' and attendance_date >= '" & Format(CDate(Date1), "dd-MMM-yyyy") & "' And attendance_date < '" & Format(CDate(Date2), "dd-MMM-yyyy") & "'"
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockBatchOptimistic
rs.Requery
NoOfGeneralDayShift = rs.GetString()
conn.Close
Dim xconnx As New ADODB.Connection
Dim xcmdx As New ADODB.Command
Dim xrsx As New ADODB.Recordset
xconnx.Open "Provider=MSDAORA.1;User ID=ccsms;PASSWORD=zzz;Persist Security Info=False"
Set xcmdx.ActiveConnection = xconnx
xcmdx.CommandText = "SELECT Count(*) from Regular_ATTENDENCE where Employee_ID = '" & Trim(eID) & "' and Shift = 'Night' and attendance_status = 'P' and attendance_date >= '" & Format(CDate(Date1), "dd-MMM-yyyy") & "' And attendance_date < '" & Format(CDate(Date2), "dd-MMM-yyyy") & "'"
xrsx.CursorLocation = adUseClient
xrsx.Open xcmdx, , adOpenStatic, adLockBatchOptimistic
xrsx.Requery
NoOfGeneralNightShift = xrsx.GetString()
xconnx.Close
Dim conn1 As New ADODB.Connection
Dim cmd1 As New ADODB.Command
Dim rs1 As New ADODB.Recordset
conn1.Open "Provider=MSDAORA.1;User ID=ccsms;PASSWORD=zzz;Persist Security Info=False"
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "SELECT Sum(Overtime_hours) from Overtime_ATTENDENCE where Employee_ID = '" & Trim(eID) & "' and Shift = 'Day' and attendance_date >= '" & Format(CDate(Date1), "dd-MMM-yyyy") & "' And attendance_date < '" & Format(CDate(Date2), "dd-MMM-yyyy") & "'"
rs1.CursorLocation = adUseClient
rs1.Open cmd1, , adOpenStatic, adLockBatchOptimistic
rs1.Requery
TotalDayOvertimeHours = rs1.GetString()
If TotalDayOvertimeHours = 0 Then
    TotalDayOvertimeHours = 0
End If
conn1.Close
Dim conn2 As New ADODB.Connection
Dim cmd2 As New ADODB.Command
Dim rs2 As New ADODB.Recordset
conn2.Open "Provider=MSDAORA.1;User ID=ccsms;PASSWORD=zzz;Persist Security Info=False"
Set cmd2.ActiveConnection = conn2
cmd2.CommandText = "SELECT Sum(Overtime_hours) from Overtime_ATTENDENCE where Employee_ID = '" & Trim(eID) & "' and Shift = 'Night' and attendance_date >= '" & Format(CDate(Date1), "dd-MMM-yyyy") & "' And attendance_date < '" & Format(CDate(Date2), "dd-MMM-yyyy") & "'"
rs2.CursorLocation = adUseClient
rs2.Open cmd2, , adOpenStatic, adLockBatchOptimistic
rs2.Requery
TotalNightOvertimeHours = rs2.GetString()
If TotalNightOvertimeHours = 0 Then
    TotalNightOvertimeHours = 0
End If
conn2.Close
PaymentForGeneralDayShift = Val(Form3.Text_BS_Day.Text) * Val(NoOfGeneralDayShift)
PaymentForGeneralNightShift = Val(Form3.Text_BS_Night.Text) * Val(NoOfGeneralNightShift)
PaymentForTotalDayOvertimeHours = Val(Form3.Text_OT_day.Text) * Val(TotalDayOvertimeHours)
PaymentForTotalNightOvertimeHours = Val(Form3.Text_OT_Night.Text) * Val(TotalNightOvertimeHours)
Form3.Text_PA.Text = ""
Form3.Text_PA.Text = Val(PaymentForGeneralDayShift) + Val(PaymentForGeneralNightShift) + Val(PaymentForTotalDayOvertimeHours) + Val(PaymentForTotalNightOvertimeHours)
End Sub

Public Sub Clinic_All_Clear()
NoOfGeneralDayShift = 0
NoOfGeneralNightShift = 0
TotalDayOvertimeHours = 0
TotalNightOvertimeHours = 0
PaymentForGeneralDayShift = 0
PaymentForGeneralNightShift = 0
PaymentForTotalDayOvertimeHours = 0
PaymentForTotalNightOvertimeHours = 0
Form3.Text_Balance.Text = "0"
Form3.Text_NPA.Text = "0"
Form3.Text_PA.Text = "0"
Form3.TextOA.Text = "0"
Call Text_PA_Change
End Sub

Public Sub PaymentSummary()
Form3.List_PaymentSummary.Clear
Form3.List_PaymentSummary.Refresh
v2 = Val(NoOfGeneralDayShift)
v3 = Val(NoOfGeneralNightShift)
v4 = Val(TotalDayOvertimeHours)
v5 = Val(TotalNightOvertimeHours)
v6 = Val(PaymentForGeneralDayShift)
v7 = Val(PaymentForGeneralNightShift)
v8 = Val(PaymentForTotalDayOvertimeHours)
v9 = Val(PaymentForTotalNightOvertimeHours)
v10 = Trim(Form3.TextOA.Text)
v11 = Trim(Form3.Text_PA.Text)
v12 = Trim(Form3.Text_NPA.Text)
'v13 = Trim(Form3.Text_PaidAmount.Text)
v14 = Val(Form3.Text_Balance.Text)
v15 = CDate(Date)
Form3.List_PaymentSummary.AddItem "No Of General Day Shift: " & v2
Form3.List_PaymentSummary.AddItem "No Of General Night Shift: " & v3
Form3.List_PaymentSummary.AddItem "Total Day Overtime Hours: " & v4
Form3.List_PaymentSummary.AddItem "Total Night Overtime Hours: " & v5
Form3.List_PaymentSummary.AddItem "Payment For General Day Shift: " & v6
Form3.List_PaymentSummary.AddItem "Payment For General Night Shift: " & v7
Form3.List_PaymentSummary.AddItem "Payment For Total Night Overtime Hours: " & v8
Form3.List_PaymentSummary.AddItem "Payment For Total Night Overtime Hours: " & v9
Form3.List_PaymentSummary.AddItem "Other Allowance: " & v10
Form3.List_PaymentSummary.AddItem "Payroll Amount: " & v11
Form3.List_PaymentSummary.AddItem "Net Payment Amount: " & v12
'Form3.List_PaymentSummary.AddItem "Paid Amount: " & v13
Form3.List_PaymentSummary.AddItem "Previous Balance (If any): " & v14
Form3.List_PaymentSummary.AddItem "As Today, Date Of Payment: " & Format(v15, "dd/MM/yyyy")
End Sub
 

