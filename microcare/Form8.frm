VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form2 
   Caption         =   "Service Wizard Window - 1/2 [Customer]"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   960
      TabIndex        =   16
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   7800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   7250
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT >"
      Height          =   375
      Left            =   8400
      TabIndex        =   14
      Top             =   7250
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "Form8.frx":17C79
      Top             =   6960
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Text            =   "Text7"
      Top             =   6500
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   4320
      Width           =   4815
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   3120
      TabIndex        =   7
      Top             =   3400
      Width           =   4815
   End
   Begin VB.TextBox Text_CustomerName 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3000
      Width           =   4815
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   3120
      TabIndex        =   5
      Top             =   1850
      Width           =   2655
   End
   Begin VB.TextBox Text_CustomerID 
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1500
      Width           =   2655
   End
   Begin VB.ComboBox Text_CustomerType 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodcy 
      Height          =   330
      Left            =   2880
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
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
      Height          =   305
      Left            =   9480
      TabIndex        =   1
      Top             =   840
      Width           =   1575
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
      Height          =   280
      Left            =   9480
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Service_ID As Integer

Private Sub Combo1_Change()
Form2.DataCombo_JobStatus.Text = Combo1.Text
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Form2.DataCombo_JobStatus.Text = Combo1.Text
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
Form2.DataCombo_JobStatus.Text = Combo1.Text
End Sub

Private Sub Combo2_Change()
Form2.DataCombo_Priority.Text = Combo2.Text
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
Form2.DataCombo_Priority.Text = Combo2.Text
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
Form2.DataCombo_Priority.Text = Combo2.Text
End Sub

Private Sub Combo3_Change()
Form2.DataCombo_ServiceType.Text = Combo3.Text
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
Form2.DataCombo_ServiceType.Text = Combo3.Text
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
Form2.DataCombo_ServiceType.Text = Combo3.Text
End Sub

Private Sub Command_Back_Click()
    Form2.Hide
    Form8.Show
End Sub

Private Sub CommandButton_Cancel_Click()
Unload Form8
Unload Form2
End Sub

Private Sub Command_Cancel_Click()
Unload Form2
On Error GoTo IfNoExistence
Unload Form8
IfNoExistence:
Exit Sub
End Sub

Private Sub Command_New_Click()
Service_ID = 0
Form2.Adodcy.Refresh
Service_ID = Val(Form2.Adodcy.Recordset.RecordCount) + 1
Form2.Command_Update.Enabled = True
Form2.Command_New.BackStyle = fmBackStyleTransparent
Form2.Command_Back.Enabled = False
Form2.Command_New.Enabled = False
Form2.Adodcy.Refresh
Form2.Adodcy.Recordset.AddNew
Form2.DataCombo1.Text = Trim(Val(Service_ID))
MsgBox Form2.DataCombo1.Text
Form2.DataCombo1.Enabled = False
MsgBox Form2.DataCombo1.Text
End Sub

Private Sub Command_Update_Click()
If Form2.Command_New.Enabled = False Then
    service_serial_no = Trim(Service_ID)
    system_date = Trim(Date)
    system_time = Trim(Time)
    Customer_ID = Trim(Customer_ID)
    Customer_Type = Trim(Customer_Type)
    c_name = Trim(Customer_Name)
    c_address = Trim(Address)
    c_phone = Trim(Phone_No)
    c_mobile = Trim(Mobile)
    c_fax = Trim(Fax_No)
    c_e_mail_id = Trim(Email)
    c_Description = Trim(Description)
    service_type = Trim(Form2.DataCombo_ServiceType.Text)
    Service_Requested = Trim(Form2.Text_SRequested.Text)
    Priority = Trim(Form2.DataCombo_Priority.Text)
    job_status = Trim(Form2.DataCombo_JobStatus.Text)
    responce_detail = Trim(Form2.Text_RResponce.Text)
    date_allotted = Trim(Form2.DTPicker_Date.Text)
    time_allotted = Trim(Form2.DTPicker_Time.Text)
    Form2.Adodcy.Recordset.Fields(0).Value = service_serial_no
    Form2.Adodcy.Recordset.Fields(1).Value = system_date
    Form2.Adodcy.Recordset.Fields(2).Value = system_time
    Form2.Adodcy.Recordset.Fields(3).Value = Customer_ID
    Form2.Adodcy.Recordset.Fields(4).Value = Customer_Type
    Form2.Adodcy.Recordset.Fields(5).Value = c_name
    Form2.Adodcy.Recordset.Fields(6).Value = c_address
    Form2.Adodcy.Recordset.Fields(7).Value = c_phone
    Form2.Adodcy.Recordset.Fields(8).Value = c_mobile
    Form2.Adodcy.Recordset.Fields(9).Value = c_fax
    Form2.Adodcy.Recordset.Fields(10).Value = c_e_mail_id
    Form2.Adodcy.Recordset.Fields(11).Value = c_Description
    Form2.Adodcy.Recordset.Fields(12).Value = service_type
    Form2.Adodcy.Recordset.Fields(13).Value = Service_Requested
    Form2.Adodcy.Recordset.Fields(14).Value = Priority
    Form2.Adodcy.Recordset.Fields(15).Value = job_status
    Form2.Adodcy.Recordset.Fields(16).Value = responce_detail
    Form2.Adodcy.Recordset.Fields(17).Value = date_allotted
    Form2.Adodcy.Recordset.Fields(18).Value = time_allotted
    Form2.Adodcy.Recordset.UpdateBatch
    Form2.Adodcy.Refresh
    Form2.DataCombo1.Enabled = True
    Form2.Command_Back.Enabled = True
    Form2.Command_New.Enabled = True
    Form2.Command_New.BackStyle = fmBackStyleOpaque
    Form2.Command_Update.Enabled = True
    MsgBox "The new service is successfully added.", vbInformation, "Message"
Else
    service_serial_no = Trim(Form2.DataCombo1.Text)
    system_date = Trim(Date)
    system_time = Trim(Time)
    Customer_ID = Trim(Customer_ID)
    Customer_Type = Trim(Customer_Type)
    c_name = Trim(Customer_Name)
    c_address = Trim(Address)
    c_phone = Trim(Phone_No)
    c_mobile = Trim(Mobile)
    c_fax = Trim(Fax_No)
    c_e_mail_id = Trim(Email)
    c_Description = Trim(Description)
    service_type = Trim(Form2.DataCombo_ServiceType.Text)
    Service_Requested = Trim(Form2.Text_SRequested.Text)
    Priority = Trim(Form2.DataCombo_Priority.Text)
    job_status = Trim(Form2.DataCombo_JobStatus.Text)
    responce_detail = Trim(Form2.Text_RResponce.Text)
    date_allotted = Trim(Form2.DTPicker_Date.Text)
    time_allotted = Trim(Form2.DTPicker_Time.Text)
    If service_serial_no = "" Then
    Exit Sub
    End If
    Form2.Adodcy.Recordset.Fields(0).Value = service_serial_no
    Form2.Adodcy.Recordset.Fields(1).Value = system_date
    Form2.Adodcy.Recordset.Fields(2).Value = system_time
    Form2.Adodcy.Recordset.Fields(3).Value = Customer_ID
    Form2.Adodcy.Recordset.Fields(4).Value = Customer_Type
    Form2.Adodcy.Recordset.Fields(5).Value = c_name
    Form2.Adodcy.Recordset.Fields(6).Value = c_address
    Form2.Adodcy.Recordset.Fields(7).Value = c_phone
    Form2.Adodcy.Recordset.Fields(8).Value = c_mobile
    Form2.Adodcy.Recordset.Fields(9).Value = c_fax
    Form2.Adodcy.Recordset.Fields(10).Value = c_e_mail_id
    Form2.Adodcy.Recordset.Fields(11).Value = c_Description
    Form2.Adodcy.Recordset.Fields(12).Value = service_type
    Form2.Adodcy.Recordset.Fields(13).Value = Service_Requested
    Form2.Adodcy.Recordset.Fields(14).Value = Priority
    Form2.Adodcy.Recordset.Fields(15).Value = job_status
    Form2.Adodcy.Recordset.Fields(16).Value = responce_detail
    Form2.Adodcy.Recordset.Fields(17).Value = date_allotted
    Form2.Adodcy.Recordset.Fields(18).Value = time_allotted
    Form2.Adodcy.Recordset.UpdateBatch
    Form2.Adodcy.Refresh
    Form2.DataCombo1.Enabled = True
    Form2.Command_Back.Enabled = True
    Form2.Command_New.Enabled = True
    Form2.Command_Update.Enabled = True
    MsgBox "The existing service is successfully updated.", vbInformation, "Message"
End If
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
If Form2.Command_New.Enabled = True And Form2.DataCombo1.Text <> "" Then
    v = "service_serial_no = '" & Trim(Form2.DataCombo1.Text) & "' and Customer_ID = '" & Trim(Customer_ID) & "'"
    Form2.Adodcy.Refresh
    Form2.Adodcy.Recordset.Filter = v
    If Form2.Adodcy.Recordset.EOF = True Then
        Form2.DataCombo1.Text = ""
        Exit Sub
    End If
End If
End Sub

Private Sub DataCombo1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Form2.Command_New.Enabled = True And Form2.DataCombo1.Text <> "" Then
    v = "service_serial_no = '" & Trim(Form2.DataCombo1.Text) & "' and Customer_ID = '" & Trim(Customer_ID) & "'"
    Form2.Adodcy.Refresh
    Form2.Adodcy.Recordset.Filter = v
    If Form2.Adodcy.Recordset.EOF = True Then
        Form2.DataCombo1.Text = ""
        Exit Sub
    End If
End If
End Sub

Private Sub DTPicker_Time_Change()
If Form2.DTPicker_Time.Text <> Null Then
Form2.DTPicker2.Value = Form2.DTPicker_Time.Text
End If
End Sub

Private Sub DTPicker_Time_Validate(Cancel As Boolean)
If Form2.DTPicker_Time.Text <> Null Then
Form2.DTPicker2.Value = Form2.DTPicker_Time.Text
End If
End Sub

Private Sub DTPicker1_Change()
Form2.DTPicker_Date.Text = CDate(DTPicker1.Value)
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
Form2.DTPicker_Date.Text = CDate(DTPicker1.Value)
End Sub

Private Sub DTPicker1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Form2.DTPicker_Date.Text = CDate(DTPicker1.Value)
End Sub

Private Sub DTPicker2_Change()
Form2.DTPicker_Time.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
Form2.DTPicker_Time.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Form2.DTPicker_Time.Text = DTPicker2.Value
End Sub


Private Sub Form_Activate()
Form2.Text_CustomerID.Text = Trim(Customer_ID)
Form2.Text_CustomerType.Text = Trim(Customer_Type)
Form2.Text_CustomerName.Text = Trim(Customer_Name)
v = "customer_id = '" & Trim(Customer_ID) & "' And Job_Status <> 'Completed'"
Form2.Adodcy.Refresh
Form2.Adodcy.Recordset.Filter = v
If Form2.Adodcy.Recordset.EOF = True Then
    Call Clinic_All_Clear
    Form2.DataCombo1.Enabled = False
    Form2.DataCombo1.ReFill
    Form2.DataCombo1.Refresh
    Exit Sub
Else
    Form2.DataCombo1.Enabled = True
    Form2.DataCombo1.ReFill
    Form2.DataCombo1.Refresh
End If
End Sub

Private Sub Form_Load()
Form2.Text_CustomerID.Text = Trim(Customer_ID)
Form2.Text_CustomerType.Text = Trim(Customer_Type)
Form2.Text_CustomerName.Text = Trim(Customer_Name)
v = "customer_id = '" & Trim(Customer_ID) & "' And Job_Status <> 'Completed'"
Form2.Adodcy.Refresh
Form2.Adodcy.Recordset.Filter = v
If Form2.Adodcy.Recordset.EOF = True Then
    Call Clinic_All_Clear
    Form2.DataCombo1.Enabled = False
    Form2.DataCombo1.ReFill
    Form2.DataCombo1.Refresh
    Exit Sub
Else
    Form2.DataCombo1.Enabled = True
    Form2.DataCombo1.ReFill
    Form2.DataCombo1.Refresh
End If
End Sub

Private Sub Timer1_Timer()
Form2.Label_Date.Caption = Format(Date, "dd/MM/yyyy")
Form2.Label_Time.Caption = Format(Now, "h:mm:ss AM/PM")
End Sub

Public Sub Clinic_All_Clear()
    Form2.DataCombo1.Text = ""
    Form2.DataCombo_ServiceType.Text = ""
    Form2.Text_SRequested.Text = ""
    Form2.DataCombo_Priority.Text = ""
    Form2.DataCombo_JobStatus.Text = ""
    Form2.Text_RResponce.Text = ""
    Form2.DTPicker1.Value = Date
    Form2.DTPicker2.Value = Time
End Sub

