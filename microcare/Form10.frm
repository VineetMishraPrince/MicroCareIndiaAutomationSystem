VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form10 
   Caption         =   "Product/Service Payment"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   360
      TabIndex        =   14
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7320
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5280
      Top             =   8160
      Width           =   1680
      _ExtentX        =   2963
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
      Left            =   3120
      Top             =   8160
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
   Begin MSAdodcLib.Adodc Adodcx 
      Height          =   375
      Left            =   1440
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   3600
   End
   Begin VB.CommandButton Command_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   10350
      TabIndex        =   11
      Top             =   7500
      Width           =   1335
   End
   Begin VB.CommandButton Command_Update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   6750
      TabIndex        =   10
      Top             =   7500
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   9
      Top             =   5115
      Width           =   11450
      _ExtentX        =   20188
      _ExtentY        =   4048
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
   Begin VB.TextBox Text_PAddress 
      Height          =   670
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   4080
      Width           =   4845
   End
   Begin VB.ListBox List3 
      Height          =   640
      Left            =   1800
      TabIndex        =   7
      Top             =   3465
      Width           =   4845
   End
   Begin VB.TextBox Text3 
      Height          =   380
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   3060
      Width           =   4850
   End
   Begin VB.ListBox DataCombo_ServiceSlNo 
      Height          =   840
      Left            =   4800
      TabIndex        =   5
      Top             =   2040
      Width           =   1840
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1670
      Width           =   1825
   End
   Begin VB.ListBox DataCombo_StaffID 
      Height          =   840
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   2670
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   2670
   End
   Begin VB.ComboBox DataCombo_CustomerType 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2670
   End
   Begin VB.Label Label_TimeOfPurchase 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   10200
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label_DateOfPurchase 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cid As String
Private Sub Command_Cancel_Click()
Unload Me
End Sub

Private Sub Command_Update_Click()
cid = Trim(Form10.DataCombo_StaffID.Text)
If Trim(cid) <> "" And Trim(Form10.DataCombo_ServiceSlNo.Text) <> "" Then
    'Form10.Adodc3x.Refresh
    'Form10.Adodc3x.Recordset.AddNew
    Form10.DataGrid1.Columns(0).Value = Trim(Form10.DataCombo_ServiceSlNo.Text)
    Form10.DataGrid1.Columns(1).Value = Trim(cid)
    Form10.DataGrid1.Columns(2).Value = Trim(Form10.DataCombo_CustomerType.Text)
End If
End Sub

Private Sub DataCombo_CustomerTypex_Change()
Form10.DataCombo_CustomerType.Text = Trim(Form10.DataCombo_CustomerTypex.Text)
End Sub

Private Sub DataCombo_CustomerTypex_Click()
Form10.DataCombo_CustomerType.Text = Trim(Form10.DataCombo_CustomerTypex.Text)
xxx = "Customer_Type = '" & Trim(Form10.DataCombo_CustomerTypex.Text) & "'"
If Trim(Form10.DataCombo_CustomerTypex.Text) = "Guarantee/Warranty" Then
    Form10.DataCombo1.ReFill
    Set Form10.DataCombo1.DataSource = Adodc1
    Set Form10.DataCombo1.RowSource = Adodc1
    Set Form10.DataCombo_StaffID.DataSource = Adodc1
    Set Form10.DataCombo_StaffID.RowSource = Adodc1
    Form10.Adodc1.Refresh
    Form10.Adodc1.Recordset.Requery
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo_StaffID.Refresh
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo1.ReFill
    Form10.DataCombo1.Refresh
    Form10.DataCombo1.ReFill
ElseIf Trim(Form10.DataCombo_CustomerTypex.Text) = "General" Then
    Form10.DataCombo1.ReFill
    Set Form10.DataCombo1.DataSource = Adodc2
    Set Form10.DataCombo1.RowSource = Adodc2
    Set Form10.DataCombo_StaffID.DataSource = Adodc2
    Set Form10.DataCombo_StaffID.RowSource = Adodc2
    Form10.Adodc2.Refresh
    Form10.Adodc2.Recordset.Requery
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo_StaffID.Refresh
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo1.ReFill
    Form10.DataCombo1.Refresh
    Form10.DataCombo1.ReFill
Else
    Form10.DataCombo1.ReFill
    Set Form10.DataCombo1.DataSource = Adodc3
    Set Form10.DataCombo1.RowSource = Adodc3
    Set Form10.DataCombo_StaffID.DataSource = Adodc3
    Set Form10.DataCombo_StaffID.RowSource = Adodc3
    Form10.Adodc3.Refresh
    Form10.Adodc3.Recordset.Requery
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo_StaffID.Refresh
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo1.ReFill
    Form10.DataCombo1.Refresh
    Form10.DataCombo1.ReFill
End If
    Fx = Adodc3x.Recordset.Fields(1).Name & " ="
    Dx = Trim(Form10.DataCombo_StaffID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc3x.Refresh
    Adodc3x.Recordset.Filter = stringTx
    If Form10.Adodc3x.Recordset.EOF = True Then
        Exit Sub
    End If
    Form10.DataGrid1.ReBind
    Call Grep_ServiceSlNoForCustomer
    'Form10.DataCombo1.Text = ""
End Sub

Private Sub DataCombo_CustomerTypex_KeyUp(KeyCode As Integer, Shift As Integer)
Form10.DataCombo_CustomerType.Text = Trim(Form10.DataCombo_CustomerTypex.Text)
xxx = "Customer_Type = '" & Trim(Form10.DataCombo_CustomerTypex.Text) & "'"
If Trim(Form10.DataCombo_CustomerTypex.Text) = "Guarantee/Warranty" Then
    Form10.DataCombo1.ReFill
    Set Form10.DataCombo1.DataSource = Adodc1
    Set Form10.DataCombo1.RowSource = Adodc1
    Set Form10.DataCombo_StaffID.DataSource = Adodc1
    Set Form10.DataCombo_StaffID.RowSource = Adodc1
    Form10.Adodc1.Refresh
    Form10.Adodc1.Recordset.Requery
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo_StaffID.Refresh
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo1.ReFill
    Form10.DataCombo1.Refresh
    Form10.DataCombo1.ReFill
ElseIf Trim(Form10.DataCombo_CustomerTypex.Text) = "General" Then
    Form10.DataCombo1.ReFill
    Set Form10.DataCombo1.DataSource = Adodc2
    Set Form10.DataCombo1.RowSource = Adodc2
    Set Form10.DataCombo_StaffID.DataSource = Adodc2
    Set Form10.DataCombo_StaffID.RowSource = Adodc2
    Form10.Adodc2.Refresh
    Form10.Adodc2.Recordset.Requery
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo_StaffID.Refresh
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo1.ReFill
    Form10.DataCombo1.Refresh
    Form10.DataCombo1.ReFill
Else
    Form10.DataCombo1.ReFill
    Set Form10.DataCombo1.DataSource = Adodc3
    Set Form10.DataCombo1.RowSource = Adodc3
    Set Form10.DataCombo_StaffID.DataSource = Adodc3
    Set Form10.DataCombo_StaffID.RowSource = Adodc3
    Form10.Adodc3.Refresh
    Form10.Adodc3.Recordset.Requery
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo_StaffID.Refresh
    Form10.DataCombo_StaffID.ReFill
    Form10.DataCombo1.ReFill
    Form10.DataCombo1.Refresh
    Form10.DataCombo1.ReFill
End If
    Fx = Adodc3x.Recordset.Fields(1).Name & " ="
    Dx = Trim(Form10.DataCombo_StaffID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc3x.Refresh
    Adodc3x.Recordset.Filter = stringTx
    If Form10.Adodc3x.Recordset.EOF = True Then
        Exit Sub
    End If
    Form10.DataGrid1.ReBind
    Call Grep_ServiceSlNoForCustomer
'Form10.DataCombo1.Text = ""
End Sub

Private Sub DataCombo_ServiceSlNo_Click()
    strTx = "service_serial_no = '" & Trim(Form10.DataCombo_ServiceSlNo.Text) & "'"
    Adodc3x.Refresh
    Adodc3x.Recordset.Filter = strTx
    If Form10.Adodc3x.Recordset.EOF = True Then
        Exit Sub
    End If
End Sub

Private Sub DataCombo_ServiceSlNo_KeyUp(KeyCode As Integer, Shift As Integer)
'    strTx = "service_serial_no = '" & Trim(Form10.DataCombo_ServiceSlNo.Text) & "'"
'   Adodc3x.Refresh
'    Adodc3x.Recordset.Filter = strTx
'    If Form10.Adodc3x.Recordset.EOF = True Then
'        Exit Sub
'    End If
End Sub

Private Sub DataCombo_StaffID_KeyUp(KeyCode As Integer, Shift As Integer)
If Trim(Form10.DataCombo_CustomerType.Text) <> "" Then
    Fx = "Customer_ID ="
    Dx = Trim(Form10.DataCombo_StaffID.Text)
    stringTx = Fx & "'" & Dx & "'"
    If Trim(Form10.DataCombo_CustomerTypex.Text) = "Guarantee/Warranty" Then
        Adodc1.Refresh
        Adodc1.Recordset.Filter = stringTx
        If Form10.Adodc1.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
            Call Service_Details_Refresh
            Call Grep_ServiceSlNoForCustomer
        End If
    ElseIf Trim(Form10.DataCombo_CustomerTypex.Text) = "General" Then
        Adodc2.Refresh
        Adodc2.Recordset.Filter = stringTx
        If Form10.Adodc2.Recordset.EOF = True Then
        Else
            Call Show_CustomerContactInfo
            Call Service_Details_Refresh
            Call Grep_ServiceSlNoForCustomer
            Exit Sub
        End If
    Else
        Adodc3.Refresh
        Adodc3.Recordset.Filter = stringTx
        If Form10.Adodc3.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
            Call Service_Details_Refresh
            Call Grep_ServiceSlNoForCustomer
        End If
    End If
End If
End Sub

Private Sub DataCombo_StaffID_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Trim(Form10.DataCombo_CustomerType.Text) <> "" Then
    Fx = "Customer_ID ="
    Dx = Trim(Form10.DataCombo_StaffID.Text)
    stringTx = Fx & "'" & Dx & "'"
    If Trim(Form10.DataCombo_CustomerTypex.Text) = "Guarantee/Warranty" Then
        Adodc1.Refresh
        Adodc1.Recordset.Filter = stringTx
        If Form10.Adodc1.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
            Call Service_Details_Refresh
            Call Grep_ServiceSlNoForCustomer
        End If
    ElseIf Trim(Form10.DataCombo_CustomerTypex.Text) = "General" Then
        Adodc2.Refresh
        Adodc2.Recordset.Filter = stringTx
        If Form10.Adodc2.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
            Call Service_Details_Refresh
            Call Grep_ServiceSlNoForCustomer
        End If
    Else
        Adodc3.Refresh
        Adodc3.Recordset.Filter = stringTx
        If Form10.Adodc3.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
            Call Service_Details_Refresh
            Call Grep_ServiceSlNoForCustomer
        End If
    End If
End If
End Sub

Private Sub DataCombo1_LostFocus()
If Trim(Form10.DataCombo_CustomerType.Text) <> "" Then
    Fx = "Customer_Name = "
    Dx = Trim(Form10.DataCombo1.Text)
    stringTx = Fx & "'" & Dx & "'"
    If Trim(Form10.DataCombo_CustomerTypex.Text) = "Guarantee/Warranty" Then
        Adodc1.Refresh
        Adodc1.Recordset.Filter = stringTx
        If Form10.Adodc1.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
        End If
    ElseIf Trim(Form10.DataCombo_CustomerTypex.Text) = "General" Then
        Adodc2.Refresh
        Adodc2.Recordset.Filter = stringTx
        If Form10.Adodc2.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
        End If
    Else
        Adodc3.Refresh
        Adodc3.Recordset.Filter = stringTx
        If Form10.Adodc3.Recordset.EOF = True Then
            Exit Sub
        Else
            Call Show_CustomerContactInfo
        End If
    End If
    Call Grep_ServiceSlNoForCustomer
End If
End Sub

Private Sub DataCombo1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Trim(Form10.DataCombo_CustomerType.Text) <> "" Then
    Fx = "Customer_Name = "
    Dx = Trim(Form10.DataCombo1.Text)
    stringTx = Fx & "'" & Dx & "'"
    If Trim(Form10.DataCombo_CustomerTypex.Text) = "Guarantee/Warranty" Then
        Adodc1.Refresh
        Adodc1.Recordset.Filter = stringTx
        If Form10.Adodc1.Recordset.EOF = True Then
            Exit Sub
        End If
    ElseIf Trim(Form10.DataCombo_CustomerTypex.Text) = "General" Then
        Adodc2.Refresh
        Adodc2.Recordset.Filter = stringTx
        If Form10.Adodc2.Recordset.EOF = True Then
            Exit Sub
        End If
    Else
        Adodc3.Refresh
        Adodc3.Recordset.Filter = stringTx
        If Form10.Adodc3.Recordset.EOF = True Then
            Exit Sub
        End If
    End If
    Call Grep_ServiceSlNoForCustomer
End If
End Sub

Private Sub Form_Load()
Set Form10.DataCombo1.DataSource = Adodcx
Set Form10.DataCombo1.RowSource = Adodcx
Set Form10.DataCombo_StaffID.DataSource = Adodcx
'Set Form10.DataCombo_StaffID.RowSource = Adodcx
Form10.DataCombo_CustomerType.Text = ""
'Form10.DataCombo_CustomerTypex.Text = ""
Form10.DataCombo1.Text = ""
Form10.DataCombo_StaffID.Text = ""
End Sub


Public Sub Show_CustomerContactInfo()
axn = Trim(Form10.DataCombo_CustomerTypex.Text)
If Trim(axn) <> "" Then
    Fx = "Customer_ID = "
    Dx = Trim(Form10.DataCombo_StaffID.Text)
    stringTx = Fx & "'" & Dx & "'"
    Adodcx.Refresh
    Adodcx.Recordset.Filter = stringTx
    If Form10.Adodcx.Recordset.EOF = True Then
        Exit Sub
    Else
        With Adodcx
            Form10.Text_PAddress.Text = Trim(.Recordset.Fields(3).Value)
        End With
    End If
End If
End Sub


Public Sub Grep_ServiceSlNoForCustomer()
cid = Trim(Form10.DataCombo_StaffID.Text)
Form10.DataCombo_ServiceSlNo.Clear
If Trim(cid) <> "" Then
    Form10.Adodc2x.Refresh
    rc = Val(Form10.Adodc2x.Recordset.RecordCount)
    If Val(rc) > 0 Then
        s = "customer_id = '" & Trim(cid) & "'"
        Form10.Adodc2x.Refresh
        Form10.Adodc2x.Recordset.Find s
        If Form10.Adodc2x.Recordset.EOF = True Then
            Exit Sub
        Else
            Form10.DataCombo_ServiceSlNo.AddItem Trim(Form10.Adodc2x.Recordset.Fields(0).Value)
        End If
upscx:
        Form10.Adodc2x.Recordset.MoveNext
        If Form10.Adodc2x.Recordset.EOF = True Then
            Exit Sub
        Else
            If Trim(cid) = Trim(Form10.Adodc2x.Recordset.Fields(3).Value) Then
                Form10.DataCombo_ServiceSlNo.AddItem Trim(Form10.Adodc2x.Recordset.Fields(3).Value)
            Else
                GoTo upscx
            End If
        End If
    End If
End If
End Sub
Private Sub Timer1_Timer()
Form10.Label_DateOfPurchase.Caption = Format(Date, "dd/MM/yyyy")
Form10.Label_TimeOfPurchase.Caption = Format(Now, "h:mm:ss AM/PM")
End Sub

Public Sub Service_Details_Refresh()
    Form10.DataCombo_ServiceSlNo.Clear
    strTx = "service_serial_no = 0"
    Adodc3x.Refresh
    Adodc3x.Recordset.Filter = strTx
    If Form10.Adodc3x.Recordset.EOF = True Then
        Form10.DataGrid1.Refresh
        Exit Sub
    End If
End Sub
 

