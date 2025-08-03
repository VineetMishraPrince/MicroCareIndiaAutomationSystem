VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Warrenty Product/Service Listing"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5040
      Top             =   8160
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.CommandButton Command_Cancel 
      Caption         =   "&CANCEL"
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   7100
      Width           =   1335
   End
   Begin VB.CommandButton Command_Update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   7100
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form9.frx":12F76
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   5230
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3201
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
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ComboBox DataCombo_CustomerID 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   8160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Left            =   4200
      Top             =   8160
   End
   Begin VB.Label Label_TimeOfPurchase 
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   10320
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label_DateOfPurchase 
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   10320
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1500
      Width           =   2295
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Cancel_Click()
Unload Me
End Sub

Private Sub Command_Update_Click()
    cid = Trim(Form9.DataCombo_CustomerID.Text)
    'Form9.Adodc2.Refresh
    'Form9.Adodc2.Recordset.AddNew
    Form9.DataGrid1.Columns(0).Value = Trim(cid)
    Form9.Adodc2.Recordset.MoveNext
    If Form9.Adodc2.Recordset.EOF = True Then
        Exit Sub
    Else
        Form9.DataGrid1.Refresh
    End If
End Sub

Private Sub DataCombo_CustomerID_KeyUp(KeyCode As Integer, Shift As Integer)
   Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Trim(Form9.DataCombo_CustomerID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    If Form9.Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
    Adodc2.Refresh
    Adodc2.Recordset.Filter = stringTx
    If Form9.Adodc2.Recordset.EOF = True Then
        Exit Sub
    End If
    Form9.DataGrid1.ReBind
End Sub

Private Sub DataCombo_CustomerID_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Trim(Form9.DataCombo_CustomerID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    If Form9.Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
    Adodc2.Refresh
    Adodc2.Recordset.Filter = stringTx
    If Form9.Adodc2.Recordset.EOF = True Then
        Exit Sub
    End If
    Form9.DataGrid1.ReBind
End Sub

Private Sub Form_Load()
   Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Trim(Form9.DataCombo_CustomerID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    If Form9.Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
    Adodc2.Refresh
    Adodc2.Recordset.Filter = stringTx
    If Form9.Adodc2.Recordset.EOF = True Then
        Exit Sub
    End If
    Form9.DataGrid1.ReBind
End Sub

Private Sub Timer1_Timer()
Form9.Label_DateOfPurchase.Caption = Format(Date, "dd/MM/yyyy")
Form9.Label_TimeOfPurchase.Caption = Format(Now, "h:mm:ss AM/PM")
End Sub
 

