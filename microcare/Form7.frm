VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Caption         =   "Customers Information"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   8040
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
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1920
   End
   Begin VB.CommandButton Command_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   10100
      TabIndex        =   15
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command_Delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   7840
      TabIndex        =   14
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command_Update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   5550
      TabIndex        =   13
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command_Add 
      Caption         =   "&ADD"
      Height          =   375
      Left            =   3275
      TabIndex        =   12
      Top             =   7320
      Width           =   1335
   End
   Begin VB.TextBox Text_Description 
      Height          =   855
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "Form7.frx":1370F
      Top             =   5950
      Width           =   5320
   End
   Begin VB.TextBox Text_EmailID 
      Height          =   400
      Left            =   3240
      TabIndex        =   10
      Text            =   "Text7"
      Top             =   5200
      Width           =   5320
   End
   Begin VB.TextBox Text_Date_of_joining 
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   4650
      Width           =   1875
   End
   Begin VB.TextBox Text_FaxNo 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   4650
      Width           =   1850
   End
   Begin VB.TextBox Text_MobileNo 
      Height          =   420
      Left            =   6720
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   4000
      Width           =   1815
   End
   Begin VB.TextBox Text_TeleNo 
      Height          =   420
      Left            =   3240
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   4000
      Width           =   1935
   End
   Begin VB.TextBox Text_PAddress 
      Height          =   630
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form7.frx":13715
      Top             =   3200
      Width           =   5295
   End
   Begin VB.TextBox Text_Name 
      Height          =   390
      Left            =   3240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2600
      Width           =   5295
   End
   Begin VB.ComboBox DataCombo_CustomerType 
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox DataCombo_StaffID 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   9840
      TabIndex        =   17
      Top             =   1150
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   9840
      TabIndex        =   16
      Top             =   520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stringTx, Fx, Dx As String

Private Sub Command_Add_Click()
'Call Soft_Expiry
    Form7.Adodc1.Refresh
    v = Form7.Adodc1.Recordset.RecordCount
    Form7.Adodc1.Recordset.AddNew
    Form7.DataCombo_StaffID.Text = Val(v) + 1
    Form7.DataCombo_StaffID.Enabled = False
    Form7.Command_Add.Enabled = False
End Sub

Private Sub Command_Cancel_Click()
    Unload Form7
End Sub

Private Sub Command_Delete_Click()
Dim respn As Integer
respn = MsgBox("Are you sure to delete this record?", vbYesNo, "Message")
If respn = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
End If
End Sub

Private Sub Command_Update_Click()
'On Error GoTo dn
Dim resp As Integer
resp = MsgBox("Are you sure to update the record?", vbYesNo + vbQuestion, "Message")
If resp = vbYes And Form7.Command_Add.Enabled = False Then
    Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Form7.DataCombo_StaffID.Text
    stringTx = Fx + "'" + Dx + "'"
    Form7.Adodc1.Recordset.Fields(0).Value = Trim(Form7.DataCombo_StaffID.Text)
    Form7.Adodc1.Recordset.Fields(1).Value = Trim(Form7.DataCombo_CustomerType.Text)
    Form7.Adodc1.Recordset.Fields(2).Value = Trim(Form7.Text_Name.Text)
    Form7.Adodc1.Recordset.Fields(3).Value = Trim(Form7.Text_PAddress.Text)
    Form7.Adodc1.Recordset.Fields(4).Value = Trim(Form7.Text_TeleNo_R.Text)
    Form7.Adodc1.Recordset.Fields(5).Value = Trim(Form7.Text_MobileNo.Text)
    Form7.Adodc1.Recordset.Fields(6).Value = Trim(Form7.Text_EmailID.Text)
    Form7.Adodc1.Recordset.Fields(7).Value = Trim(Form7.Text_FaxNo.Text)
    Form7.Adodc1.Recordset.Fields(8).Value = Trim(Form7.Text_Date_of_joining.Text)
    Form7.Adodc1.Recordset.Fields(9).Value = Trim(Form7.Text_Description.Text)
    Form7.Adodc1.Recordset.UpdateBatch
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    Form7.DataCombo_StaffID.Enabled = True
ElseIf resp = vbYes And Form7.Command_Add.Enabled = True Then
Form7.Adodc1.Recordset.Fields(0).Value = Trim(Form7.DataCombo_StaffID.Text)
    Form7.Adodc1.Recordset.Fields(1).Value = Trim(Form7.DataCombo_CustomerType.Text)
    Form7.Adodc1.Recordset.Fields(2).Value = Trim(Form7.Text_Name.Text)
    Form7.Adodc1.Recordset.Fields(3).Value = Trim(Form7.Text_PAddress.Text)
    Form7.Adodc1.Recordset.Fields(4).Value = Trim(Form7.Text_TeleNo_R.Text)
    Form7.Adodc1.Recordset.Fields(5).Value = Trim(Form7.Text_MobileNo.Text)
    Form7.Adodc1.Recordset.Fields(6).Value = Trim(Form7.Text_EmailID.Text)
    Form7.Adodc1.Recordset.Fields(7).Value = Trim(Form7.Text_FaxNo.Text)
    Form7.Adodc1.Recordset.Fields(8).Value = Trim(Form7.Text_Date_of_joining.Text)
    Form7.Adodc1.Recordset.Fields(9).Value = Trim(Form7.Text_Description.Text)
    Form7.Adodc1.Recordset.UpdateBatch
    Form7.DataCombo_StaffID.Enabled = True
End If
Form7.Command_Add.Enabled = True
Exit Sub
dn:
MsgBox "Invalid Entry...! Try Again.", vbExclamation, "ERROR"
Exit Sub
End Sub

Private Sub DataCombo_CustomerTypex_Change()
'Form7.DataCombo_CustomerType.Text = Trim(Form7.DataCombo_CustomerTypex.Text)
End Sub

Private Sub DataCombo_CustomerTypex_Click()
Form7.DataCombo_CustomerType.Text = Trim(Form7.DataCombo_CustomerTypex.Text)
End Sub

Private Sub DataCombo_CustomerTypex_KeyUp(KeyCode As Integer, Shift As Integer)
Form7.DataCombo_CustomerType.Text = Trim(Form7.DataCombo_CustomerTypex.Text)
End Sub




Private Sub DataCombo_StaffID_KeyUp(KeyCode As Integer, Shift As Integer)
If Form7.Command_Add.Enabled = True Then
    Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Trim(Form7.DataCombo_StaffID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    If Form7.Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
End If
End Sub

Private Sub DataCombo_StaffID_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Form7.Command_Add.Enabled = True Then
    Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Trim(Form7.DataCombo_StaffID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    If Form7.Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
End If
End Sub

Private Sub Text_Date_of_joining_LostFocus()
If Trim(Form7.Text_Date_of_joining.Text) <> "" And IsDate(Form7.Text_Date_of_joining.Text) = True Then
    Form7.Text_Date_of_joining.Text = Format(Form7.Text_Date_of_joining.Text, "d-mmm-yyyy")
End If
End Sub

Private Sub Text_DateOFBirth_LostFocus()
If Trim(Text_DateOFBirth.Text) <> "" And IsDate(Text_DateOFBirth.Text) = True Then
    Text_DateOFBirth.Text = Format(Text_DateOFBirth.Text, "d-mmm-yyyy")
End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Timer1_Timer()
Form7.Label_DateOfPurchase.Caption = " " & Format(Date, "dd/MM/yyyy") & " "
Form7.Label_TimeOfPurchase.Caption = " " & Format(Now, "h:mm:ss AM/PM") & " "
End Sub

