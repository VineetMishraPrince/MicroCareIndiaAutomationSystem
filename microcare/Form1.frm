VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Staff's Details"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text_Sex 
      Height          =   285
      Left            =   6240
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3360
      Top             =   8160
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.CommandButton Command_Cancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   24
      Top             =   7370
      Width           =   1335
   End
   Begin VB.CommandButton Command_Delete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   23
      Top             =   7370
      Width           =   1335
   End
   Begin VB.CommandButton Command_Update 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   22
      Top             =   6750
      Width           =   1335
   End
   Begin VB.CommandButton Command_Add 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   6750
      Width           =   1335
   End
   Begin VB.PictureBox Picture_Photo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3760
      Left            =   8520
      ScaleHeight     =   3735
      ScaleWidth      =   2865
      TabIndex        =   20
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton Command_PastePhoto 
      Caption         =   "&Paste Photo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   19
      Top             =   2270
      Width           =   1335
   End
   Begin VB.OptionButton Option_F 
      BackColor       =   &H80000010&
      Caption         =   "Option2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   10080
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton Option_M 
      BackColor       =   &H80000010&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text_Description 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Text            =   "Form1.frx":E332
      Top             =   6840
      Width           =   5175
   End
   Begin VB.TextBox Text_Date_of_joining 
      Height          =   400
      Left            =   3120
      TabIndex        =   15
      Text            =   "Text14"
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox Text_Nationality 
      Height          =   400
      Left            =   6720
      TabIndex        =   14
      Text            =   "Text13"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text_DateOFBirth 
      Height          =   400
      Left            =   3120
      TabIndex        =   13
      Text            =   "Text12"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text_EmailID 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3120
      TabIndex        =   12
      Text            =   "Text11"
      Top             =   5420
      Width           =   5175
   End
   Begin VB.TextBox Text_FaxNo 
      Height          =   400
      Left            =   6720
      TabIndex        =   11
      Text            =   "Text10"
      Top             =   4950
      Width           =   1575
   End
   Begin VB.TextBox Text_MobileNo 
      Height          =   400
      Left            =   3120
      TabIndex        =   10
      Text            =   "Text9"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text_Tel_O 
      Height          =   400
      Left            =   6720
      TabIndex        =   9
      Text            =   "Text8"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text_TeleNo_R 
      Height          =   400
      Left            =   3120
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text_C_Address 
      Height          =   640
      Left            =   3120
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   3720
      Width           =   5175
   End
   Begin VB.TextBox Text_PAddress 
      Height          =   640
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":E339
      Top             =   3000
      Width           =   5175
   End
   Begin VB.TextBox Text_Qualification 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2520
      Width           =   5175
   End
   Begin VB.TextBox Text_specialization 
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text_Designation 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text_Name 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1530
      Width           =   5175
   End
   Begin VB.ComboBox DataCombo_StaffID 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stringTx, Fx, Dx As String

Private Sub Command_Add_Click()
'Call Soft_Expiry
'    Text_PastePhotoComment.Visible = True
    Command_PastePhoto.Enabled = True
    'Frame1_ChooseSex.Visible = True
    Form1.Adodc1.Recordset.AddNew
    Form1.Command_Add.Enabled = False
End Sub

Private Sub Command_Cancel_Click()
    Unload Form1
End Sub

Private Sub Command_Delete_Click()
Dim respn As Integer
respn = MsgBox("Are you sure to delete this record?", vbYesNo, "Message")
If respn = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
End If
End Sub


Private Sub Command_PastePhoto_Click()
    F = "C:\My Documents\My Pictures\HR\x.jpg"
    Form1.Picture_Photo.Picture = LoadPicture(F)
    'Text_PastePhotoComment.Visible = False
End Sub

Private Sub Command_Update_Click()
'On Error GoTo dn
Dim resp As Integer
resp = MsgBox("Are you sure to update the record?", vbYesNo + vbQuestion, "Message")
If resp = vbYes And Form1.Command_Add.Enabled = False Then
    Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Form1.DataCombo_StaffID.Text
    stringTx = Fx + "'" + Dx + "'"
    Form1.Adodc1.Recordset.Fields("Employee_ID").Value = Trim(Form1.DataCombo_StaffID.Text)
    Form1.Adodc1.Recordset.Fields("Employee_Name").Value = Trim(Form1.Text_Name.Text)
    Form1.Adodc1.Recordset.Fields("Sex").Value = Trim(Form1.Text_Sex.Text)
    Form1.Adodc1.Recordset.Fields("Date_of_Birth").Value = Trim(Form1.Text_DateOFBirth.Text)
    Form1.Adodc1.Recordset.Fields("DESIGNATION").Value = Trim(Form1.Text_Designation.Text)
    Form1.Adodc1.Recordset.Fields("specialization").Value = Trim(Form1.Text_specialization.Text)
    Form1.Adodc1.Recordset.Fields("Qualification").Value = Trim(Form1.Text_Qualification.Text)
    Form1.Adodc1.Recordset.Fields("Nationality").Value = Trim(Form1.Text_Nationality.Text)
    Form1.Adodc1.Recordset.Fields("Permanent_Address").Value = Trim(Form1.Text_PAddress.Text)
    Form1.Adodc1.Recordset.Fields("Correspondence_Address").Value = Trim(Form1.Text_C_Address.Text)
    Form1.Adodc1.Recordset.Fields("Phone_No_R").Value = Trim(Form1.Text_TeleNo_R.Text)
    Form1.Adodc1.Recordset.Fields("Phone_No_O").Value = Trim(Form1.Text_Tel_O.Text)
    Form1.Adodc1.Recordset.Fields("Mobile").Value = Trim(Form1.Text_MobileNo.Text)
    Form1.Adodc1.Recordset.Fields("Email").Value = Trim(Form1.Text_EmailID.Text)
    Form1.Adodc1.Recordset.Fields("Fax_No").Value = Trim(Form1.Text_FaxNo.Text)
    Form1.Adodc1.Recordset.Fields("Date_of_Joining").Value = Trim(Form1.Text_Date_of_joining.Text)
    Form1.Adodc1.Recordset.Fields("Description").Value = Trim(Form1.Text_Description.Text)
    Form1.Adodc1.Recordset.UpdateBatch
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    F = "C:\My Documents\My Pictures\HR\x.jpg"
    If F <> "" Then
        R = MsgBox("Are You Sure to update this Photo?", vbYesNo, "Message")
            If R = vbYes Then
                Adodc1.Recordset.Find (stringTx)
                    Open F For Binary Access Read As #1
                    ReDim img(FileLen(F) - 1)
                    Get #1, , img()
                    Close #1
                    Adodc1.Recordset.Fields("S_Photo").AppendChunk img
                    Picture_Photo.Picture = LoadPicture(F)
                    Adodc1.Recordset.MoveFirst
                    Adodc1.Recordset.MoveLast
                    Adodc1.Refresh
                    Adodc1.Recordset.Find (stringTx)
        End If
    End If
ElseIf resp = vbYes And Form1.Command_Add.Enabled = True Then
    Form1.Adodc1.Recordset.Fields("Employee_ID").Value = Trim(Form1.DataCombo_StaffID.Text)
    Form1.Adodc1.Recordset.Fields("Employee_Name").Value = Trim(Form1.Text_Name.Text)
    Form1.Adodc1.Recordset.Fields("Sex").Value = Trim(Form1.Text_Sex.Text)
    Form1.Adodc1.Recordset.Fields("Date_of_Birth").Value = Trim(Form1.Text_DateOFBirth.Text)
    Form1.Adodc1.Recordset.Fields("DESIGNATION").Value = Trim(Form1.Text_Designation.Text)
    Form1.Adodc1.Recordset.Fields("specialization").Value = Trim(Form1.Text_specialization.Text)
    Form1.Adodc1.Recordset.Fields("Qualification").Value = Trim(Form1.Text_Qualification.Text)
    Form1.Adodc1.Recordset.Fields("Nationality").Value = Trim(Form1.Text_Nationality.Text)
    Form1.Adodc1.Recordset.Fields("Permanent_Address").Value = Trim(Form1.Text_PAddress.Text)
    Form1.Adodc1.Recordset.Fields("Correspondence_Address").Value = Trim(Form1.Text_C_Address.Text)
    Form1.Adodc1.Recordset.Fields("Phone_No_R").Value = Trim(Form1.Text_TeleNo_R.Text)
    Form1.Adodc1.Recordset.Fields("Phone_No_O").Value = Trim(Form1.Text_Tel_O.Text)
    Form1.Adodc1.Recordset.Fields("Mobile").Value = Trim(Form1.Text_MobileNo.Text)
    Form1.Adodc1.Recordset.Fields("Email").Value = Trim(Form1.Text_EmailID.Text)
    Form1.Adodc1.Recordset.Fields("Fax_No").Value = Trim(Form1.Text_FaxNo.Text)
    Form1.Adodc1.Recordset.Fields("Date_of_Joining").Value = Trim(Form1.Text_Date_of_joining.Text)
    Form1.Adodc1.Recordset.Fields("Description").Value = Trim(Form1.Text_Description.Text)
    Form1.Adodc1.Recordset.UpdateBatch
    F = "C:\My Documents\My Pictures\HR\x.jpg"
    If F <> "" Then
        R = MsgBox("Are You Sure to update this Photo?", vbYesNo, "Message")
            If R = vbYes Then
                Fx = Adodc1.Recordset.Fields(0).Name & " ="
                Dx = Trim(Form1.DataCombo_StaffID.Text)
                stringTx = Fx + "'" + Dx + "'"
                Adodc1.Refresh
                Adodc1.Recordset.Find (stringTx)
                    Open F For Binary Access Read As #1
                    ReDim img(FileLen(F) - 1)
                    Get #1, , img()
                    Close #1
                    Adodc1.Recordset.Fields("S_Photo").AppendChunk img
                    Form1.Picture_Photo.Picture = LoadPicture(F)
                    Adodc1.Recordset.MoveFirst
                    Adodc1.Recordset.MoveLast
                    Adodc1.Refresh
                    Adodc1.Recordset.Find (stringTx)
            End If
    End If
End If
Form1.Command_Add.Enabled = True
Text_PastePhotoComment.Visible = False
Exit Sub
dn:
MsgBox "Invalid Entry...! Try Again.", vbExclamation, "ERROR"
Exit Sub
End Sub

Private Sub DataCombo_StaffID_KeyUp(KeyCode As Integer, Shift As Integer)
If Form1.Command_Add.Enabled = True Then
    Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Trim(Form1.DataCombo_StaffID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    If Form1.Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
End If
End Sub

Private Sub DataCombo_StaffID_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Form1.Command_Add.Enabled = True Then
    Fx = Adodc1.Recordset.Fields(0).Name & " ="
    Dx = Trim(Form1.DataCombo_StaffID.Text)
    stringTx = Fx + "'" + Dx + "'"
    Adodc1.Refresh
    Adodc1.Recordset.Find (stringTx)
    If Form1.Adodc1.Recordset.EOF = True Then
        Exit Sub
    End If
End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
'    Text_PastePhotoComment.Visible = False
    'Frame1_ChooseSex.Visible = False
    Call Text_Sex_Change
End Sub

Private Sub Option_F_Click()
    Form1.Text_Sex.Text = "F"
End Sub

Private Sub Option_M_Click()
    Form1.Text_Sex.Text = "M"
End Sub

Private Sub Text_Date_of_joining_LostFocus()
If Trim(Form1.Text_Date_of_joining.Text) <> "" And IsDate(Form1.Text_Date_of_joining.Text) = True Then
    Form1.Text_Date_of_joining.Text = Format(Form1.Text_Date_of_joining.Text, "d-mmm-yyyy")
End If
End Sub

Private Sub Text_DateOFBirth_LostFocus()
If Trim(Text_DateOFBirth.Text) <> "" And IsDate(Text_DateOFBirth.Text) = True Then
    Text_DateOFBirth.Text = Format(Text_DateOFBirth.Text, "d-mmm-yyyy")
End If
End Sub

Private Sub Text_Sex_Change()
If Trim(UCase(Form1.Text_Sex.Text)) = "M" Then
    Option_M.Value = True
Else
    Option_F.Value = True
End If
End Sub
 

Private Sub Text14_Change()

End Sub
