VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form12 
   Caption         =   "Profit/Loss Information"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form12"
   MDIChild        =   -1  'True
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CommandButton_Go 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command_Cancel 
      Caption         =   "Close"
      Height          =   735
      Left            =   5280
      TabIndex        =   3
      Top             =   7220
      Width           =   1095
   End
   Begin VB.ListBox ListBox1 
      Appearance      =   0  'Flat
      Height          =   5295
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19202051
      CurrentDate     =   38857
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   400
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19202051
      CurrentDate     =   38857
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Cancel_Click()
Unload Me
End Sub

Private Sub CommandButton_Go_Click()
Form12.ListBox1.Visible = True
Form12.ListBox1.Clear
sd = Format(CDate(Form12.DTPicker1.Value), "dd-MMM-yyyy")
ed = Format(CDate(Form12.DTPicker2.Value), "dd-MMM-yyyy")
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
conn.Open "Provider=MSDAORA.1;User ID=ccsms;PASSWORD=zzz;Persist Security Info=False"
Set cmd.ActiveConnection = conn
DP = "select sum(Paid_amount) from Employee_Payroll where Payment_Date >= '" & Format(CDate(sd), "dd-MMM-yyyy") & "' and Payment_Date <= '" & Format(CDate(ed), "dd-MMM-yyyy") & "'"
cmd.CommandText = DP
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockBatchOptimistic
If rs.RecordCount > 0 Then
    rs.Requery
    If rs.EOF = False Then
        D_Payment = rs.GetString(adClipString)
        GoTo gsd
    Else
        D_Payment = 0
    End If
gsd:
    Form12.ListBox1.AddItem "Payroll:->"
    Form12.ListBox1.AddItem "       To Employees: " & Val(D_Payment)
End If
rs.Close
conn.Close
Form12.ListBox1.AddItem "-----------------------------------------"
conn.Open "Provider=MSDAORA.1;User ID=ccsms;PASSWORD=zzz;Persist Security Info=False"
Set cmd.ActiveConnection = conn
Services_charges = "select sum(payed_amount) from SERVICE_CHARGES"
cmd.CommandText = Services_charges
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockBatchOptimistic
If rs.RecordCount > 0 Then
    rs.Requery
    If rs.EOF = False Then
        Services_charges = rs.GetString(adClipString)
    Else
        Services_charges = 0
    End If
    Form12.ListBox1.AddItem "Services charges:-> "
    Form12.ListBox1.AddItem "       Charge = " & Val(Services_charges)
End If
rs.Close
conn.Close
Form12.ListBox1.AddItem "-----------------------------------------"
conn.Open "Provider=MSDAORA.1;User ID=ccsms;PASSWORD=zzz;Persist Security Info=False"
Set cmd.ActiveConnection = conn
Appt = "select sum(price) from WARRENTY_PRODUCTS where warrenty_start_date >= '" & Format(CDate(sd), "dd-MMM-yyyy") & "' and warrenty_start_date <= '" & Format(CDate(ed), "dd-MMM-yyyy") & "'"
cmd.CommandText = Appt
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockBatchOptimistic
If rs.RecordCount > 0 Then
    rs.Requery
    If rs.EOF = True Then
        App_Charge = 0
        GoTo l
    Else
        On Error GoTo l
        App_Charge = rs.GetString(adClipString)
    End If
l:
    Form12.ListBox1.AddItem "Net Promised Products/Services:-> "
    Form12.ListBox1.AddItem "       Amount = " & Val(App_Charge)
End If
rs.Close
conn.Close
Form12.ListBox1.AddItem "-----------------------------------------"
pfa = Val(Val(App_Charge) + Val(Services_charges)) - Val(D_Payment)
Form12.ListBox1.AddItem "-----------------------------------------"
Form12.ListBox1.AddItem "Profit/Loss Amount = " & pfa
End Sub

Private Sub DTPicker1_Change()
If CDate(Form12.DTPicker1.Value) > CDate(Form12.DTPicker2.Value) Then
    Form12.DTPicker1.Value = Form12.DTPicker2.Value
End If
'Form12.ListBox1.Visible = False
End Sub
Private Sub DTPicker1_Click()
'Form12.Shape1.Visible = False
End Sub

Private Sub DTPicker2_Change()
If CDate(Form12.DTPicker2.Value) > CDate(Date) Then
    Form12.DTPicker2.Value = Date
End If
'Form12.ListBox1.Visible = False
End Sub

Private Sub DTPicker2_Click()
'Form12.Shape1.Visible = False
End Sub

Private Sub Form_Load()
Form12.DTPicker1.Value = Date
Form12.DTPicker2.Value = Date
'Form12.ListBox1.Visible = False
End Sub
 

