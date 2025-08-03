VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form8 
   Caption         =   "Service Wizard Window - 2/2 [Service]"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "form2.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   9960
      TabIndex        =   20
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<< BACK"
      Height          =   375
      Left            =   8400
      TabIndex        =   19
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   8400
      TabIndex        =   18
      Text            =   "Combo3"
      Top             =   6240
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   8400
      TabIndex        =   17
      Text            =   "Combo2"
      Top             =   5520
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8400
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   400
      Left            =   8160
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   3450
      Width           =   2600
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   400
      Left            =   8160
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   2700
      Width           =   2600
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   400
      Left            =   8160
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   1970
      Width           =   2600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&new"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   7200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   19202050
      CurrentDate     =   38854
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   6570
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   19202051
      CurrentDate     =   38854
   End
   Begin VB.TextBox Text3 
      Height          =   1400
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "form2.frx":16986
      Top             =   5055
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   1400
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "form2.frx":1698C
      Top             =   3600
      Width           =   4095
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   3120
      TabIndex        =   3
      Top             =   1400
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   3120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1020
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   6200
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   5450
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8400
      TabIndex        =   14
      Top             =   4610
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   2600
      Left            =   7920
      Top             =   1520
      Width           =   3135
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
      Left            =   9705
      TabIndex        =   1
      Top             =   880
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
      Height          =   285
      Left            =   9720
      TabIndex        =   0
      Top             =   520
      Width           =   1575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
