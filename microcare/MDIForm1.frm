VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Microcare Call Center Service Managemnt System"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mServices 
      Caption         =   "Service"
      Begin VB.Menu mServices_smServeceWizard 
         Caption         =   "Service Wizard"
      End
      Begin VB.Menu mServices_smWarrentyProductService 
         Caption         =   "Waranty Product/Service"
      End
      Begin VB.Menu mService_smExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu maccounts 
      Caption         =   "Accounts"
      Begin VB.Menu mAccounts_smProductServicePayment 
         Caption         =   "product/Service Payment"
      End
   End
   Begin VB.Menu mhrmanagement 
      Caption         =   "HR Management"
      Begin VB.Menu mHRManagement_smEmployees 
         Caption         =   "Employee"
         Begin VB.Menu mHRManagement_smEmployees_bsmEmployeeDetails 
            Caption         =   "Employee Details"
         End
         Begin VB.Menu mHRManagement_smEmployees_bsmDutyAllocation 
            Caption         =   "Duty Allocation"
         End
         Begin VB.Menu mHRManagement_smEmployees_bsmPaymentDetails 
            Caption         =   "Salary Details"
         End
         Begin VB.Menu sep 
            Caption         =   "-"
         End
         Begin VB.Menu mHRManagement_smEmployees_bsmAttendance 
            Caption         =   "Attandance"
         End
         Begin VB.Menu sep1 
            Caption         =   "-"
         End
         Begin VB.Menu mHRManagement_smEmployees_bsmPayroll 
            Caption         =   "Payroll"
         End
      End
   End
   Begin VB.Menu mEnquiry 
      Caption         =   "Enquiry"
      Begin VB.Menu mEnquiry_smJobStatus 
         Caption         =   "Job Status"
      End
      Begin VB.Menu mEnquiry_smPLInfo 
         Caption         =   "Profit/ Loss Information"
      End
      Begin VB.Menu mEnquiry_smprodavail 
         Caption         =   "Current Product Availablity"
      End
   End
   Begin VB.Menu mreports 
      Caption         =   "Reports"
      Begin VB.Menu mReport_smjobsToDo 
         Caption         =   "Jobs To Do"
      End
   End
   Begin VB.Menu mtools 
      Caption         =   "Tools"
      Begin VB.Menu mTools_smBackUp 
         Caption         =   "Backup"
      End
      Begin VB.Menu mutility 
         Caption         =   "Utilities"
         Begin VB.Menu bsmCalculator 
            Caption         =   "Calculator"
         End
         Begin VB.Menu bsmMediaPlayer 
            Caption         =   "Media Player"
         End
         Begin VB.Menu bsmNotepad 
            Caption         =   "Notpad"
         End
         Begin VB.Menu bsmWorkPad 
            Caption         =   "Wordpad"
         End
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "Help"
      Begin VB.Menu mHelp_smHelpTopics 
         Caption         =   "Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mHelp_smAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
If UCase(UserProfile) = UCase("Service Provider") Then
MDIForm1.mAccounts_smProductServicePayment.Enabled = False
MDIForm1.mEnquiry_smPLInfo.Enabled = False
MDIForm1.mHRManagement_smEmployees.Enabled = False
MDIForm1.mServices_smWarrentyProductService.Enabled = False
MDIForm1.mTools_smBackUp.Enabled = False
End If
End Sub

Private Sub mEnquiry_smPLInfo_Click()
Form12.Show
End Sub

Private Sub mEnquiry_smprodavail_Click()
Form13.Show
End Sub

Private Sub mReport_smjobsToDo_Click()
DataReport1.Show
End Sub

Private Sub bsmCalculator_Click()
Shell ("calc.exe")
End Sub

Private Sub bsmMediaPlayer_Click()
Shell ("C:\Program Files\Windows Media Player\mplayer2.exe")
End Sub

Private Sub bsmNotepad_Click()
Shell ("notepad.exe")
End Sub

Private Sub bsmWorkPad_Click()
Shell ("D:\Program Files\Windows NT\Accessories\wordpad.exe")
End Sub

Private Sub mAccounts_smProductServicePayment_Click()
Form10.Show
End Sub

Private Sub mEnquiry_smJobStatus_Click()
Form11.Show
End Sub

Private Sub mHelp_smAbout_Click()
frmabout.Show
End Sub

Private Sub mHelp_smHelpTopics_Click()
On Error GoTo xxxxx
Form_F1.Show
xxxxx:
Exit Sub
Open "C:\Project system\Call Centre Service Management System\Design\Help Files\APMS Help.hlp" For Random As file
End Sub

Private Sub mHRManagement_smCustomer_Click()
Form7.Show
End Sub

Private Sub mHRManagement_smEmployees_bsmAttendance_Click()
Form6.Show
End Sub

Private Sub mHRManagement_smEmployees_bsmDutyAllocation_Click()
Form4.Show
End Sub

Private Sub mHRManagement_smEmployees_bsmEmployeeDetails_Click()
Form1.Show
End Sub

Private Sub mHRManagement_smEmployees_bsmPaymentDetails_Click()
Form5.Show
End Sub

Private Sub mHRManagement_smEmployees_bsmPayroll_Click()
Form3.Show
End Sub

Private Sub mService_smExit_Click()
Dim eRpl As Integer
eRpl = MsgBox("Are you sure to exit?", vbYesNo + vbQuestion, "Message")
If eRpl = vbYes Then
    Unload Me
    End
End If
End Sub

Private Sub mServices_smServeceWizard_Click()
Form2.Show
End Sub

Private Sub mServices_smWarrentyProductService_Click()
Form9.Show
End Sub
 

