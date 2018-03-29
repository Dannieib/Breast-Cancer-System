VERSION 5.00
Begin VB.Form menu 
   BackColor       =   &H00808000&
   Caption         =   "MAIN MENU"
   ClientHeight    =   6150
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Lucida Handwriting"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "menu.frx":0000
      Top             =   0
      Width           =   9300
   End
   Begin VB.Menu MNUPATIENT 
      Caption         =   "PATIENT"
      Begin VB.Menu FRMDIAGNOSE 
         Caption         =   "DIAGNOSE"
      End
   End
   Begin VB.Menu MNUEXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNUABOUT 
      Caption         =   "ABOUT"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FRMDIAGNOSE_Click()
Unload Me
diagnose.Show
End Sub

Private Sub MNUABOUT_Click()
MsgBox " BREAST CANCER DIANOSER..VERSION 1.0... POWERED BY:VASTSOFTS INC.", , "BREAST CANCER"
End Sub

Private Sub MNUEXIT_Click()
Dim EXITT As String
EXITT = MsgBox("SURE TO EXIT", vbYesNo, "EXIT")
If EXITT = vbYes Then
Unload Me
Else: Me.Refresh
End If
End Sub
