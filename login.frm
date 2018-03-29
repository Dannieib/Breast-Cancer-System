VERSION 5.00
Begin VB.Form login 
   Caption         =   "LOGIN"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Caption         =   "VASTSOFT INC."
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim EXITT As String
EXITT = MsgBox("SURE TO EXIT", vbYesNo, "EXIT")
If EXITT = vbYes Then
Unload Me
Else: Me.Refresh
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
If Text1.Text = "pius" And Text2.Text = "pius" Then
MsgBox "LOGIN SUCCESSFUL", , "LOGIN"
Unload Me
menu.Show
Else
MsgBox "INVALID PARAMETERS", , "LOGIN"
Me.Refresh
End If
End Sub
