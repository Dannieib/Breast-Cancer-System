VERSION 5.00
Begin VB.Form splash 
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   9495
      Begin VB.Timer Timer1 
         Interval        =   4000
         Left            =   360
         Top             =   4920
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "VASTSOFTS INC."
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   4
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "BREAST CANCER DIAGNOSING SYSTEM"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   3480
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "OF"
         BeginProperty Font 
            Name            =   "Goudy Stout"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "DESIGN AND IMPLEMENTATION "
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   8295
      End
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
login.Show
End Sub
