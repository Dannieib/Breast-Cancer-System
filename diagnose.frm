VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form diagnose 
   Caption         =   "DIAGNOSE PATIENT(s)"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   14370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "LAST RECORD"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   37
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "FIRST RECORD"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   36
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "NEXT RECORD"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   35
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "PREVIOUS RECORD"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   34
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   33
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   11760
      TabIndex        =   32
      Top             =   7080
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      TabIndex        =   31
      Top             =   7080
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DIAGNOSE"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   30
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   29
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   360
      TabIndex        =   28
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "DIAGNOSE PATIENT"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   7320
      TabIndex        =   21
      Top             =   120
      Width           =   6855
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   23
         Text            =   "PLEASE SELECT..."
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label PRESCRIPTION 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   6615
      End
      Begin VB.Label Label13 
         Caption         =   "PRESCRIPTION"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   26
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label RESULT 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   6615
      End
      Begin VB.Label Label11 
         Caption         =   "RESULT"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   24
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "SYMPTHOMS"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "REGISTER PATIENT"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CheckBox Check2 
         Caption         =   "SINGLE"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "MARRIED"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   5280
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   18
         Top             =   4680
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   17
         Top             =   4080
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   16
         Top             =   3480
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   15
         Top             =   2880
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   2280
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   32899073
         CurrentDate     =   42947
      End
      Begin VB.OptionButton Option2 
         Caption         =   "FEMALE"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MALE"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   11
         Top             =   1080
         Width           =   3975
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
         Height          =   420
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label9 
         Caption         =   "MARITAL STATUS"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "STATE OF ORIGIN"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "NATIONALITY"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "DATE OF BIRTH"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "LAST NAME"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "FIRST NAME"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "diagnose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Dim REC As New ADODB.Recordset

Private Sub Command1_Click()
Dim EXITT As String
EXITT = MsgBox("SURE TO EXIT", vbYesNo, "EXIT")
If EXITT = vbYes Then
Unload Me
Else: Me.Refresh
End If
End Sub

Private Sub Command10_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or RESULT.Caption = "" Or PRESCRIPTION.Caption = "" Then
MsgBox "EMPTY FEILD(s) DISCOVERED", , "DIAGNOSE"
Else
 REC.MoveFirst
display
End If

End Sub

Private Sub Command11_Click()

If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or RESULT.Caption = "" Or PRESCRIPTION.Caption = "" Then
MsgBox "EMPTY FEILD(s) DISCOVERED", , "DIAGNOSE"
Else
REC.MoveLast
display
End If

End Sub

Private Sub Command12_Click()
REC.AddNew
End Sub

Sub Clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
DTPicker1.Value = "12/31/1999"
Combo1.Text = "PLEASE SELECT..."
PRESCRIPTION.Caption = ""
RESULT.Caption = ""
MsgBox "ALL FIELDS CLEARED", , "CANCER DIAGNOSE"
End Sub

Private Sub Command2_Click()
REC.AddNew
Clear
End Sub

Private Sub Command3_Click()
If Combo1.Text = "PLEASE SELECT..." Then
MsgBox " PLEASE CHOOSE SYMPTHOMS", , "PATIENT(s)"
ElseIf Combo1.Text = " REDNESS OF THE BREAST?" Then
RESULT.Caption = "COULD BE REACTION OF GENETICS OR MUCH ALCOHOLIC CONSUMPTION"
PRESCRIPTION.Caption = "TAKE 2 TABLETS OF HORMONE THERAPY AND IF REDNESS PERSIST, CONSULT YOUR PHYSICIAN"
ElseIf Combo1.Text = "PAINS ON THE BREAST OR ROUND THE ARMPIT?" Then
RESULT.Caption = "COULD BE BODY WEIGHT OR YOU COULD BE AGING"
PRESCRIPTION.Caption = "TAKE 2 TABLETS OF PANADOL AND BIOLOGICAL THERAPY TO EASE PAINS"
ElseIf Combo1.Text = "DISCHARGE OF UNNECCESSARY FLUID?" Then
RESULT.Caption = "COULD BE MUCH ALCOHOLIC INTAKE OR RADIATION TO EXPOSURE"
PRESCRIPTION.Caption = "TAKE RADIATION THERAPY OR YOU UNDERGO A SURGERY"
ElseIf Combo1.Text = "RASH ON QTHE NIPPLE?" Then
RESULT.Caption = "COULD BE INFECTION OR DUCTAL CARCINOMA CANCER"
PRESCRIPTION.Caption = " UNDERGO A SURGERY OR YOU TAKE CHEMOTHERAPY THOUGH MAY BE HARMFUL TO SKIN"
ElseIf Combo1.Text = "CHANGE IN THE SIZE OF THE NIPPLE?" Then
RESULT.Caption = "CAUSED BY DENSED BREAST TISSUE"
PRESCRIPTION.Caption = "UNDERGO A SURGERY IMMEDIATELY"
ElseIf Combo1.Text = "PEELING OF SKIN ON THE BREAST SKIN?" Then
RESULT.Caption = "COULD BE MUCH ALCOHOLIC INTAKE AND HORMONE TREATMENT"
PRESCRIPTION.Caption = "TAKE HORMONE AND GENETIC THERAPY"
End If
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or RESULT.Caption = "" Or PRESCRIPTION.Caption = "" Then
MsgBox "EMPTY FEILD(s) DISCOVERED", , "DIAGNOSE"
Else
REC.Fields("FIRST_NAME").Value = Text1.Text
REC.Fields("LAST_NAME").Value = Text2.Text
If Option1.Value = True Then
REC.Fields("GENDER") = Option1.Caption
Else
REC.Fields("GENDER") = Option2.Caption
End If
REC.Fields("DOB").Value = DTPicker1.Value
REC.Fields("PHONE").Value = Text3.Text
REC.Fields("ADDRESS").Value = Text4.Text
REC.Fields("NATIONALITY").Value = Text5.Text
REC.Fields("STATE").Value = Text6.Text
If Check1.Value = True Then
REC.Fields("MARITAL") = Check1.Caption
Else
REC.Fields("MARITAL") = Check2.Caption
End If
REC.Fields("SYMPTHOMS").Value = Combo1.Text
REC.Fields("RESULT").Value = RESULT.Caption
REC.Fields("PRESCRIPTION").Value = PRESCRIPTION.Caption
MsgBox "SUCCESSFULLY SAVED TO DATABASE", , "DIAGNOSE"
REC.Update
REFRESHF
End If
End Sub


Sub display()

Text1.Text = REC!first_name
Text2.Text = REC!last_name
If REC!gender = "male" Then
Option1.Value = True
Else: Option2.Value = True
End If
DTPicker1.Value = REC!dob
Text3.Text = REC!phone
Text4.Text = REC!address
Text5.Text = REC!nationality
Text6.Text = REC!State
If REC!marital = "married" Then
Check1.Value = True
Else
End If
Combo1.Text = REC!sympthoms
RESULT.Caption = REC!RESULT
PRESCRIPTION.Caption = REC!PRESCRIPTION


End Sub
Sub REFRESHF()
REC.Close
REC.Open "SELECT * FROM DIAGNOSE", CON, adOpenDynamic, adLockPessimistic
If Not REC.EOF Then
REC.MoveNext
display
Else
MsgBox "NO RECORD FOUND", , "REFRESH"
End If
End Sub
Private Sub Command6_Click()
Dim QUIT As String
QUIT = MsgBox("SURE TO DELETE THIS RECORD?", vbYesNo, "DELETE CONFIRMATION")
If vbYes Then
REC.Delete adAffectCurrent
MsgBox "DELETED SUCCESSFULLY", , "RECORD DELETED"
REC.Update
REFRESHF
Else
MsgBox "RECORD NOT WIPED!", , "DELETION"
End If
End Sub

Private Sub Command7_Click()
Me.PrintForm
End Sub


Private Sub Command8_Click()

If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or RESULT.Caption = "" Or PRESCRIPTION.Caption = "" Then
MsgBox "EMPTY FEILD(s) DISCOVERED", , "DIAGNOSE"
Else

REC.MovePrevious
If REC.BOF Then
REC.MoveLast
display
Else
display
End If
End If
End Sub

Private Sub Command9_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or RESULT.Caption = "" Or PRESCRIPTION.Caption = "" Then
MsgBox "EMPTY FEILD(s) DISCOVERED", , "DIAGNOSE"
Else

REC.MoveNext
If REC.EOF Then
REC.MoveFirst
display
Else
display
End If
End If
End Sub

Private Sub Form_Load()
CON.Open "provider=microsoft.jet.oledb.4.0;Data Source=E:\BREAST CANCER DIAGNOSE\BREAST_CANCER.mdb;PERSIST SECURITY INFO = FALSE"
REC.Open "SELECT * FROM DIAGNOSE", CON, adOpenDynamic, adLockPessimistic


Combo1.AddItem " REDNESS OF THE BREAST?"
Combo1.AddItem "PAINS ON THE BREAST OR ROUND THE ARMPIT?"
Combo1.AddItem "DISCHARGE OF UNNECCESSARY FLUID?"
Combo1.AddItem "RASH ON QTHE NIPPLE?"
Combo1.AddItem "CHANGE IN THE SIZE OF THE NIPPLE?"
Combo1.AddItem "PEELING OF SKIN ON THE BREAST SKIN?"
display
End Sub

