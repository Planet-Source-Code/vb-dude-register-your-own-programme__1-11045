VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register Programme"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "SampleReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SampleReg.frx":0442
   ScaleHeight     =   2385
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Hidden 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register!"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Register Code: ED23498-28"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Name: Reynard Chan"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Register Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Oops! Your Trial Version is over. Please register this programme now."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4845
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "Reynard Chan" Or Text2.Text = "ED23498-28" Then
'Save the text inside the textboxes
SaveSetting "MyApp", "General", "User", Text1.Text
SaveSetting "MyApp", "General", "User2", Text2.Text
MsgBox "Thanks for Registering!", vbExclamation, "Register"
'Unload Form
Unload Me
Else
MsgBox "Arrr... Wrong. Try again.", vbInformation, "Register"
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
'Get rid of all text
Text2.Text = ""
Text1.Text = ""
SaveSetting "MyApp", "General", "User2", Text2.Text
SaveSetting "MyApp", "General", "User", Text1.Text
'Reset the Register
Hidden.Text = "d"
SaveSetting "MyApp", "General", "Reset", Hidden.Text
'Unload form
Unload Me
End Sub

Private Sub Command3_Click()
MsgBox "Sample on How to Make a Registration Form For Your Programme" & _
vbCrLf & "By Reynard Chan, Age 12" & _
vbCrLf & "Made with Visual Basic 6" & _
vbCrLf & "Located at: www.planet-source-code.com/vb/" & _
vbCrLf & "Please send comments at that site or: vbdude@dbzfreak.cjb.net", vbInformation, "Register Sample"
End Sub

Private Sub Form_Load()
'Check if reseted
Hidden.Text = GetSetting("MyApp", "General", "Reset", Hidden.Text)
'Check if registered or not.
Text2.Text = GetSetting("MyApp", "General", "User2", Text2.Text)
Text1.Text = GetSetting("MyApp", "General", "User", Text1.Text)
If Text1.Text = "Reynard Chan" Or Text2.Text = "ED23498-28" Then
MsgBox "You've registered the programme! Woopee! ", vbInformation, "Registered!"
Else
MsgBox "Please register this programme now", vbExclamation, "Unregistered"
End If
End Sub
