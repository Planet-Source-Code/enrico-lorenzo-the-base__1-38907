VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Photo"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   1050
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1770
      Left            =   240
      Pattern         =   "*.bmp;*.gif;*.jpg"
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   330
      Left            =   2520
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drives"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   465
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Mid$(Dir1, Len(Dir1), 1) = "\" Then
    Form1.Text8 = Dir1 + File1
Else
    Form1.Text8 = Dir1 + "\" + File1
End If

Form1.Enabled = True
Form1.Command12.Enabled = True
Form1.Image1.Visible = True
Form4.Hide
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If Not File1 = "BlankImage.jpg" Then
    Command1.Enabled = True
End If
End Sub

