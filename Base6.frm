VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
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
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   240
      Picture         =   "Base6.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   945
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Data1.Recordset.Index = "PrimaryKey"
    Form1.Data1.Recordset.Seek "=", Text1

If Form1.Data1.Recordset.NoMatch = False Then
    Form1.Enabled = True
    Form1.Command14.Visible = True
    Form1.Command1.Enabled = False
    Form1.Command2.Enabled = False
    Form1.Command3.Enabled = False
    Form1.Command4.Enabled = False
    Form1.Command5.Enabled = False
    Form1.Command6.SetFocus
    Unload Me
Else
    Form1.Text8 = "BlankImage.jpg"
    Text1 = ""
    Text1.SetFocus
End If

If Not Form1.Text8 = "BlankImage.jpg" Then
    Form1.Image1.Visible = True
    Form1.Image1.Picture = LoadPicture(Form1.Text8)
Else
    Form1.Image1.Visible = False
End If
    
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form1.Data1.Refresh
Form1.Command3.Enabled = True
Form1.Command4.Enabled = True

If Not Form1.Text8 = "BlankImage.jpg" Then
    Form1.Image1.Visible = True
    Form1.Image1.Picture = LoadPicture(Form1.Text8)
Else
    Form1.Image1.Visible = False
End If

Unload Me
End Sub

Private Sub Text1_Change()
If IsNumeric(Text1) = True And Len(Text1) = 11 Then
    Command1.Enabled = True
    Command1.SetFocus
Else
    Command1.Enabled = False
End If
End Sub
