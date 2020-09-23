VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Student Number"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "11 numeric characters"
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      DataField       =   "Last Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "any character"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "First Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "(never leave it blank"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataField       =   "Middle Initial"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   ")"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text5 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "any character (never leave it blank)"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      DataField       =   "Telephone"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "numeric characters"
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   240
      Picture         =   "Base3.frx":0000
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "Base3.frx":030A
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   240
      Picture         =   "Base3.frx":0614
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   240
      Picture         =   "Base3.frx":091E
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label9 
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
      TabIndex        =   14
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      TabIndex        =   13
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1920
      TabIndex        =   12
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3720
      TabIndex        =   11
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M. I."
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5520
      TabIndex        =   10
      Top             =   1200
      Width           =   285
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      TabIndex        =   9
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      TabIndex        =   8
      Top             =   2280
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Often Mistakes"
      Height          =   210
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   1080
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Unload Me
End Sub

