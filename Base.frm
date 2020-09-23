VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THE BASE"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Base.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "&Search"
      Height          =   375
      Left            =   6000
      TabIndex        =   36
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      DataField       =   "Comment"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4800
      Width           =   3735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Delete P&hoto"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   33
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Browse Photo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   28
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      DataField       =   "Student Photo"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "datStudent"
      Connect         =   "Access"
      DatabaseName    =   "CS303P.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Students"
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sa&ve"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Add"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Last >>"
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "< &Previous"
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< &First"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Browse"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   5400
      Width           =   5415
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Birthdate"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "Telephone"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      DataField       =   "Middle Initial"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text3 
      DataField       =   "First Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "Last Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Student Number"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Cale&ndar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Photo"
      Height          =   1215
      Left            =   5880
      TabIndex        =   32
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Caption         =   "REFRESH"
      Height          =   375
      Left            =   6000
      TabIndex        =   39
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Maintenance"
      Height          =   3135
      Left            =   5880
      TabIndex        =   21
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Based from the student database format of STI College"
      Height          =   210
      Left            =   240
      TabIndex        =   38
      Top             =   6480
      Width           =   3960
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R. A. Baltazar and ENRICO X"
      Height          =   210
      Left            =   240
      TabIndex        =   37
      Top             =   6240
      Width           =   2070
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   240
      Picture         =   "Base.frx":030A
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
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
      TabIndex        =   35
      Top             =   4920
      Width           =   825
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PHOTO"
      Height          =   210
      Left            =   600
      TabIndex        =   26
      Top             =   1200
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   240
      Picture         =   "Base.frx":0614
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   240
      Picture         =   "Base.frx":091E
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "Base.frx":0C28
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "Base.frx":0F32
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Height          =   210
      Left            =   2520
      TabIndex        =   31
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      Height          =   210
      Left            =   1920
      TabIndex        =   30
      Top             =   1440
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   240
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate"
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
      TabIndex        =   15
      Top             =   4320
      Width           =   750
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
      TabIndex        =   12
      Top             =   3720
      Width           =   525
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
      TabIndex        =   10
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M. I."
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5280
      TabIndex        =   9
      Top             =   2640
      Width           =   285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3600
      TabIndex        =   8
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1920
      TabIndex        =   7
      Top             =   2640
      Width           =   765
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
      TabIndex        =   3
      Top             =   2400
      Width           =   465
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
      TabIndex        =   2
      Top             =   1800
      Width           =   945
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   240
      Picture         =   "Base.frx":123C
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim activate, record

Private Sub Command1_Click()
activate = 0
Data1.Recordset.MoveFirst
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command4.SetFocus
Image1.Picture = LoadPicture(Text8)

If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Visible = True
End If
End Sub

Private Sub Command10_Click()
Form1.Enabled = False
Form4.Show
End Sub

Private Sub Command11_Click()
Form1.Enabled = False
Form5.Calendar1.Value = Text7
Form5.Show
End Sub

Private Sub Command11_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    Text6.SetFocus
End If
End Sub

Private Sub Command12_Click()
Command12.Enabled = False
Text8 = "BlankImage.jpg"
Image1.Visible = False
End Sub

Private Sub Command13_Click()
Form1.Enabled = False
Form6.Show
End Sub

Private Sub Command14_Click()
Data1.Refresh
activate = 0

If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Visible = True
    Image1.Picture = LoadPicture(Text8)
End If

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = True

If Data1.Recordset.RecordCount = 0 Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command7.Enabled = False
    Command6.Enabled = False
    Command13.Enabled = False
End If

If Data1.Recordset.RecordCount = 1 Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End If

End Sub

Private Sub Command2_Click()
record = record + 1

Data1.Recordset.MovePrevious
Command3.Enabled = True
Command4.Enabled = True
Image1.Picture = LoadPicture(Text8)
If record = Data1.Recordset.RecordCount Then
    activate = 0
    Command1.Enabled = False
    Command2.Enabled = False
    Data1.Recordset.MoveFirst
    Image1.Picture = LoadPicture(Text8)
End If

If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Visible = True
End If
End Sub

Private Sub Command3_Click()
If activate = 0 Then
    record = Data1.Recordset.RecordCount
    activate = 1
End If

record = record - 1
Command1.Enabled = True
Command2.Enabled = True
Data1.Recordset.MoveNext
Image1.Picture = LoadPicture(Text8)
If record = 1 Then
    Command2.SetFocus
    Command3.Enabled = False
    Command4.Enabled = False
    Data1.Recordset.MoveLast
    Image1.Picture = LoadPicture(Text8)
End If

If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Visible = True
End If
End Sub

Private Sub Command4_Click()
record = 1
activate = 1
Data1.Recordset.MoveLast
Command1.Enabled = True
Command1.SetFocus
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Image1.Picture = LoadPicture(Text8)

If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Visible = True
End If

End Sub

Private Sub Command5_Click()
Data1.Recordset.AddNew
Image1.Picture = LoadPicture(Text8)
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command13.Enabled = False
Text1.Locked = False
Text1.SetFocus
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text9.Locked = False
End Sub

Private Sub Command6_Click()
Data1.Recordset.Edit
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command13.Enabled = False
Text1.Locked = False
Text1.SetFocus
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text9.Locked = False

If Text8 = "BlankImage.jpg" Then
    Command12.Enabled = False
Else
    Command12.Enabled = True
End If
End Sub

Private Sub Command7_Click()
activate = 0
Data1.Recordset.Delete
Data1.Refresh
Image1.Picture = LoadPicture(Text8)
Label11 = Data1.Recordset.RecordCount
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True

If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Visible = True
End If
    
If Data1.Recordset.RecordCount = 0 Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command13.Enabled = False
End If

If Data1.Recordset.RecordCount = 1 Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End If
End Sub

Private Sub Command8_Click()
If Len(Text6) = 0 Then
    Text6 = "NULL"
End If

If Len(Text9) = 0 Then
    Text9 = "NC"
End If

If Not IsNumeric(Text1) Or Not Len(Text1) = 11 Then
    Form1.Enabled = True
    Form2.Show
ElseIf Not IsNumeric(Text6) And Not Text6 = "NULL" Then
    Form1.Enabled = False
    Form2.Show
ElseIf Not Len(Text2) > 0 Or Not Len(Text3) > 0 Or Not Len(Text4) > 0 Or Not Len(Text5) > 0 Then
    Form1.Enabled = False
    Form2.Show
ElseIf Len(Text7) = 0 Then
    Form1.Enabled = False
    Form2.Show
Else
    If Len(Text8) = 0 Then
        Text8 = "BlankImage.jpg"
    End If
    
    Data1.Recordset.Update
    Data1.Refresh
    activate = 0
    Image1.Picture = LoadPicture(Text8)
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = False
    Command9.Enabled = False
    Command10.Enabled = False
    Command11.Enabled = False
    Command12.Enabled = False
    Command13.Enabled = True
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Text4.Locked = True
    Text5.Locked = True
    Text6.Locked = True
    Text9.Locked = True
    Label11 = Data1.Recordset.RecordCount
    
    If Text8 = "BlankImage.jpg" Then
        Image1.Visible = False
    Else
        Image1.Visible = True
    End If
 
    If Data1.Recordset.RecordCount = 1 Then
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
    End If
    
End If
End Sub

Private Sub Command9_Click()
activate = 0
Data1.Recordset.CancelUpdate
Data1.Refresh
Image1.Picture = LoadPicture(Text8)

If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Visible = True
End If

If Data1.Recordset.RecordCount = 0 Then
    Command3.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command13.Enabled = False
ElseIf Data1.Recordset.RecordCount = 1 Then
    Command3.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = True
    Command7.Enabled = True
Else
    Command3.Enabled = True
    Command4.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command13.Enabled = True
End If

Command5.Enabled = True
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text9.Locked = True
End Sub

Private Sub Form_Activate()
If Text8 = "BlankImage.jpg" Then
    Image1.Visible = False
Else
    Image1.Picture = LoadPicture(Text8)
End If

activate = 0
Command1.Enabled = False
Command2.Enabled = False
Label11 = Data1.Recordset.RecordCount

If Data1.Recordset.RecordCount = 0 Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command7.Enabled = False
    Command6.Enabled = False
    Command13.Enabled = False
End If

If Data1.Recordset.RecordCount = 1 Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form4
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'Enter Key
    Text2.SetFocus
End If
End Sub




Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
ElseIf KeyAscii = 8 And Len(Text2) = 0 Then 'Backspace Key
    Text1.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
ElseIf KeyAscii = 8 And Len(Text3) = 0 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
ElseIf KeyAscii = 8 And Len(Text4) = 0 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6.SetFocus
ElseIf KeyAscii = 8 And Len(Text5) = 0 Then
    Text4.SetFocus
End If
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Command11.Enabled = True Then
    Command11.SetFocus
ElseIf KeyAscii = 8 And Len(Text6) = 0 Then
    Text5.SetFocus
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 And Len(Text9) = 0 Then
    Command11.SetFocus
End If
End Sub
