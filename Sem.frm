VERSION 5.00
Begin VB.Form frm13_2_2 
   Caption         =   "Semester Grades"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSemGrade 
      Caption         =   "Calculate Grades"
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   2160
      Width           =   1332
   End
   Begin VB.PictureBox picGrades 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   4395
      TabIndex        =   14
      Top             =   3960
      Width           =   4452
   End
   Begin VB.Frame fraType 
      Caption         =   "Type of Registration"
      Height          =   735
      Left            =   600
      TabIndex        =   11
      Top             =   1200
      Width           =   3495
      Begin VB.OptionButton optPF 
         Caption         =   "Pass/Fail"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optReg 
         Caption         =   "Regular"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   2160
      Width           =   1092
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display Single Grade"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtFinal 
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtMidterm 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtSSN 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Student"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label lblFinal 
      Caption         =   "Final"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblMidterm 
      Caption         =   "Midterm"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblSSN 
      Caption         =   "SSN"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frm13_2_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pupil As Object
Dim section As New Collection

Private Sub Form_Load()
  optReg = True
End Sub

Private Sub cmdAdd_Click()
  If optReg.Value Then
      Set pupil = New CSTudent
    Else
      Set pupil = New CPFStudent
  End If
  'Read the Values stored in the Text boxes
  pupil.Name = txtName
  pupil.SocSecNum = txtSSN
  pupil.midGrade = Val(txtMidterm)
  pupil.finGrade = Val(txtFinal)
  section.Add pupil, txtSSN.Text
  'Clear Text Boxes
  txtName.Text = ""
  txtSSN.Text = ""
  txtMidterm.Text = ""
  txtFinal.Text = ""
End Sub

Private Sub cmdSemGrade_Click()
  Dim i As Integer, grade As String
  picGrades.Cls
  For i = 1 To section.Count
    picGrades.Print section.Item(i).Name; Tab(28); section.Item(i).SocSecNum; _
                    Tab(48); section.Item(i).SemGrade
  Next i
End Sub

Private Sub cmdDisplay_Click()
  Dim ssn As String
  ssn = InputBox("Enter the student's social security number.")
  picGrades.Cls
  picGrades.Print section.Item(ssn).Name; Tab(28); section.Item(ssn).SocSecNum(); _
                  Tab(48); section.Item(ssn).SemGrade
End Sub

Private Sub cmdQuit_Click()
  End
End Sub
