VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPFStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pass/Fail Student Class Module
Private m_name As String
Private m_ssn As String
Private m_midterm As Single
Private m_final As Single

Property Get Name() As String
  Name = m_name
End Property

Property Let Name(ByVal vName As String)
  m_name = vName
End Property

Property Get SocSecNum() As String
  SocSecNum = m_ssn
End Property

Property Let SocSecNum(ByVal vNum As String)
  m_ssn = vNum
End Property

Property Let midGrade(ByVal vGrade As Single)
  m_midterm = vGrade
End Property

Property Let finGrade(ByVal vGrade As Single)
  m_final = vGrade
End Property

Public Function SemGrade() As String
  Dim grade As Single
  grade = (m_midterm + m_final) / 2
  grade = Round(grade)   'Round the grade
  If grade >= 60 Then
      SemGrade = "Pass"
    Else
      SemGrade = "Fail"
  End If
End Function

