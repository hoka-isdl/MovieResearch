VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "�ݒ�"
   ClientHeight    =   3794
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4305
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If OptionButton1.Value = True Then
  gender = 1
Else
  gender = 2
End If
If ComboBox1.Text = "20��" Then
  age = 1
ElseIf ComboBox1.Text = "30��" Then
  age = 2
ElseIf ComboBox1.Text = "40��" Then
  age = 3
ElseIf ComboBox1.Text = "50��ȏ�" Then
  age = 4
End If
Unload Me
UserForm1.Show
End Sub



Private Sub UserForm_Initialize()
With UserForm2.Controls("ComboBox1")
.AddItem "20��"
.AddItem "30��"
.AddItem "40��"
.AddItem "50��ȏ�"
End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
  UserForm7.Show
End If
End Sub

