VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "�A���P�[�g�񓚉��"
   ClientHeight    =   6839
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5516
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
count = count + 1
If count = 1 Then
  answer = answer + 1
End If
Sheet2.Cells(answer + 1, 1) = answer
Sheet2.Cells(answer + 1, 137) = age
Sheet2.Cells(answer + 1, 138) = gender
If OptionButton1.Value = True Then
  Sheet2.Cells(answer + 1, m) = 1
ElseIf OptionButton2.Value = True Then
  Sheet2.Cells(answer + 1, m) = 2
ElseIf OptionButton3.Value = True Then
  Sheet2.Cells(answer + 1, m) = 3
ElseIf OptionButton4.Value = True Then
  Sheet2.Cells(answer + 1, m) = 4
ElseIf OptionButton5.Value = True Then
  Sheet2.Cells(answer + 1, m) = 5
End If
Unload Me
UserForm1.Show
End Sub

Private Sub UserForm_Initialize()
TextBox1.WordWrap = True '�܂�Ԃ�������
TextBox1.MultiLine = True '�����s������
TextBox1.ScrollBars = fmScrollBarsVertical
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
  UserForm7.Show
End If
End Sub

