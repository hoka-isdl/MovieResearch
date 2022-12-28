VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ホーム画面"
   ClientHeight    =   6972
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6748
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ListBox1_Click()
    If ListBox1.Text <> "" Then
        CommandButton5.Enabled = True
    Else
        CommandButton5.Enabled = False
    End If
    For j = 0 To ListBox2.ListCount - 1
        ListBox2.Selected(j) = False
    Next j
   
    
End Sub
Private Sub CommandButton2_Click()

For j = 0 To ListBox1.ListCount - 1
    ListBox1.Selected(j) = False
Next j
Me.Hide
UserForm3.Show
End Sub

Private Sub CommandButton3_Click()
'Unload Me
For j = 0 To ListBox1.ListCount - 1
    ListBox1.Selected(j) = False
Next j
Me.Hide

UserForm5.Show
End Sub

Private Sub CommandButton4_Click()
Unload Me
UserForm2.Show
End Sub

Private Sub CommandButton5_Click()
If ListBox1.Text <> "" Then
    s = ListBox1.Text
    v = 1
    With UserForm1.Controls("ListBox2")
    For j = 0 To .ListCount - 1
        If s = .List(j) Then
          v = 2
          ListBox1.Selected(j) = False
        Exit For
        End If
        
    Next j
   
    
  End With
    
    'UserForm4.ToggleButton1.Value = False
ElseIf ListBox2.Text <> "" Then
    s = ListBox2.Text
    v = 2
    For j = 0 To ListBox1.ListCount - 1
      ListBox2.Selected(j) = False
    Next j

    'UserForm4.ToggleButton1.Value = True
End If

'Unload Me
For j = 0 To ListBox1.ListCount - 1
    ListBox1.Selected(j) = False
Next j
Me.Hide
UserForm4.Show

End Sub

Private Sub ListBox2_Change()
If ListBox2.Text <> "" Then
    CommandButton5.Enabled = True
Else
    CommandButton5.Enabled = False
End If
For j = 0 To ListBox1.ListCount - 1
    ListBox1.Selected(j) = False
Next j
End Sub



Private Sub UserForm_Initialize()
Dim i, j, k, l, eva As Integer
Dim e As String
count = 0
movie = 0
Do
  e = Sheet1.Cells(movie + 2, 1).Value
  If e = "" Then
    Exit Do
  End If
  movie = movie + 1
Loop
answer = 0
Do
  e = Sheet2.Cells(answer + 2, 1).Value
  If e = "" Then
    Exit Do
  End If
  answer = answer + 1
Loop
ReDim Data1(1 To movie) As Integer
ReDim Data2(1 To movie) As Integer
CommandButton5.Enabled = False
For i = 2 To movie + 1
 eva = 0
  For j = 2 To answer + 1
    If Sheet2.Cells(j, 137) = age And Sheet2.Cells(j, 138) = gender Then
      If Sheet2.Cells(j, i) = 1 Or Sheet2.Cells(j, i) = 2 Or Sheet2.Cells(j, i) = 3 Then
        eva = eva + 1
      End If
    End If
  Next
  Data1(i - 1) = eva
  Data2(i - 1) = i
Next

For k = 1 To movie - 1
  For l = k + 1 To movie
    If Data1(l) > Data1(k) Then
      tmp1 = Data1(l)
      tmp2 = Data2(l)
      Data1(l) = Data1(k)
      Data2(l) = Data2(k)
      Data1(k) = tmp1
      Data2(k) = tmp2
    End If
  Next
Next
With UserForm1.Controls("ListBox1")
For i = 1 To 7
.AddItem Sheet1.Cells(Data2(i), 2)
Next
  
End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
  UserForm7.Show
End If
End Sub
