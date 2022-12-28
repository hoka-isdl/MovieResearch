VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "映画詳細ページ"
   ClientHeight    =   11816
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8309.001
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Unload Me
UserForm1.CommandButton5.Enabled = False
UserForm1.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me
UserForm6.Show
End Sub

Private Sub ToggleButton1_Click()

If v = 1 Then
  With UserForm1.Controls("ListBox2")
    .AddItem s
  End With
  v = 2
  UserForm4.ToggleButton1.Caption = "マイリストから削除"
ElseIf v = 2 Then
  With UserForm1.Controls("ListBox2")
    For j = 0 To .ListCount - 1
        If s = .List(j) Then
          .RemoveItem j
        Exit For
        End If
    Next j
  End With
  v = 1
  UserForm4.ToggleButton1.Caption = "マイリストへ追加"
End If
'ToggleButton1.Value = Not ToggleButton1.Value
End Sub

Private Sub UserForm_Initialize()
Dim i, k, p, x As Integer
Dim j As Integer
Dim ppoint As Integer: ppoint = 0
Dim npoint As Integer: npoint = 0
Dim a As Integer: a = 0
Dim b As Integer: b = 0
Dim c As Integer: c = 0
Dim d As Integer: d = 0
Dim men As Integer: men = 0
Dim women As Integer: women = 0
Dim a_par As Double: a_par = 0
Dim b_par As Double: b_par = 0
Dim c_par As Double: c_par = 0
Dim d_par As Double: d_par = 0
Dim men_par As Double: men_par = 0
Dim women_par As Double: women_par = 0
Dim TestStr As String

UserForm4.Label1.Caption = s
For m = 2 To movie + 1
  If Sheet1.Cells(m, 2) = s Then
    UserForm4.Label3.Caption = Sheet1.Cells(m, 6)
    UserForm4.Label5.Caption = Sheet1.Cells(m, 7)
    UserForm4.Label6.Caption = Sheet1.Cells(m, 8)
    UserForm4.Label11.Caption = Sheet1.Cells(m, 3) & "年"
    For j = 2 To answer + 1
      If Sheet2.Cells(j, m) = 1 Or Sheet2.Cells(j, m) = 2 Then
        If Sheet2.Cells(j, 137) = 1 Then
          a = a + 1
        ElseIf Sheet2.Cells(j, 137) = 2 Then
          b = b + 1
        ElseIf Sheet2.Cells(j, 137) = 3 Then
          c = c + 1
        ElseIf Sheet2.Cells(j, 137) = 4 Then
          d = d + 1
        End If
        If Sheet2.Cells(j, 138) = 1 Then
          men = men + 1
        ElseIf Sheet2.Cells(j, 138) = 2 Then
          women = women + 1
        End If
      End If
    Next
    
     ' テスト用の文字列を指定
    TestStr = Sheet1.Cells(m, 8)
    Exit For
  End If
Next
MeCabExecToSheet TestStr, Sheet10, 1
 p = 1
 x = 1
Do
  e = Sheet10.Cells(p, 2).Value
  If e = "" Then
    Exit Do
  End If
  If e = "名詞" Then
    Sheet11.Cells(x, 1).Value = Sheet10.Cells(p, 1).Value
    x = x + 1
  End If
  p = p + 1
Loop
p = 1
Sheet10.Cells.Clear
Do
  e = Sheet11.Cells(p, 1).Value
  If e = "" Then
    Exit Do
  End If
  For k = 1 To 13314
    If e = Sheet4.Cells(k, 1) Then
      If Sheet4.Cells(k, 2) = "p" Then
        ppoint = ppoint + 1
        Exit For
      ElseIf Sheet4.Cells(k, 2) = "n" Then
        npoint = npoint + 1
        Exit For
      End If
    End If
  Next
  p = p + 1
Loop
Sheet11.Cells.Clear
If ppoint > npoint Then
  UserForm4.Label16.Caption = "ポジティブ度"
  UserForm4.Label17.Caption = Round((ppoint / (ppoint + npoint)) * 100, 2) & "%"
Else
  UserForm4.Label16.Caption = "ネガティブ度"
  UserForm4.Label17.Caption = Round((npoint / (ppoint + npoint)) * 100, 2) & "%"
End If
a_par = (a / (a + b + c + d)) * 100
b_par = (b / (a + b + c + d)) * 100
c_par = (c / (a + b + c + d)) * 100
d_par = (d / (a + b + c + d)) * 100
men_par = (men / (men + women)) * 100
women_par = (women / (men + women)) * 100
UserForm4.Label14.Caption = "20代　" & Round(a_par, 2) & "%" & vbCrLf & "30代　" & Round(b_par, 2) & "%" & vbCrLf & "40代　" & Round(c_par, 2) & "%" & vbCrLf & "50代以上　" & Round(d_par, 2) & "%"
UserForm4.Label12.Caption = "男性　" & Round(men_par, 2) & "%" & vbCrLf & "女性　" & Round(women_par, 2) & "%"

If v = 2 Then
  'ToggleButton1.Value = True
  UserForm4.ToggleButton1.Caption = "マイリストから削除"
ElseIf v = 1 Then
  'ToggleButton1.Value = False
  UserForm4.ToggleButton1.Caption = "マイリストへ追加"
End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
  UserForm7.Show
End If
End Sub

