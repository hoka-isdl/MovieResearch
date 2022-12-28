VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "キーワード検索"
   ClientHeight    =   6272
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   8358.001
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim keyWord As String
Dim i As Integer
Dim k As Integer

Dim flag1 As Boolean
Dim flag2 As Boolean

With UserForm5.Controls("TextBox1")
keyWord = .Text
'Set myObj = myRange.Find(keyWord, LookAt:=xlPart)
End With
For i = 2 To movie + 1
        If InStr(Sheet1.Cells(i, 8).Value, keyWord) > 0 Then
           ListBox1.AddItem Sheet1.Cells(i, 2)
        End If
Next i
'同意語検索
For j = 1 To 232
  If InStr(Sheet6.Cells(j, 2).Value, keyWord) > 0 Then
    For i = 2 To movie + 1
        If InStr(Sheet1.Cells(i, 8).Value, Sheet6.Cells(j, 3)) > 0 Then
            For k = 0 To ListBox1.ListCount - 1
              If Sheet1.Cells(i, 2) = ListBox1.List(k) Then
                flag1 = True
              End If
            Next
            If flag1 = False Then
             ListBox1.AddItem Sheet1.Cells(i, 2)
            End If
           
        End If
    Next i
  End If
Next

'類義語検索
For j = 1 To 11767
  If InStr(Sheet7.Cells(j, 2).Value, keyWord) > 0 Then
    For i = 2 To movie + 1
        If InStr(Sheet1.Cells(i, 8).Value, Sheet7.Cells(j, 3)) > 0 Then
            For k = 0 To ListBox1.ListCount - 1
              If Sheet1.Cells(i, 2) = ListBox1.List(k) Then
                flag2 = True
              End If
            Next
            
            If flag2 = False Then
             ListBox1.AddItem Sheet1.Cells(i, 2)
            End If
           
        End If
    Next i
  End If
Next

If ListBox1.ListCount = 0 Then
ListBox1.AddItem "検索結果なし"
End If


End Sub

Private Sub CommandButton2_Click()
With UserForm5.Controls("ListBox1")
s = .Text
End With
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
Unload Me
UserForm4.Show



End Sub


Private Sub CommandButton3_Click()
Unload Me
UserForm1.Show
End Sub

Private Sub TextBox1_Change()
With UserForm5.Controls("ListBox1")
.Clear
End With
End Sub
Private Sub ListBox1_Change()
    If ListBox1.Text <> "" Then
        CommandButton2.Enabled = True
    Else
        CommandButton2.Enabled = False
    End If
End Sub



Private Sub UserForm_Initialize()
CommandButton2.Enabled = False
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
  UserForm7.Show
End If
End Sub

