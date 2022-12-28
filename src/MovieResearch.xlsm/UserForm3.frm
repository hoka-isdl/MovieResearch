VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "ジャンル検索"
   ClientHeight    =   5894
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4515
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()
Unload Me
UserForm1.Show
End Sub

Private Sub ListBox1_Change()
    If ListBox1.Text <> "" Then
        CommandButton2.Enabled = True
    Else
        CommandButton2.Enabled = False
    End If
End Sub
Private Sub ComboBox1_Change()
With UserForm3.Controls("ListBox1")
.Clear
End With
End Sub

Private Sub ComboBox2_Change()
With UserForm3.Controls("ListBox1")
.Clear
End With
End Sub

Private Sub CommandButton1_Click()
Dim i, j, t, k, l, o, n, h, tmp1, tmp2 As Integer: h = 1
ReDim Data1(1 To movie) As Integer
ReDim Data2(1 To movie) As Integer

With UserForm3.Controls("ListBox1")
   

   '人気度順
   If ComboBox2.ListIndex = 0 Then
    For i = 2 To movie + 1
        If Sheet1.Cells(i, 4) = ComboBox1.Text Or Sheet1.Cells(i, 5) = ComboBox1.Text Or ComboBox1.ListIndex = 0 Then
         t = 0
         For j = 2 To answer + 1
         tj = 1
          If tj = "" Then
            tj = 0
          End If
          t = tj + t
         Next
         Data1(h) = t
         Data2(h) = i
         h = h + 1
         .AddItem Sheet1.Cells(i, 2)
       
         
        End If
        
    Next
   
   End If
   '古い年度順
   If ComboBox2.ListIndex = 1 Or ComboBox2.ListIndex = 2 Then
     
     For i = 2 To movie + 1
       If Sheet1.Cells(i, 4) = ComboBox1.Text Or Sheet1.Cells(i, 5) = ComboBox1.Text Or ComboBox1.ListIndex = 0 Then
         Data1(h) = Sheet1.Cells(i, 3)
         Data2(h) = i
         h = h + 1
         .AddItem Sheet1.Cells(i, 2)
       End If
     Next
    End If
    
    If ComboBox2.ListIndex = 1 Then
     For k = 1 To .ListCount - 1
      For l = k + 1 To .ListCount
        If Data1(l) < Data1(k) Then
           tmp1 = Data1(l)
           tmp2 = Data2(l)
           Data1(l) = Data1(k)
           Data2(l) = Data2(k)
           Data1(k) = tmp1
           Data2(k) = tmp2
        End If
      Next
     Next
     n = .ListCount
    .Clear
     For o = 1 To n
      .AddItem Sheet1.Cells(Data2(o), 2)
    Next
   End If
   
    If ComboBox2.ListIndex = 0 Or ComboBox2.ListIndex = 2 Then
     For k = 1 To .ListCount - 1
      For l = k + 1 To .ListCount
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
     n = .ListCount
    .Clear
     For o = 1 To n
      .AddItem Sheet1.Cells(Data2(o), 2)
    Next
   End If

End With

End Sub


Private Sub CommandButton2_Click()

s = ListBox1.Text
 With UserForm1.Controls("ListBox2")
 v = 1
    For j = 0 To .ListCount - 1
        If s = .List(j) Then
          v = 2
        Exit For
        End If
        
    Next j
   
    
  End With
Unload Me
UserForm4.Show

End Sub

Private Sub UserForm_Initialize()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim flag1 As Boolean
Dim flag2 As Boolean
CommandButton2.Enabled = False
With UserForm3.Controls("ComboBox1")
.AddItem "総合"
.ListIndex = 0
  For i = 2 To movie + 1
      flag1 = False
      flag2 = False
    If .ListCount = 0 Then
      .AddItem Sheet1.Cells(i, 4)
      
    Else
      
      For j = 0 To .ListCount - 1
        If Sheet1.Cells(i, 4) = .List(j) Then
          flag1 = True
          
        End If
        If Sheet1.Cells(i, 5) = .List(j) Then
          flag2 = True
          
        End If
      Next j
      
      If flag1 = False Then
      .AddItem Sheet1.Cells(i, 4)
      End If
      If flag2 = False And Sheet1.Cells(i, 5) <> "" Then
      .AddItem Sheet1.Cells(i, 5)
      End If
    End If
  Next i
End With


With UserForm3.Controls("ComboBox2")
.AddItem "人気度順"
.AddItem "古い年度順"
.AddItem "新しい年度順"
.ListIndex = 0
End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
  UserForm7.Show
End If
End Sub

