VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "TestEditor Lite Edition"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3390
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CommandButton1_Click()
    Dim a, b, i, j, cnt, qnum As Integer
    Dim w, q, an, t As Worksheet
    
    Set w = Worksheets("単語帳")
    Set q = Worksheets("問題")
    Set an = Worksheets("解答")
    Set t = Worksheets("tmp")
    Application.ScreenUpdating = False
    q.Cells(4, 5) = TextBox1.Text
    q.Cells(4, 7) = TextBox2.Text
    q.Cells(4, 13) = TextBox1.Text
    q.Cells(4, 15) = TextBox2.Text
    q.Cells(4, 21) = TextBox1.Text
    q.Cells(4, 23) = TextBox2.Text
    an.Cells(4, 5) = TextBox1.Text
    an.Cells(4, 7) = TextBox2.Text
    an.Cells(4, 13) = TextBox1.Text
    an.Cells(4, 15) = TextBox2.Text
    an.Cells(4, 21) = TextBox1.Text
    an.Cells(4, 23) = TextBox2.Text
    q.Cells(4, 8) = "/" + TextBox3.Text + "点"
    q.Cells(4, 16) = "/" + TextBox3.Text + "点"
    q.Cells(4, 24) = "/" + TextBox3.Text + "点"
    an.Cells(4, 8) = "/" + TextBox3.Text + "点"
    an.Cells(4, 16) = "/" + TextBox3.Text + "点"
    an.Cells(4, 24) = "/" + TextBox3.Text + "点"
    qnum = TextBox3.Text
    q.Range("C7:F26") = ""
    q.Range("K7:N26") = ""
    q.Range("S7:V27") = ""
    an.Range("C7:F26") = ""
    an.Range("K7:N26") = ""
    an.Range("S7:V27") = ""
    t.Range("A:A").ClearContents
    cnt = 1
    Call random
    If (TextBox2.Text - TextBox1.Text) < (qnum - 1) Then
        MsgBox "問題数は範囲にある単語以下に設定してください", vbExclamation, "注意"
        UserForm1.Hide
        UserForm1.Show
    End If
    
    If OptionButton1.Value = True Then
        '問題用紙と解答用紙作成
        For i = 1 To qnum
            For a = 2 To 201 '単語数増えたらtoの後ろの数増やす
                If t.Cells(i, 1) = w.Cells(a, 3) Then
                    If cnt < 21 Then
                        q.Cells(cnt + 6, 3) = w.Cells(a, 1)
                        an.Cells(cnt + 6, 3) = w.Cells(a, 1)
                        an.Cells(cnt + 6, 6) = w.Cells(a, 2)
                        cnt = cnt + 1
                    ElseIf 21 <= cnt And cnt <= 40 Then
                        q.Cells(cnt - 14, 11) = w.Cells(a, 1)
                        an.Cells(cnt - 14, 11) = w.Cells(a, 1)
                        an.Cells(cnt - 14, 14) = w.Cells(a, 2)
                        cnt = cnt + 1
                    ElseIf 41 <= cnt And cnt <= 60 Then
                        q.Cells(cnt - 34, 19) = w.Cells(a, 1)
                        an.Cells(cnt - 34, 19) = w.Cells(a, 1)
                        an.Cells(cnt - 34, 22) = w.Cells(a, 2)
                        cnt = cnt + 1
                    End If
                End If
            Next a
        Next i
    ElseIf OptionButton2.Value = True Then
        For i = 1 To qnum
            For b = 2 To 201 '単語数増えたらToの後ろの数増やす
                If t.Cells(i, 1) = w.Cells(b, 3) Then
                    If cnt < 21 Then
                        q.Cells(cnt + 6, 3) = w.Cells(b, 2)
                        an.Cells(cnt + 6, 3) = w.Cells(b, 2)
                        an.Cells(cnt + 6, 6) = w.Cells(b, 1)
                        cnt = cnt + 1
                    ElseIf 21 <= cnt And cnt < 40 Then
                         q.Cells(cnt - 14, 11) = w.Cells(b, 2)
                        an.Cells(cnt - 14, 11) = w.Cells(b, 2)
                        an.Cells(cnt - 14, 14) = w.Cells(b, 1)
                        cnt = cnt + 1
                    ElseIf 41 <= cnt And cnt < 60 Then
                        q.Cells(cnt - 34, 18) = w.Cells(b, 2)
                        an.Cells(cnt - 34, 18) = w.Cells(b, 2)
                        an.Cells(cnt - 34, 22) = w.Cells(b, 1)
                        cnt = cnt + 1
                    End If
                End If
            Next b
        Next i
    End If
    Application.ScreenUpdating = True
    Unload UserForm1
End Sub

Sub random()
    
    Dim i, j, Min, Max, qnum, tmp As Integer
    Min = TextBox1.Text
    Max = TextBox2.Text
    qnum = TextBox3.Text
    ReDim numlist(Min To Max) As Boolean
    Dim t As Worksheet
    Set t = Worksheets("tmp")
    
    
        If Min = 1 Then
            Randomize
                
                
            For i = Min To Max
                Do
                    tmp = Int((Max - Min + 1) * Rnd + Min)  '乱数生成の暗号
                Loop Until numlist(tmp) = False
                
                numlist(tmp) = True
                For j = 1 To qnum
                    If t.Cells(j, 1) = "" Then
                        t.Cells(j, 1) = tmp
                        GoTo nxt1
                    End If
                Next j
nxt1:
            Next i
        
        Else
            Randomize
                
                
            For i = Min To Max - 1
                Do
                    tmp = Int((Max - Min + 1) * Rnd + Min)  '乱数生成の暗号
                Loop Until numlist(tmp) = False
                
                numlist(tmp) = True
                For j = 1 To qnum
                    If t.Cells(j, 1) = "" Then
                        t.Cells(j, 1) = tmp
                        GoTo nxt2
                    End If
                Next j
nxt2:
            Next i
        End If
    
End Sub
