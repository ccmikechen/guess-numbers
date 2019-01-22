Public Class Form1
    Dim num, C, L(2), F, dF(9) As String
    Dim A(2), B(2) As Integer
    Dim TN As Boolean = False

    Sub check(ByVal n, ByVal r)
        TextBox2.Text &= n
        Dim tA, tB As Integer
        For i = 1 To 4
            If Mid(n, i, 1) = Mid(num, i, 1) Then : tA += 1
            ElseIf num.Contains(Mid(n, i, 1)) Then : tB += 1
            End If
        Next

        A(r) = tA : B(r) = tB
        TextBox2.Text &= vbTab
        If tA > 0 Then TextBox2.Text &= tA & "A"
        If tB > 0 Then TextBox2.Text &= tB & "B"
        If tA = 0 And tB = 0 Then TextBox2.Text &= "-" : F &= n
        If tA + tB = 4 Then C = n
        TextBox2.Text &= vbCrLf
        If tA = 4 Then TN = True
    End Sub

    Sub clear()
        num = "" : C = "" : F = "" : TN = False
        ReDim L(2), dF(9), A(2), B(2)
        TextBox2.Clear()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        clear()
        num = TextBox1.Text
        If Len(num) <> 4 Then MsgBox("請輸入四位不重複數字") : Exit Sub
        Dim n1, n2, n3, n4 As Integer
        n1 = Mid(num, 1, 1) : n2 = Mid(num, 2, 1)
        n3 = Mid(num, 3, 1) : n4 = Mid(num, 4, 1)
        If n1 = n2 Or n1 = n3 Or n1 = n4 Or n2 = n3 Or n2 = n4 Or n3 = n4 Then
            MsgBox("數字不可重複") : Exit Sub
        End If
        Do While Len(C) < 10
            Randomize()
            Dim d = Int(Rnd() * 10)
            If InStr(C, d) = 0 Then C &= d
        Loop
        L(1) = Mid(C, 1, 4) : L(2) = Mid(C, 5, 4) : L(0) = Mid(C, 9, 2)
        check(L(1), 1) : If TN Then Exit Sub
        check(L(2), 2) : If TN Then Exit Sub
        If A(1) + B(1) + A(2) + B(2) = 4 Then F &= L(0)

        Dim r As Integer = 3
        Do Until TN
            Dim N, CT(4) As String
            For i = 1 To 4 : CT(i) = C : Next
            ReDim Preserve L(r), A(r), B(r)
            Dim f3 As Integer = 1

            N = ""
            Dim f2 As Boolean
            For i = 1 To 4
                For Each j As String In CT(i)
                    f2 = True
                    If InStr(F, j) = 0 And InStr(N, j) = 0 Then
                        For k = 1 To r - 1

                            If InStr(L(k), j) > 0 Then
                                Dim tA As Integer = A(k)
                                Dim tB As Integer = B(k)

                                For Each k2 As String In L(k)
                                    If InStr(N, k2) > 0 Then
                                        If InStr(N, k2) = InStr(L(k), k2) Then tA -= 1 Else tB -= 1
                                    End If
                                Next
                                If InStr(L(k), j) = i And tA = 0 Then f2 = False : Exit For
                                If InStr(L(k), j) <> i And tB = 0 Then f2 = False : Exit For

                            End If
                        Next
                        If f2 Then N &= j : Exit For
                    End If
                    f2 = False
                Next j
                If f2 = False Then
                    CT(i) = C
                    i -= 2
                    CT(i + 1) = CT(i + 1).Replace(Mid(N, i + 1, 1), "")
                    N = Mid(N, 1, i)
                End If

            Next i

            check(N, r)

            L(r) = N : r += 1
        Loop
    End Sub
End Class
