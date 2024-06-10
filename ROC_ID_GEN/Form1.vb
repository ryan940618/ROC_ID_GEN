Public Class Form1
    Dim meet As Integer = 0
    Dim mcheck As Integer = 0
    Dim count As Integer
    Dim gencount As Integer
    Dim gencountcheck As Integer
    Private Sub Textbox8_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox8.TextChanged
        TextBox8.Select(TextBox8.TextLength, -1)
        TextBox8.ScrollToCaret()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        Form2.Show()
    End Sub

    Function Generate()
        BackgroundWorker1.RunWorkerAsync()
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        gencount = 1
        gencountcheck = 0
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Check_Length() = True Then
            If Check_City() = True Then
                If Check_Sex() = True Then
                    If Check_Serial() = True Then
                        If Check_Check() = True Then
                            TextBox3.Text = "正確"
                        Else
                            TextBox3.Text = "檢查碼錯誤"
                        End If
                    Else
                        TextBox3.Text = "流水號錯誤"
                    End If
                Else
                    TextBox3.Text = "性別錯誤"
                End If
            Else
                TextBox3.Text = "縣市錯誤"
            End If
        Else
            TextBox3.Text = "長度錯誤"
        End If
        TextBox8.Text = TextBox8.Text & "[" & String.Format("{0:tt:hh:mm:ss}", DateTime.Now) & "]Checked:" & TextBox4.Text & " Result:" & TextBox3.Text & vbNewLine
    End Sub

    Function Check_City()
        Dim i As Integer
        For i = 0 To 25
            If TextBox4.Text.Substring(0, 1) = ComboBox1.Items(i).ToString.Substring(1, 1) Then
                TextBox5.Text = ComboBox1.Items(i)
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function

    Function Check_Length()
        If TextBox4.TextLength = 10 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function Check_Sex()
        Dim i As Integer
        For i = 0 To 1
            If TextBox4.Text.Substring(1, 1) = (i + 1) Then
                TextBox6.Text = ComboBox2.Items(i)
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function

    Function Check_Serial()
        If IsNumeric(TextBox4.Text.Substring(2, 7)) Then
            If TextBox4.Text.Substring(2, 7).Length = 7 Then
                TextBox7.Text = TextBox4.Text.Substring(2, 7)
                Return True
                Exit Function
            End If
        End If
        Return False
    End Function

    Function Check_Check()
        '計算縣市的加權值()
        Dim city As Integer
        Dim i As Integer = 0
        For i = 0 To 25
            If ComboBox1.Items(i).ToString.Substring(1, 1) = TextBox4.Text.Substring(0, 1) Then
                city = (i + 10) / 10 + ((i + 10) Mod 10) * 9
                Exit For
            End If
        Next
        '計算性別的加權值
        Dim sex As Integer
        i = 0
        For i = 1 To 2
            If i = TextBox4.Text.Substring(1, 1) Then
                sex = sex + i * 8
            End If
        Next
        '計算流水號的加權值
        Dim serial As Integer
        For i = 0 To 6
            serial = serial + (7 - i) * TextBox4.Text.Substring(i + 2, 1)
        Next
        '計算全部的加權值
        Dim all As Integer
        all = city + sex + serial
        all = all Mod 10
        all = 10 - all
        all = all Mod 10
        '算出檢查碼
        Dim check As Integer
        check = all
        Label9.Text = "加權碼: city=" & city & ",sex=" & sex & ",serial=" & serial & ",check=" & check
        If TextBox4.Text.Substring(9, 1) = all Then
            Return True
            Exit Function
        End If
        Return False
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox8.Text = ""
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        gencount = 10
        gencountcheck = 0
        BackgroundWorker1.RunWorkerAsync()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False
        Button11.Enabled = False
        TextBox4.Enabled = False
        Button12.Enabled = True
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        gencount = 100
        gencountcheck = 0
        BackgroundWorker1.RunWorkerAsync()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False
        Button11.Enabled = False
        TextBox4.Enabled = False
        Button12.Enabled = True
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Form2.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If ComboBox1.Enabled = True Then
            ComboBox1.SelectedText = "隨機"
            ComboBox1.Enabled = False
        Else
            ComboBox1.SelectedIndex = 0
            ComboBox1.Enabled = True
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Enabled = True Then
            ComboBox2.SelectedText = "隨機"
            ComboBox2.Enabled = False
        Else
            ComboBox2.SelectedIndex = 0
            ComboBox2.Enabled = False
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        gencount = 1000
        gencountcheck = 0
        BackgroundWorker1.RunWorkerAsync()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False
        Button11.Enabled = False
        TextBox4.Enabled = False
        Button12.Enabled = True
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        gencount = 10000
        gencountcheck = 0
        BackgroundWorker1.RunWorkerAsync()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False
        Button11.Enabled = False
        TextBox4.Enabled = False
        Button12.Enabled = True
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        gencount = 100000
        gencountcheck = 0
        BackgroundWorker1.RunWorkerAsync()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False
        Button11.Enabled = False
        TextBox4.Enabled = False
        Button12.Enabled = True
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        meet = 0
        If Check_Length() = True Then
            If Check_City() = True Then
                If Check_Sex() = True Then
                    If Check_Serial() = True Then
                        If Check_Check() = True Then
                            TextBox3.Text = "正確"
                            mcheck = 1
                        Else
                            TextBox3.Text = "檢查碼錯誤"
                        End If
                    Else
                        TextBox3.Text = "流水號錯誤"
                    End If
                Else
                    TextBox3.Text = "性別錯誤"
                End If
            Else
                TextBox3.Text = "縣市錯誤"
            End If
        Else
            TextBox3.Text = "長度錯誤"
        End If
        TextBox8.Text = TextBox8.Text & "[" & String.Format("{0:tt:hh:mm:ss}", DateTime.Now) & "]GenUntillMeet_Checked:" & TextBox4.Text & " Result:" & TextBox3.Text & vbNewLine
        If mcheck = 1 Then
            BackgroundWorker2.RunWorkerAsync()
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button8.Enabled = False
            Button9.Enabled = False
            Button10.Enabled = False
            Button11.Enabled = False
            TextBox4.Enabled = False
            Button12.Enabled = True
        End If
    End Sub

    Function MetCheck()
        If TextBox2.Text = TextBox4.Text Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        gencountcheck = gencountcheck + 1
        Form1.CheckForIllegalCrossThreadCalls = False
        Dim city_value As Integer
        Dim sex_value As Integer
        Dim random As New Random
        Dim i As Integer
        If CheckBox1.Checked = True Then
            city_value = random.Next(0, 26)
        Else
            For i = 0 To 26
                If ComboBox1.SelectedIndex = i Then
                    city_value = i
                    Exit For
                End If
            Next
        End If

        If CheckBox2.Checked = True Then
            sex_value = random.Next(0, 2)
        Else
            For i = 0 To 1
                If ComboBox2.SelectedIndex = i Then
                    sex_value = i
                    Exit For
                End If
            Next
        End If

        '計算縣市的加權值
        Dim city As Integer
        i = city_value
        city = (i + 10) / 10 + ((i + 10) Mod 10) * 9

        '計算性別的加權值
        Dim sex As Integer
        i = sex_value
        sex = sex + (i + 1) * 8

        '計算流水號的加權值
        Dim rng As New Random
        Dim number As Integer = rng.Next(1, 10000000)
        Dim digits As String = number.ToString("0000000")
        Dim serial As Integer
        TextBox1.Text = digits
        For i = 0 To 6
            serial = serial + (7 - i) * digits.Substring(i, 1)
        Next
        '計算全部的加權值
        Dim all As Integer
        all = city + sex + serial
        all = all Mod 10
        all = 10 - all
        all = all Mod 10
        '算出檢查碼
        Dim check As Integer
        check = all
        TextBox2.Text = ComboBox1.Items(city_value).ToString.Substring(1, 1) & (sex_value + 1) & TextBox1.Text & check
        count = count + 1
        TextBox8.Text = TextBox8.Text & "[" & String.Format("{0:tt:hh:mm:ss}", DateTime.Now) & "]Generated:" & TextBox2.Text & " Count:" & count & vbNewLine
        Label10.Text = "加權碼: city=" & city & ",sex=" & sex & ",serial=" & serial & ",check=" & check
        Threading.Thread.Sleep(ComboBox3.Items(ComboBox3.SelectedIndex))
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Form2.TextBox1.Text = Form2.TextBox1.Text & TextBox2.Text & vbNewLine
        Form2.TextBox2.Text = Form2.TextBox2.Text & TextBox2.Text & ":" & count & vbNewLine
        If Not gencountcheck = gencount Then
            BackgroundWorker1.RunWorkerAsync()
        Else
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            Button1.Enabled = True
            Button2.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
            Button8.Enabled = True
            Button9.Enabled = True
            Button10.Enabled = True
            Button11.Enabled = True
            TextBox4.Enabled = True
            Button12.Enabled = False
        End If
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Form1.CheckForIllegalCrossThreadCalls = False
        Dim city_value As Integer
        Dim sex_value As Integer
        Dim random As New Random
        Dim i As Integer
     
            For i = 0 To 26
            If TextBox4.Text.Substring(0, 1) = ComboBox1.Items(i).ToString.Substring(1, 1) Then
                city_value = i
                Exit For
            End If
            Next

            For i = 0 To 1
            If i + 1 = TextBox4.Text.Substring(1, 1) Then
                sex_value = i
                Exit For
            End If
            Next

        '計算縣市的加權值
        Dim city As Integer
        i = city_value
        city = (i + 10) / 10 + ((i + 10) Mod 10) * 9

        '計算性別的加權值
        Dim sex As Integer
        i = sex_value
        sex = sex + (i + 1) * 8

        '計算流水號的加權值
        Dim rng As New Random
        Dim number As Integer = rng.Next(1, 10000000)
        Dim digits As String = number.ToString("0000000")
        Dim serial As Integer
        TextBox1.Text = digits
        For i = 0 To 6
            serial = serial + (7 - i) * digits.Substring(i, 1)
        Next
        '計算全部的加權值
        Dim all As Integer
        all = city + sex + serial
        all = all Mod 10
        all = 10 - all
        all = all Mod 10
        '算出檢查碼
        Dim check As Integer
        check = all
        TextBox2.Text = ComboBox1.Items(city_value).ToString.Substring(1, 1) & (sex_value + 1) & TextBox1.Text & check
        count = count + 1
        TextBox8.Text = TextBox8.Text & "[" & String.Format("{0:tt:hh:mm:ss}", DateTime.Now) & "]GenUntillMeet_Generated:" & TextBox2.Text & " Count:" & count & vbNewLine
        Label10.Text = "加權碼: city=" & city & ",sex=" & sex & ",serial=" & serial & ",check=" & check
        Threading.Thread.Sleep(ComboBox3.Items(ComboBox3.SelectedIndex))
    End Sub


    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Form2.TextBox1.Text = Form2.TextBox1.Text & TextBox2.Text & vbNewLine
        Form2.TextBox2.Text = Form2.TextBox2.Text & TextBox2.Text & ":" & count & vbNewLine
        If TextBox2.Text = TextBox4.Text Then
            meet = 1
        End If
        If meet = 1 Then
            If TextBox2.Text = TextBox4.Text Then
                TextBox8.Text = TextBox8.Text & "[" & String.Format("{0:tt:hh:mm:ss}", DateTime.Now) & "]GenUntillMeet_Met:" & TextBox4.Text & vbNewLine
            Else
                TextBox8.Text = TextBox8.Text & "[" & String.Format("{0:tt:hh:mm:ss}", DateTime.Now) & "]GenUntillMeet_Stopped:" & TextBox2.Text & vbNewLine
            End If
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            Button1.Enabled = True
            Button2.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
            Button8.Enabled = True
            Button9.Enabled = True
            Button10.Enabled = True
            Button11.Enabled = True
            TextBox4.Enabled = True
            Button12.Enabled = False
        Else
            BackgroundWorker2.RunWorkerAsync()
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If BackgroundWorker1.IsBusy = True Then
            gencountcheck = gencount
        End If
        If BackgroundWorker2.IsBusy = True Then
            meet = 1
        End If
    End Sub
End Class
