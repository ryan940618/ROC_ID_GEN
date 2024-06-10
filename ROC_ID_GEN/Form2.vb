Public Class Form2
    Private Sub Textbox1_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox1.TextChanged
        TextBox1.Select(TextBox1.TextLength, -1)
        TextBox1.ScrollToCaret()
    End Sub
    Private Sub Textbox2_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.Select(TextBox2.TextLength, -1)
        TextBox2.ScrollToCaret()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Clipboard.SetText(TextBox1.Text)
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = ""
    End Sub
End Class