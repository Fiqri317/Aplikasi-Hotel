Public Class Login
    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox2.PasswordChar = "*"c
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim username As String = TextBox1.Text
        Dim password As String = TextBox2.Text

        If username = "admin" And password = "123" Then
            MessageBox.Show("Login Sukses", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Dim formHotel As New Hotel
            formHotel.Show()

            Me.Hide()
        Else
            MessageBox.Show("Username atau password salah!", "Login Gagal", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
