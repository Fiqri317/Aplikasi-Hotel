Imports System
Imports System.Data
Imports System.Data.OleDb
Public Class Hotel
    Dim _koneksiString As String
    Dim _koneksi As New OleDbConnection
    Dim komandambil As New OleDbCommand
    Dim datatabelku As New DataTable
    Dim dataadapterku As New OleDbDataAdapter
    Dim x As String

    Private Sub dgv_coba_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles dgv_coba.CellFormatting
        If dgv_coba.Columns(e.ColumnIndex).Name = "Tanggal Checkin" Then
            If e.Value IsNot Nothing Then
                Dim tgl As DateTime = Convert.ToDateTime(e.Value)
                e.Value = tgl.ToString("dd MMMM yyyy", New System.Globalization.CultureInfo("id-ID"))
                e.FormattingApplied = True
            End If
        End If

        If dgv_coba.Columns(e.ColumnIndex).Name = "Tanggal Checkout" Then
            If e.Value IsNot Nothing Then
                Dim tgl As DateTime = Convert.ToDateTime(e.Value)
                e.Value = tgl.ToString("dd MMMM yyyy", New System.Globalization.CultureInfo("id-ID"))
                e.FormattingApplied = True
            End If
        End If
    End Sub

    Private Sub Hotel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _koneksiString = "Provider=Microsoft.Jet.OleDb.4.0;" + "Data Source=D:\Campus\Semester V\Kecerdasan Komputasi\Aplikasi Hotel\database\Hotel.mdb;"
        _koneksi.ConnectionString = _koneksiString
        _koneksi.Open()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text

        komandambil.CommandText = "SELECT * FROM Hotel"
        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)
        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        HitungTotalHari()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        HitungTotalHari()
    End Sub

    Private Sub HitungTotalHari()
        Dim tglCheckIn As DateTime = DateTimePicker1.Value
        Dim tglCheckOut As DateTime = DateTimePicker2.Value
        Dim totalHari As Integer = (tglCheckOut - tglCheckIn).Days

        If totalHari >= 0 Then
            TextBox3.Text = totalHari.ToString() + " Hari"
        Else
            MessageBox.Show("Tanggal Checkout tidak boleh sebelum Tanggal Checkin!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox3.Clear()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cmdTambah As New OleDbCommand
        Dim tanya As String
        Dim x As DataRow
        cmdTambah.Connection = _koneksi
        cmdTambah.CommandText = "INSERT INTO " + "Hotel ([No Identitas], Nama, [Tanggal Checkin], [Tanggal Checkout], Total)" +
            "VALUES ('" + TextBox1.Text + "','" + TextBox2.Text + "','" + DateTimePicker1.Text + "','" +
            DateTimePicker2.Text + "','" + TextBox3.Text + " ')"
        tanya = MessageBox.Show("Data Ini di Simpan ?", "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If tanya = vbYes Then
            cmdTambah.ExecuteNonQuery()
            x = datatabelku.NewRow
            x("No Identitas") = TextBox1.Text
            x("Nama") = TextBox2.Text
            x("Tanggal Checkin") = DateTimePicker1.Text
            x("Tanggal Checkout") = DateTimePicker2.Text
            x("Total") = TextBox3.Text
            datatabelku.Rows.Add(x)
            Bs_coba.DataSource = Nothing
            Bs_coba.DataSource = datatabelku

            dgv_coba.Refresh()
            Bs_coba.MoveFirst()
        End If
    End Sub

    Private Sub dgv_coba_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_coba.CellContentClick
        TextBox1.Text = dgv_coba.CurrentRow.Cells(0).Value.ToString()
        TextBox2.Text = dgv_coba.CurrentRow.Cells(1).Value.ToString()
        DateTimePicker1.Value = Convert.ToDateTime(dgv_coba.CurrentRow.Cells(2).Value)
        DateTimePicker2.Value = Convert.ToDateTime(dgv_coba.CurrentRow.Cells(3).Value)
        TextBox3.Text = dgv_coba.CurrentRow.Cells(4).Value.ToString()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim cmdHapus As New OleDbCommand
        cmdHapus.Connection = _koneksi
        cmdHapus.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Akan di Hapus ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdHapus.CommandText = "DELETE FROM " + "Hotel WHERE [No Identitas]=" + TextBox1.Text
            cmdHapus.ExecuteNonQuery()
        End If
        Bs_coba.RemoveCurrent()
        dgv_coba.Refresh()

        TextBox1.Clear()
        TextBox2.Clear()
        DateTimePicker1.Value = DateTime.Now
        DateTimePicker2.Value = DateTime.Now
        TextBox3.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmdUpdate As New OleDbCommand
        cmdUpdate.Connection = _koneksi
        cmdUpdate.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Ingin di Perbarui?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdUpdate.CommandText = "UPDATE Hotel SET " +
                "Nama = '" + TextBox2.Text + "', " +
                "[Tanggal Checkin] = '" + DateTimePicker1.Text + "', " +
                "[Tanggal Checkout] = '" + DateTimePicker2.Text + "', " +
                "Total = '" + TextBox3.Text + "' " +
                "WHERE [No Identitas] = " + TextBox1.Text  '
            cmdUpdate.ExecuteNonQuery()
            Dim rowToUpdate As DataRow = datatabelku.Select("[No Identitas] = " + TextBox1.Text).FirstOrDefault()
            If rowToUpdate IsNot Nothing Then
                rowToUpdate("No Identitas") = TextBox1.Text
                rowToUpdate("Nama") = TextBox2.Text
                rowToUpdate("Tanggal Checkin") = DateTimePicker1.Text
                rowToUpdate("Tanggal Checkout") = DateTimePicker2.Text
                rowToUpdate("Total") = TextBox3.Text
            End If
            Bs_coba.DataSource = Nothing
            Bs_coba.DataSource = datatabelku
            dgv_coba.Refresh()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox2.Clear()
        DateTimePicker1.Value = DateTime.Now
        DateTimePicker2.Value = DateTime.Now
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.Items.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        datatabelku.Clear()
        Dim kolomPencarian As String = ""
        If ComboBox1.SelectedItem Is Nothing Then
            MessageBox.Show("Pilih kolom pencarian terlebih dahulu!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Select Case ComboBox1.SelectedItem.ToString()
            Case "No Identitas"
                kolomPencarian = "[No Identitas]"
            Case "Nama"
                kolomPencarian = "Nama"
            Case "Tanggal Checkin"
                kolomPencarian = "[Tanggal Checkin]"
            Case "Tanggal Checkout"
                kolomPencarian = "[Tanggal Checkout]"
            Case "Total"
                kolomPencarian = "Total"
            Case Else
                MessageBox.Show("Kolom pencarian tidak valid!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
        End Select

        ' Pastikan ada inputan pencarian
        If TextBox4.Text = "" Then
            MessageBox.Show("Masukkan data pencarian!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text
        komandambil.CommandText = "SELECT * FROM Hotel WHERE " + kolomPencarian + " LIKE '%" + TextBox4.Text + "%'"
        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)
        dgv_coba.Refresh()
        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
        Bs_coba.MoveFirst()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
        Login.Show()
        Login.TextBox1.Clear()
        Login.TextBox2.Clear()
    End Sub
End Class