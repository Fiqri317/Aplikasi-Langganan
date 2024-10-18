Imports System
Imports System.Data
Imports System.Data.OleDb
Public Class Form1
    Dim _koneksiString As String
    Dim _koneksi As New OleDbConnection
    Dim komandambil As New OleDbCommand
    Dim datatabelku As New DataTable
    Dim dataadapterku As New OleDbDataAdapter
    Dim x As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _koneksiString = "Provider=Microsoft.Jet.OleDb.4.0;" + "Data Source=D:\Campus\Semester V\Kecerdasan Komputasi\Aplikasi Langganan Anime\database\Langganan.mdb;"
        _koneksi.ConnectionString = _koneksiString
        _koneksi.Open()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text

        komandambil.CommandText = "SELECT * FROM Pengguna"
        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)
        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        priceSubscription()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        priceSubscription()
    End Sub

    Private Sub priceSubscription()
        Dim duration As String = ComboBox1.Text
        Dim levelSubscription As String = If(RadioButton1.Checked, "Pro", "Premium")
        Dim price As String = ""

        Select Case duration
            Case "7 Days"
                price = If(levelSubscription = "Pro", "20000", "30000")
            Case "30 Days"
                price = If(levelSubscription = "Pro", "50000", "100000")
            Case "6 Months"
                price = If(levelSubscription = "Pro", "250000", "500000")
            Case "1 Years"
                price = If(levelSubscription = "Pro", "1000000", "1500000")
        End Select

        TextBox4.Text = price
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cmdTambah As New OleDbCommand
        Dim tanya As String
        Dim x As DataRow
        cmdTambah.Connection = _koneksi
        cmdTambah.CommandText = "INSERT INTO " + "Pengguna (Username, [Password], Email, Duration, [Subscription Level], Price, [Subscription Start Date], [Subscription End Date], Status, Payment)" +
            "VALUES ('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + ComboBox1.Text + "','" +
            If(RadioButton1.Checked, "Pro", "Premium") + "','" + TextBox4.Text + "','" + DateTimePicker1.Text + "','" +
            DateTimePicker2.Text + "','" + If(RadioButton3.Checked, "Active", "Expired") + "','" + ComboBox2.Text + " ')"
        tanya = MessageBox.Show("Data Ini di Simpan ?", "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If tanya = vbYes Then
            cmdTambah.ExecuteNonQuery()
            x = datatabelku.NewRow
            x("Username") = TextBox1.Text
            x("Password") = TextBox2.Text
            x("Email") = TextBox3.Text
            x("Duration") = ComboBox1.Text
            x("Subscription Level") = If(RadioButton1.Checked, "Pro", "Premium")
            x("Price") = TextBox4.Text
            x("Subscription Start Date") = DateTimePicker1.Text
            x("Subscription End Date") = DateTimePicker2.Text
            x("Status") = If(RadioButton3.Checked, "Active", "Expired")
            x("Payment") = ComboBox2.Text
            datatabelku.Rows.Add(x)
            Bs_Coba.DataSource = Nothing
            Bs_Coba.DataSource = datatabelku

            dgv_coba.Refresh()
            Bs_Coba.MoveFirst()
        End If
    End Sub

    Private Sub dgv_coba_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_coba.CellContentClick
        If dgv_coba.CurrentRow IsNot Nothing Then
            TextBox1.Text = dgv_coba.CurrentRow.Cells(0).Value.ToString()
            TextBox2.Text = dgv_coba.CurrentRow.Cells(1).Value.ToString()
            TextBox3.Text = dgv_coba.CurrentRow.Cells(2).Value.ToString()
            ComboBox1.Text = dgv_coba.CurrentRow.Cells(3).Value.ToString()

            Dim subscriptionLevel As String = dgv_coba.CurrentRow.Cells(4).Value.ToString()
            RadioButton1.Checked = (subscriptionLevel = "Pro")
            RadioButton2.Checked = (subscriptionLevel = "Premium")

            TextBox4.Text = dgv_coba.CurrentRow.Cells(5).Value.ToString()
            DateTimePicker1.Value = Convert.ToDateTime(dgv_coba.CurrentRow.Cells(6).Value)
            DateTimePicker2.Value = Convert.ToDateTime(dgv_coba.CurrentRow.Cells(7).Value)

            Dim status As String = dgv_coba.CurrentRow.Cells(8).Value.ToString()
            RadioButton3.Checked = (status = "Active")
            RadioButton4.Checked = (status = "Expired")

            ComboBox2.Text = dgv_coba.CurrentRow.Cells(9).Value.ToString()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim cmdHapus As New OleDbCommand
        cmdHapus.Connection = _koneksi
        cmdHapus.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Akan di Hapus ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        Dim username As String = TextBox1.Text.Trim().Replace("'", "''")
        cmdHapus.CommandText = "DELETE FROM Pengguna WHERE Username = '" & username & "'"
        cmdHapus.ExecuteNonQuery()


        Dim rowToDelete As DataRow = datatabelku.Select("Username='" & username & "'").FirstOrDefault()
        If rowToDelete IsNot Nothing Then
            datatabelku.Rows.Remove(rowToDelete)
        End If

        Bs_Coba.DataSource = Nothing
        Bs_Coba.DataSource = datatabelku
        dgv_coba.Refresh()


        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.SelectedIndex = -1
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        TextBox4.Clear()
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        ComboBox2.SelectedIndex = -1
        TextBox1.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        datatabelku.Clear()
        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text
        komandambil.CommandText = "SELECT * FROM " + "Pengguna WHERE Username LIKE '%" + TextBox5.Text + "%'"

        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)

        dgv_coba.Refresh()
        Bs_Coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_Coba
        Bs_Coba.MoveFirst()
    End Sub
End Class
