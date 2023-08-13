Imports System.Data.Odbc
Public Class FormDataPelanggan
    Private Sub tampil_data()
        da = New OdbcDataAdapter("select Kode_Pelanggan, Nama_Pelanggan, Alamat_Pelanggan, No_Telp from tb_pelanggan", con)
        ds = New DataSet
        da.Fill(ds, "tb_pelanggan")
        DataGridView1.DataSource = ds.Tables("tb_pelanggan")
    End Sub
    Private Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub
    Sub delete()
        Dim a As String = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        If a = "" Then
            MsgBox("Data User yang dihapus belum dipilih")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan Kode_Pelanggan = " & a & " ? ", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK) Then
                Koneksi()
                cmd = New OdbcCommand("delete from tb_pelanggan where Kode_Pelanggan='" & a & "'", con)
                cmd.ExecuteNonQuery()
                MsgBox("Menghapus data user BERHASIL", vbInformation, "INFORMASI")
                con.Close()
                tampil_data()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call Koneksi()
        Dim simpan As String = "insert into tb_pelanggan values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
        cmd = New OdbcCommand(simpan, con)
        cmd.ExecuteNonQuery()
        MsgBox("Data berhasil disimpan")
        Call tampil_data()
        Call bersih()
    End Sub

    Private Sub FormDataPelanggan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call tampil_data()
        Call bersih()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim data As Integer
        data = DataGridView1.CurrentRow.Index

        TextBox1.Text = DataGridView1.Item(0, data).Value
        TextBox2.Text = DataGridView1.Item(1, data).Value
        TextBox3.Text = DataGridView1.Item(2, data).Value
        TextBox4.Text = DataGridView1.Item(3, data).Value
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Call Koneksi()
        Dim edit As String = "update tb_pelanggan set Nama_Pelanggan= '" & TextBox2.Text & "',Alamat_Pelanggan= '" & TextBox3.Text & "',No_Telp= '" & TextBox4.Text & "' where Kode_Pelanggan='" & TextBox1.Text & "'"
        cmd = New OdbcCommand(edit, con)
        cmd.ExecuteNonQuery()
        MsgBox("Data berhasil diubah", vbInformation, "Pemberitahuan")
        Call tampil_data()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Call delete()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged
        Call Koneksi()
        cmd = New OdbcCommand("select * from tb_pelanggan where Kode_Pelanggan Like '%" & TextBox5.Text & "%'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Call Koneksi()
            da = New OdbcDataAdapter("select * from tb_pelanggan where Kode_Pelanggan Like '%" & TextBox5.Text & "%'", con)
            ds = New DataSet
            da.Fill(ds, "tb_pelanggan")
            DataGridView1.DataSource = ds.Tables("tb_pelanggan")
            DataGridView1.ReadOnly = True
        End If
    End Sub

    Private Sub Label5_Click(sender As System.Object, e As System.EventArgs) Handles Label5.Click

    End Sub
End Class