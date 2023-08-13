Imports System.Data.Odbc
Public Class FormDataPewangi
    Private Sub tampil_data()
        da = New OdbcDataAdapter("select Nama_Pewangi from tb_pewangi", con)
        ds = New DataSet
        da.Fill(ds, "tb_pewangi")
        DataGridView1.DataSource = ds.Tables("tb_pewangi")
    End Sub
    Private Sub bersih()
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
    Sub delete()
        Dim a As String = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        If a = "" Then
            MsgBox("Data Pewangi yang dihapus belum dipilih")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan Nama_Pewangi = " & a & " ? ", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK) Then
                Koneksi()
                cmd = New OdbcCommand("delete from tb_pewangi where Nama_Pewangi='" & a & "'", con)
                cmd.ExecuteNonQuery()
                MsgBox("Menghapus data pewangi BERHASIL", vbInformation, "INFORMASI")
                con.Close()
                tampil_data()
            End If
        End If
    End Sub
    Private Sub FormDataPewangi_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call tampil_data()
        Call bersih()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call Koneksi()
        Dim simpan As String = "insert into tb_pewangi values ('" & TextBox2.Text & "')"
        cmd = New OdbcCommand(simpan, con)
        cmd.ExecuteNonQuery()
        MsgBox("Data berhasil disimpan")
        Call tampil_data()
        Call bersih()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim data As Integer
        data = DataGridView1.CurrentRow.Index

        TextBox2.Text = DataGridView1.Item(0, data).Value
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Call delete()
        Call bersih()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox3.TextChanged
        Call Koneksi()
        cmd = New OdbcCommand("select * from tb_pewangi where Nama_Pewangi Like '%" & TextBox3.Text & "%'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Call Koneksi()
            da = New OdbcDataAdapter("select * from tb_pewangi where Nama_Pewangi Like '%" & TextBox3.Text & "%'", con)
            ds = New DataSet
            da.Fill(ds, "tb_pewangi")
            DataGridView1.DataSource = ds.Tables("tb_pewangi")
            DataGridView1.ReadOnly = True
        End If
    End Sub
End Class