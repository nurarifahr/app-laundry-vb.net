Imports System.Data.Odbc
Public Class FormDaftarPelanggan
    Private Sub tampil_data()
        da = New OdbcDataAdapter("select Kode_Pelanggan, Nama_Pelanggan, Alamat_Pelanggan, No_Telp from tb_pelanggan", con)
        ds = New DataSet
        da.Fill(ds, "tb_pelanggan")
        DataGridView1.DataSource = ds.Tables("tb_pelanggan")
    End Sub
    Private Sub isi_pelanggan()
        Dim data As Integer
        data = DataGridView1.CurrentRow.Index

        FormOrderCucianMasuk.TextBox2.Text = DataGridView1.Item(0, data).Value
        FormOrderCucianMasuk.TextBox3.Text = DataGridView1.Item(1, data).Value
        FormOrderCucianMasuk.TextBox4.Text = DataGridView1.Item(2, data).Value
    End Sub
    
    Private Sub FormDaftarPelanggan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call tampil_data()
        TextBox5.Text = ""
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Call isi_pelanggan()
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

End Class