Imports System.Data.Odbc
Public Class FormDaftarJasa
    Private Sub tampil_data()
        da = New OdbcDataAdapter("select Kode_Jasa, Nama_Jasa, Satuan, Harga from tb_jasa", con)
        ds = New DataSet
        da.Fill(ds, "tb_jasa")
        DataGridView1.DataSource = ds.Tables("tb_jasa")
    End Sub
    Private Sub bersih()
        TextBox5.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
    End Sub
    Private Sub isi_jasa()
        Dim data As Integer
        data = DataGridView1.CurrentRow.Index

        FormOrderCucianMasuk.TextBox9.Text = DataGridView1.Item(0, data).Value
        FormOrderCucianMasuk.TextBox10.Text = DataGridView1.Item(1, data).Value
        FormOrderCucianMasuk.TextBox11.Text = DataGridView1.Item(2, data).Value
        FormOrderCucianMasuk.TextBox12.Text = DataGridView1.Item(3, data).Value
    End Sub

    Private Sub FormDaftarJasa_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call tampil_data()
        Call bersih()
    End Sub

    Private Sub TextBox5_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged
        Call Koneksi()
        If (RadioButton1.Checked = True) Then
            cmd = New OdbcCommand("select * from tb_jasa where Nama_Jasa Like '%" & TextBox5.Text & "%'", con)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                Call Koneksi()
                da = New OdbcDataAdapter("select * from tb_jasa where Nama_Jasa Like '%" & TextBox5.Text & "%'", con)
                ds = New DataSet
                da.Fill(ds, "tb_jasa")
                DataGridView1.DataSource = ds.Tables("tb_jasa")
                DataGridView1.ReadOnly = True
            End If
        Else
            cmd = New OdbcCommand("select * from tb_jasa where Kode_Jasa Like '%" & TextBox5.Text & "%'", con)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                Call Koneksi()
                da = New OdbcDataAdapter("select * from tb_jasa where Kode_Jasa Like '%" & TextBox5.Text & "%'", con)
                ds = New DataSet
                da.Fill(ds, "tb_jasa")
                DataGridView1.DataSource = ds.Tables("tb_jasa")
                DataGridView1.ReadOnly = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Call isi_jasa()
        FormOrderCucianMasuk.TextBox13.Focus()
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class