Imports System.Data.Odbc
Public Class FormDataOrder
    Private Sub tampil_data()
        da = New OdbcDataAdapter("select No_Order, Tgl_Order, Tgl_Rencana_Selesai, Jml_Pakaian, Total_Harga, Nama_Pewangi, Kode_Pelanggan, Nama_Pelanggan, Status from tb_pesanan order by No_Order asc", con)
        ds = New DataSet
        da.Fill(ds, "tb_pesanan")
        DataGridView1.DataSource = ds.Tables("tb_pesanan")
    End Sub
    Private Sub isi_order()
        Dim data As Integer
        data = DataGridView1.CurrentRow.Index
        If DataGridView1.Item(8, data).Value = "Sudah diambil" Then
            MsgBox("Transaksi Telah Selesai!")
        Else
            FormTransaksiCucianKeluar.TextBox3.Text = DataGridView1.Item(0, data).Value
            FormTransaksiCucianKeluar.TextBox6.Text = DataGridView1.Item(1, data).Value
            FormTransaksiCucianKeluar.TextBox7.Text = DataGridView1.Item(2, data).Value
            FormTransaksiCucianKeluar.TextBox9.Text = DataGridView1.Item(3, data).Value
            FormTransaksiCucianKeluar.TextBox12.Text = DataGridView1.Item(4, data).Value
            FormTransaksiCucianKeluar.TextBox8.Text = DataGridView1.Item(5, data).Value
            FormTransaksiCucianKeluar.TextBox4.Text = DataGridView1.Item(6, data).Value
            FormTransaksiCucianKeluar.TextBox5.Text = DataGridView1.Item(7, data).Value
            FormTransaksiCucianKeluar.ComboBox1.Text = DataGridView1.Item(8, data).Value

            da = New OdbcDataAdapter("select tb_pesanan_detail.No_Order, tb_pesanan_detail.Kode_Jasa, tb_pesanan_detail.Nama_Jasa, tb_pesanan_detail.Satuan, tb_pesanan_detail.Harga, tb_pesanan_detail.Jumlah, tb_pesanan_detail.Sub_Total from tb_pesanan_detail inner join tb_pesanan on tb_pesanan_detail.No_Order = tb_pesanan.No_Order where tb_pesanan.No_Order = '" & DataGridView1.Item(0, data).Value & "'", con)
            ds = New DataSet
            da.Fill(ds, "tb_pesanan_detail")
            FormTransaksiCucianKeluar.DataGridView1.DataSource = ds.Tables("tb_pesanan_detail")
            FormTransaksiCucianKeluar.DataGridView1.ReadOnly = True

            FormTransaksiCucianKeluar.DataGridView1.Columns("No_Order").HeaderText = "No Order"
            FormTransaksiCucianKeluar.DataGridView1.Columns("Kode_Jasa").HeaderText = "Kode Jasa"
            FormTransaksiCucianKeluar.DataGridView1.Columns("Nama_Jasa").HeaderText = "Nama Jasa"
            FormTransaksiCucianKeluar.DataGridView1.Columns("Satuan").HeaderText = "Satuan"
            FormTransaksiCucianKeluar.DataGridView1.Columns("Harga").HeaderText = "Harga"
            FormTransaksiCucianKeluar.DataGridView1.Columns("Jumlah").HeaderText = "Jumlah"
            FormTransaksiCucianKeluar.DataGridView1.Columns("Sub_Total").HeaderText = "Sub Total"
        End If
    End Sub
    Private Sub FormDataOrder_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call tampil_data()
    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Call isi_order()
        Dim data As Integer
        data = DataGridView1.CurrentRow.Index
        Me.Close()
    End Sub

    Private Sub TextBox5_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox5.TextChanged
        Call Koneksi()
        cmd = New OdbcCommand("select * from tb_pesanan where No_Order Like '%" & TextBox5.Text & "%'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Call Koneksi()
            da = New OdbcDataAdapter("select * from tb_pesanan where No_Order Like '%" & TextBox5.Text & "%'", con)
            ds = New DataSet
            da.Fill(ds, "tb_pesanan")
            DataGridView1.DataSource = ds.Tables("tb_pesanan")
            DataGridView1.ReadOnly = True
        End If
    End Sub
End Class