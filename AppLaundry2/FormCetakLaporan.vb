Imports System.Data.Odbc
Public Class FormCetakLaporan
    Private Sub FormCetakLaporan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            FormCetakOrder.Show()
            FormCetakOrder.tampilSemuaData()
        ElseIf RadioButton2.Checked = True Then
            FormCetakOrder.Show()
            FormCetakOrder.tampilSudahDiambil()
        Else
            FormCetakOrder.Show()
            FormCetakOrder.tampilBelumDiambil()
        End If
    End Sub

    'Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged
    '   da = New OdbcDataAdapter("select No_Order, Tgl_Order, Tgl_Rencana_Selesai, Jml_Pakaian, Total_Harga, Nama_Pelanggan, Status from tb_pesanan", con)
    '  ds = New DataSet
    ' da.Fill(ds, "tb_pesanan")
    ' DataGridView1.DataSource = ds.Tables("tb_pesanan")
    'End Sub

    ' Private Sub RadioButton2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton2.CheckedChanged
    '    da = New OdbcDataAdapter("select No_Order, Tgl_Order, Tgl_Rencana_Selesai, Jml_Pakaian, Total_Harga, Nama_Pelanggan, Status from tb_pesanan where Status ='" & RadioButton2.Text & "'"", con)
    '     ds = New DataSet
    '      da.Fill(ds, "tb_pesanan")
    '       DataGridView1.DataSource = ds.Tables("tb_pesanan")
    '    End Sub

    'Private Sub RadioButton3_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton3.CheckedChanged
    'da = New OdbcDataAdapter("select No_Order, Tgl_Order, Tgl_Rencana_Selesai, Jml_Pakaian, Total_Harga, Nama_Pelanggan, Status from tb_pesanan where Status ='" & RadioButton3.Text & "'"", con)
    ' ds = New DataSet
    '  da.Fill(ds, "tb_pesanan")
    '   DataGridView1.DataSource = ds.Tables("tb_pesanan")
    'End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class