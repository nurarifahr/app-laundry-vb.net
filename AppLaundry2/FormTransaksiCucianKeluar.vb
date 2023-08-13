Imports System.Data.Odbc
Imports System.Drawing.Printing
Public Class FormTransaksiCucianKeluar
    Dim WithEvents PD As New PrintDocument
    Dim PPD As New PrintPreviewDialog
    Private Sub bersih()
        DataGridView1.Columns.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = "0"
        TextBox11.Text = "0"
        TextBox12.Text = "0"
        ComboBox1.Text = "  "
        TextBox8.Text = ""
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        FormDataOrder.ShowDialog()
    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged
        Dim total_harga As Single
        Dim dibayar As Single
        Dim kembali As Single

        If TextBox10.Text = "" Then
            TextBox11.Text = "0"
        Else
            total_harga = CSng(TextBox12.Text)
            dibayar = CSng(TextBox10.Text)

            kembali = dibayar - total_harga
        End If
        TextBox11.Text = kembali
    End Sub

    Private Sub FormTransaksiCucianKeluar_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Call Koneksi()
        Dim simpan As String = "insert into tb_transaksi values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox10.Text & "', '" & TextBox11.Text & "','" & FormLogin.TextBox1.Text & "')"
        cmd = New OdbcCommand(simpan, con)
        cmd.ExecuteNonQuery()

        If ComboBox1.Text = "Sudah diambil" Then
            Call Koneksi()
            Dim edit As String = "update tb_pesanan set Status= '" & ComboBox1.Text & "'where No_Order='" & TextBox3.Text & "'"
            cmd = New OdbcCommand(edit, con)
            cmd.ExecuteNonQuery()
        End If
        MsgBox("Data berhasil disimpan")
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Call bersih()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        PPD.Document = PD
        PPD.ShowDialog()
    End Sub

    Private Sub PD_BeginPrint(sender As Object, e As System.Drawing.Printing.PrintEventArgs) Handles PD.BeginPrint
        Dim pagestup As New PageSettings
        pagestup.PaperSize = New PaperSize("Custom", 500, 250)
        PD.DefaultPageSettings = pagestup
    End Sub

    Private Sub PD_PrintPage(sender As System.Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PD.PrintPage
        Dim f5 As New Font("Times New Rowan", 5, FontStyle.Regular)
        Dim f5b As New Font("Times New Rowan", 5, FontStyle.Bold)
        Dim f6 As New Font("Times New Rowan", 6, FontStyle.Regular)
        Dim f6b As New Font("Times New Rowan", 6, FontStyle.Bold)
        Dim f8 As New Font("Times New Rowan", 8, FontStyle.Regular)
        Dim f8b As New Font("Times New Rowan", 8, FontStyle.Bold)
        Dim f4 As New Font("Times New Rowan", 4, FontStyle.Italic)

        Dim leftmargin As Integer = PD.DefaultPageSettings.Margins.Left
        Dim centermargin As Integer = PD.DefaultPageSettings.PaperSize.Width / 2
        Dim rightmargin As Integer = PD.DefaultPageSettings.PaperSize.Width

        Dim kanan As New StringFormat
        Dim tengah As New StringFormat
        kanan.Alignment = StringAlignment.Far
        tengah.Alignment = StringAlignment.Center

        Dim garisPutus As String
        garisPutus = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Dim garisLurus As String
        garisLurus = "__________________________________________________________________________"

        e.Graphics.DrawString("LAUNDRY", f8b, Brushes.Black, 5, 5)
        e.Graphics.DrawString("Jl. Indonesia Maju No. 22 Surakarta", f6, Brushes.Black, 5, 16)
        e.Graphics.DrawString("Telp. 085231313056", f6, Brushes.Black, 5, 26)

        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, 30)
        e.Graphics.DrawString("Nota Transaksi (Cucian Keluar)", f6b, Brushes.Black, centermargin, 43, tengah)
        e.Graphics.DrawString(TextBox2.Text, f5, Brushes.Black, rightmargin - 15, 45, kanan)
        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, 45)

        e.Graphics.DrawString("No. Transaksi", f5b, Brushes.Black, 5, 59)
        e.Graphics.DrawString(TextBox1.Text, f5b, Brushes.Black, 90, 59)
        e.Graphics.DrawString("Tgl. Transaksi", f5, Brushes.Black, 5, 68)
        e.Graphics.DrawString(TextBox2.Text, f5, Brushes.Black, 90, 68)
        e.Graphics.DrawString("No.Order", f5, Brushes.Black, 5, 77)
        e.Graphics.DrawString(TextBox3.Text, f5, Brushes.Black, 90, 77)
        e.Graphics.DrawString("Tgl. Order", f5, Brushes.Black, 5, 86)
        e.Graphics.DrawString(TextBox6.Text, f5, Brushes.Black, 90, 86)
        e.Graphics.DrawString("Jumlah Pakaian", f5, Brushes.Black, 5, 95)
        e.Graphics.DrawString(TextBox9.Text, f5, Brushes.Black, 90, 95)

        e.Graphics.DrawString("Nama Pewangi", f5, Brushes.Black, 300, 59)
        e.Graphics.DrawString(TextBox8.Text, f5, Brushes.Black, 390, 59)
        e.Graphics.DrawString("Nama Pelanggan", f5, Brushes.Black, 300, 68)
        e.Graphics.DrawString(TextBox5.Text, f5, Brushes.Black, 390, 68)
        e.Graphics.DrawString("Status Cucian", f5, Brushes.Black, 300, 77)
        e.Graphics.DrawString(ComboBox1.Text, f5, Brushes.Black, 390, 77)
        e.Graphics.DrawString("Petugas", f5, Brushes.Black, 300, 86)
        e.Graphics.DrawString(FormLogin.TextBox1.Text, f5, Brushes.Black, 390, 86)

        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, 97)

        e.Graphics.DrawString("Kode Jasa", f5, Brushes.Black, 5, 112)
        e.Graphics.DrawString("Nama Jasa", f5, Brushes.Black, 70, 112)
        e.Graphics.DrawString("Satuan", f5, Brushes.Black, 170, 112)
        e.Graphics.DrawString("Harga", f5, Brushes.Black, 250, 112)
        e.Graphics.DrawString("Jumlah", f5, Brushes.Black, 350, 112)
        e.Graphics.DrawString("Sub Total", f5, Brushes.Black, 450, 112)

        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, 113)

        DataGridView1.AllowUserToAddRows = False
        Dim tinggi As Integer
        For baris As Integer = 0 To DataGridView1.RowCount - 1
            tinggi += 15
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(1).Value.ToString, f5, Brushes.Black, 5, 113 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(2).Value.ToString, f5, Brushes.Black, 70, 113 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(3).Value.ToString, f5, Brushes.Black, 170, 113 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(4).Value.ToString, f5, Brushes.Black, 250, 113 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(5).Value.ToString, f5, Brushes.Black, 350, 113 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(6).Value.ToString, f5, Brushes.Black, rightmargin - 15, 113 + tinggi, kanan)

            e.Graphics.DrawString(garisPutus, f5, Brushes.Black, 5, 119 + tinggi)
        Next
        tinggi = 129 + tinggi

        e.Graphics.DrawString("Terima Kasih Telah Menggunakan Jasa Kami", f4, Brushes.Red, 20, tinggi + 12)
        e.Graphics.DrawString("Total Harga", f5b, Brushes.Black, 380, tinggi)
        e.Graphics.DrawString(TextBox12.Text, f5b, Brushes.Black, rightmargin - 15, tinggi, kanan)
        e.Graphics.DrawString("Dibayar", f5, Brushes.Black, 380, tinggi + 12)
        e.Graphics.DrawString(TextBox10.Text, f5, Brushes.Black, rightmargin - 15, tinggi + 12, kanan)
        e.Graphics.DrawString("Kembali", f5b, Brushes.Black, 380, tinggi + 24)
        e.Graphics.DrawString(TextBox11.Text, f5b, Brushes.Black, rightmargin - 15, tinggi + 24, kanan)
        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, tinggi + 26)
    End Sub
End Class