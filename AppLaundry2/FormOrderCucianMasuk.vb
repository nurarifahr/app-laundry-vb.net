Imports System.Data.Odbc
Imports System.Drawing.Printing
Public Class FormOrderCucianMasuk
    Dim WithEvents PD As New PrintDocument
    Dim PPD As New PrintPreviewDialog
    Sub TampilPewangi()
        cmd = New OdbcCommand("select Nama_Pewangi from tb_pewangi", con)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr.Item("Nama_Pewangi"))
        Loop
    End Sub
    Private Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = 0
        TextBox13.Text = 0
        TextBox14.Text = 0
        ComboBox1.Text = "-Pilih-"
        ComboBox2.Text = "-Pilih-"
        DataGridView1.Rows.Clear()
        TextBox8.Text = "0"
    End Sub
    Private Sub bersih_jasa()
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox13.Focus()
    End Sub

    Private Sub hitungtotal()
        'TextBox8.Text = "Rp.    " & (From row As DataGridViewRow In DataGridView1.Rows Where row.Cells(5).FormattedValue.ToString() <> String.Empty Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
        Dim hitung As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitung = hitung + DataGridView1.Rows(i).Cells(5).Value
            TextBox8.Text = hitung
        Next
    End Sub

    Private Sub FormOrderCucianMasuk_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call TampilPewangi()
        Call hitungtotal()
        Call bersih()
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        FormDaftarPelanggan.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        FormDaftarJasa.ShowDialog()
    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged
        Dim harga As Single
        Dim jumlah As Integer
        Dim subTotal As Single

        If TextBox13.Text = "" Then
            TextBox14.Text = 0
        Else
            harga = CSng(TextBox12.Text)
            jumlah = CInt(TextBox13.Text)

            subTotal = harga * jumlah
        End If
        TextBox14.Text = subTotal
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        DataGridView1.Rows.Add(New String() {TextBox9.Text, TextBox10.Text, TextBox12.Text, TextBox11.Text, TextBox13.Text, TextBox14.Text})
        Call hitungtotal()
        Call bersih_jasa()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Or TextBox1.Text = "" Then
            MsgBox("Data wajib diisi semua !")
        Else
            Dim simpanMasuk As String = "insert into tb_pesanan values ('" & TextBox1.Text & "', '" & TextBox5.Text & "', '" & TextBox6.Text & "', '" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & ComboBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & FormLogin.TextBox1.Text & "', '" & ComboBox2.Text & "')"
            cmd = New OdbcCommand(simpanMasuk, con)
            cmd.ExecuteNonQuery()

            For baris As Integer = 0 To DataGridView1.Rows.Count - 2
                Dim simpanDetail As String = "insert into tb_pesanan_detail values('" & TextBox1.Text & "', '" & DataGridView1.Rows(baris).Cells(0).Value & "', '" & DataGridView1.Rows(baris).Cells(1).Value & "', '" & DataGridView1.Rows(baris).Cells(3).Value & "', '" & DataGridView1.Rows(baris).Cells(2).Value & "', '" & DataGridView1.Rows(baris).Cells(4).Value & "', '" & DataGridView1.Rows(baris).Cells(5).Value & "')"
                cmd = New OdbcCommand(simpanDetail, con)
                cmd.ExecuteNonQuery()
            Next
            MsgBox("Data berhasil disimpan")
        End If
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
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
        Dim f10 As New Font("Times New Rowan", 10, FontStyle.Regular)
        Dim f10b As New Font("Times New Rowan", 10, FontStyle.Bold)
        Dim f12 As New Font("Times New Rowan", 12, FontStyle.Bold)

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

        e.Graphics.DrawString("Nota Order (Cucian Masuk)", f6b, Brushes.Black, centermargin, 32, tengah)
        e.Graphics.DrawString(TextBox5.Text, f6, Brushes.Black, 5, 46)

        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, 48)

        e.Graphics.DrawString("No. Order", f5b, Brushes.Black, 5, 62)
        e.Graphics.DrawString(TextBox1.Text, f5b, Brushes.Black, 90, 62)
        e.Graphics.DrawString("Tgl. Order", f5, Brushes.Black, 5, 71)
        e.Graphics.DrawString(TextBox5.Text, f5, Brushes.Black, 90, 71)
        e.Graphics.DrawString("Tgl. Rencana Selesai", f5, Brushes.Black, 5, 80)
        e.Graphics.DrawString(TextBox6.Text, f5, Brushes.Black, 90, 80)
        e.Graphics.DrawString("Jumlah Pakaian", f5, Brushes.Black, 5, 89)
        e.Graphics.DrawString(TextBox7.Text, f5, Brushes.Black, 90, 89)

        e.Graphics.DrawString("Nama Pewangi", f5, Brushes.Black, 300, 62)
        e.Graphics.DrawString(ComboBox1.Text, f5, Brushes.Black, 390, 62)
        e.Graphics.DrawString("Nama Pelanggan", f5, Brushes.Black, 300, 71)
        e.Graphics.DrawString(TextBox3.Text, f5, Brushes.Black, 390, 71)
        e.Graphics.DrawString("Status Cucian", f5, Brushes.Black, 300, 80)
        e.Graphics.DrawString(ComboBox2.Text, f5, Brushes.Black, 390, 80)
        e.Graphics.DrawString("Petugas", f5, Brushes.Black, 300, 89)
        e.Graphics.DrawString(FormLogin.TextBox1.Text, f5, Brushes.Black, 390, 89)

        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, 98)

        e.Graphics.DrawString("Kode Jasa", f5, Brushes.Black, 5, 113)
        e.Graphics.DrawString("Nama Jasa", f5, Brushes.Black, 70, 113)
        e.Graphics.DrawString("Satuan", f5, Brushes.Black, 170, 113)
        e.Graphics.DrawString("Harga", f5, Brushes.Black, 250, 113)
        e.Graphics.DrawString("Jumlah", f5, Brushes.Black, 350, 113)
        e.Graphics.DrawString("Sub Total", f5, Brushes.Black, 450, 113)

        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, 114)

        DataGridView1.AllowUserToAddRows = False
        Dim tinggi As Integer
        For baris As Integer = 0 To DataGridView1.RowCount - 1
            tinggi += 15
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(0).Value.ToString, f5, Brushes.Black, 5, 114 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(1).Value.ToString, f5, Brushes.Black, 70, 114 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(3).Value.ToString, f5, Brushes.Black, 170, 114 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(2).Value.ToString, f5, Brushes.Black, 250, 114 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(4).Value.ToString, f5, Brushes.Black, 350, 114 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(5).Value.ToString, f5, Brushes.Black, rightmargin - 15, 114 + tinggi, kanan)

            e.Graphics.DrawString(garisPutus, f5, Brushes.Black, 5, 120 + tinggi)
        Next
        tinggi = 130 + tinggi

        e.Graphics.DrawString("Catatan : Harap nota ini dibawa saat pengambilan cucian", f4, Brushes.Red, 5, tinggi)
        e.Graphics.DrawString("Total Harga", f5b, Brushes.Black, 380, tinggi)
        e.Graphics.DrawString(TextBox8.Text, f5b, Brushes.Black, rightmargin - 15, tinggi, kanan)
        e.Graphics.DrawString(garisLurus, f8b, Brushes.Black, 5, tinggi + 1)

    End Sub

End Class