Imports System.Data.Odbc
Public Class FormMenuUtama
    Private Sub LOGOUTToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LOGOUTToolStripMenuItem.Click
        MsgBox("Anda yakin akan keluar", MsgBoxStyle.YesNo)
        Me.Close()
        FormLogin.ShowDialog()
        
    End Sub

    Private Sub EXITToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EXITToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub USERToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles USERToolStripMenuItem.Click
        FormDataPetugas.ShowDialog()
    End Sub

    Private Sub PELANGGANToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PELANGGANToolStripMenuItem.Click
        FormDataPelanggan.ShowDialog()
    End Sub

    Private Sub JASAToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles JASAToolStripMenuItem.Click
        FormDataJasa.ShowDialog()
    End Sub

    Private Sub PEWANGIToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PEWANGIToolStripMenuItem.Click
        FormDataPewangi.ShowDialog()
    End Sub

    Private Sub ORDERCUCIANMASUKToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ORDERCUCIANMASUKToolStripMenuItem.Click
        FormOrderCucianMasuk.ShowDialog()
    End Sub

    Private Sub CUCIANKELUARToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CUCIANKELUARToolStripMenuItem.Click
        FormTransaksiCucianKeluar.ShowDialog()
    End Sub

    Private Sub LAPORANToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LAPORANToolStripMenuItem.Click
        FormCetakLaporan.ShowDialog()
    End Sub

    Private Sub GANTIPASSWORDToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GANTIPASSWORDToolStripMenuItem.Click
        FormGantiPassword.ShowDialog()
    End Sub
End Class