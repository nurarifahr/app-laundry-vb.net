Imports System.Data
Imports System.Data.Odbc
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.ReportSource
Public Class FormCetakOrder
    Dim con As OdbcConnection
    Dim dr As OdbcDataReader
    Dim da As OdbcDataAdapter
    Dim ds As DataSet
    Dim dt As DataTable
    Dim cmd As OdbcCommand
    Dim rpt As New ReportDocument

    Sub koneksi()
        con = New OdbcConnection
        con.ConnectionString = "dsn=dblaundry_uas"
        con.Open()
    End Sub

    Public Sub tampilSemuaData()
        koneksi()
        rpt.Load(Application.StartupPath & "\LaporanCetakOrder.rpt")
        da = New OdbcDataAdapter("select No_Order, Tgl_Order, Tgl_Rencana_Selesai, Jml_Pakaian, Total_Harga, Nama_Pelanggan, Status from tb_pesanan", con)
        dt = New DataTable
        da.Fill(dt)

        rpt.SetDataSource(dt)

        CrystalReportViewer1.ReportSource = rpt
        CrystalReportViewer1.Refresh()
    End Sub
    Public Sub tampilSudahDiambil()
        koneksi()
        rpt.Load(Application.StartupPath & "\LaporanCetakOrder.rpt")
        da = New OdbcDataAdapter("select No_Order, Tgl_Order, Tgl_Rencana_Selesai, Jml_Pakaian, Total_Harga, Nama_Pelanggan, Status from tb_pesanan where Status ='" & FormCetakLaporan.RadioButton2.Text & "'", con)
        dt = New DataTable
        da.Fill(dt)

        rpt.SetDataSource(dt)

        CrystalReportViewer1.ReportSource = rpt
        CrystalReportViewer1.Refresh()
    End Sub
    Public Sub tampilBelumDiambil()
        koneksi()
        rpt.Load(Application.StartupPath & "\LaporanCetakOrder.rpt")
        da = New OdbcDataAdapter("select No_Order, Tgl_Order, Tgl_Rencana_Selesai, Jml_Pakaian, Total_Harga, Nama_Pelanggan, Status from tb_pesanan where Status ='" & FormCetakLaporan.RadioButton3.Text & "'", con)
        dt = New DataTable
        da.Fill(dt)

        rpt.SetDataSource(dt)

        CrystalReportViewer1.ReportSource = rpt
        CrystalReportViewer1.Refresh()
    End Sub
    Private Sub FormCetakOrder_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        koneksi()
    End Sub
End Class