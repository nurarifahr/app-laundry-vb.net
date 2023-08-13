Imports System.Data.Odbc
Module Module1
    Public con As OdbcConnection
    Public dr As OdbcDataReader
    Public da As OdbcDataAdapter
    Public ds As DataSet
    Public dt As DataTable
    Public cmd As OdbcCommand

    Public Sub Koneksi()
        con = New OdbcConnection
        con.ConnectionString = "dsn=dblaundry_uas"
        If con.State = ConnectionState.Closed Then con.Open()
    End Sub
End Module
