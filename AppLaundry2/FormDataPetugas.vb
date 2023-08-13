Imports System.Data.Odbc
Public Class FormDataPetugas
    Private Sub tampil_data()
        da = New OdbcDataAdapter("select Kode_Petugas, Nama_Petugas, Username from tb_user", con)
        ds = New DataSet
        da.Fill(ds, "tb_user")
        DataGridView1.DataSource = ds.Tables("tb_user")
        TextBox4.UseSystemPasswordChar = True
    End Sub
    Private Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox4.Enabled = True
        TextBox5.Text = ""
    End Sub
    Sub delete()
        Dim a As String = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        If a = "" Then
            MsgBox("Data User yang dihapus belum dipilih")
        Else
            If (MessageBox.Show("Anda yakin menghapus data dengan Kode_Petugas = " & a & " ? ", "Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK) Then
                Koneksi()
                cmd = New OdbcCommand("delete from tb_user where Kode_Petugas='" & a & "'", con)
                cmd.ExecuteNonQuery()
                MsgBox("Menghapus data user BERHASIL", vbInformation, "INFORMASI")
                con.Close()
                tampil_data()
            End If
        End If
    End Sub
    Private Sub otomatis()
        Call Koneksi()
        Dim tampil As String = "select * from tb_user order by Kode_Petugas desc"
        cmd = New OdbcCommand(tampil, con)
        dr = cmd.ExecuteReader()
        dr.Read()
        If Not dr.HasRows Then
            TextBox1.Text = "USR" + "0001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(dr.Item("Kode_Petugas").ToString, 4, 3)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "USR000" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 2 Then
                TextBox1.Text = "USR00" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 3 Then
                TextBox1.Text = "USR0" & TextBox1.Text & ""
            End If
        End If

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call Koneksi()
        Dim simpan As String = "insert into tb_user values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
        cmd = New OdbcCommand(simpan, con)
        cmd.ExecuteNonQuery()
        MsgBox("Data berhasil disimpan")
        Call tampil_data()
        Call bersih()
    End Sub

    Private Sub FormDataUser_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call tampil_data()
        Call bersih()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Call delete()
        Call bersih()
    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim data As Integer
        data = DataGridView1.CurrentRow.Index

        TextBox1.Text = DataGridView1.Item(0, data).Value
        TextBox2.Text = DataGridView1.Item(1, data).Value
        TextBox3.Text = DataGridView1.Item(2, data).Value
        TextBox4.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Call Koneksi()
        Dim edit As String = "update tb_user set Nama_Petugas= '" & TextBox2.Text & "',Username= '" & TextBox3.Text & "' where Kode_Petugas='" & TextBox1.Text & "'"
        cmd = New OdbcCommand(edit, con)
        cmd.ExecuteNonQuery()
        MsgBox("Data berhasil diubah", vbInformation, "Pemberitahuan")
        Call tampil_data()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Call Koneksi()
        cmd = New OdbcCommand("select * from tb_user where Nama_Petugas Like '%" & TextBox5.Text & "%'", con)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Call Koneksi()
            da = New OdbcDataAdapter("select * from tb_user where Nama_Petugas Like '%" & TextBox5.Text & "%'", con)
            ds = New DataSet
            da.Fill(ds, "tb_user")
            DataGridView1.DataSource = ds.Tables("tb_user")
            DataGridView1.ReadOnly = True
        End If
        Call bersih()
    End Sub

    Private Sub TextBox4_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox4.TextChanged
        TextBox4.UseSystemPasswordChar = True
    End Sub

End Class