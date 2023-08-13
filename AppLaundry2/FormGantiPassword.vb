Imports System.Data.Odbc
Public Class FormGantiPassword
    Private Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        CheckBox1.Checked = False
        CheckBox2.Checked = False
    End Sub
    Private Sub FormGantiPassword_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.UseSystemPasswordChar = True
        TextBox2.UseSystemPasswordChar = True
        Call bersih()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            cmd = New OdbcCommand("select * from tb_user where Username ='" & FormLogin.TextBox1.Text & "' and Password_Petugas='" & TextBox1.Text & "'", con)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                TextBox2.Enabled = True
                TextBox3.Enabled = True
            Else
                MsgBox("Password Lama Salah !")
                TextBox1.Text = ""
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Password Baru Harus di Isi")
        Else
            If TextBox2.Text <> TextBox3.Text Then
                MsgBox("Password Baru dan Kofirmasi Harus Sama !")
            Else
                Call Koneksi()
                Dim editData As String = "update tb_user set Password_Petugas='" & TextBox2.Text & "' where Username='" & FormLogin.TextBox1.Text & "'"
                cmd = New OdbcCommand(editData, con)
                cmd.ExecuteNonQuery()
                MsgBox("Password Berhasil Di Update")
            End If
        End If
        Call bersih()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.UseSystemPasswordChar = False
        Else
            TextBox1.UseSystemPasswordChar = True
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.UseSystemPasswordChar = False
        Else
            TextBox2.UseSystemPasswordChar = True
        End If
    End Sub
End Class