Public Class formDataJenisUser

    Dim databaru As Boolean
    Dim selectDataBase, vJenisUser, tmpString As String
    Dim indexSatuan, indexKategori, indexLokasi As Integer

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub FormDataUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        Dim appPath As String = Application.StartupPath()
        'Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        'Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Bokehlicia-Captiva-System-users.ico")
        IsiGrid()
        TextBox2.Enabled = False
    End Sub
    Sub IsiGrid()
        selectDataBase = "SELECT * from tjenisuser"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tjenisuser")
        DataGridView1.DataSource = (DS.Tables("tjenisuser"))
        DataGridView1.Enabled = True
    End Sub
    Sub Bersih()

        TextBox2.Text = ""
        TextBox3.Text = ""

    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bersih()
        TextBox3.Focus()
        databaru = True
        TextBox3.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = True

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT * FROM tjenisuser WHERE JenisUser LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tjenisuser")
            DataGridView1.DataSource = (DS.Tables("tjenisuser"))
            DataGridView1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim simpan As String
        Dim pesan As String

        If TextBox3.Text = "" Then Exit Sub
        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, "IDSF")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO tjenisuser(IDJenisUser,JenisUser) " +
                     "VALUES (LAST_INSERT_ID(),'" & TextBox3.Text & "')"

            Button2.Enabled = True
            Button3.Enabled = False
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, "IDSF")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE tjenisuser SET JenisUser= '" & TextBox3.Text & "' WHERE IDJenisUser= '" & TextBox2.Text & "' "

            Button2.Enabled = True
            Button3.Enabled = False
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        TextBox3.Enabled = False
    End Sub
    Private Sub jalankansql(ByVal sQL As String)
        Dim objcmd As New System.Data.Odbc.OdbcCommand
        bukaDB()
        Try
            objcmd.Connection = konek
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQL
            objcmd.ExecuteNonQuery()
            objcmd.Dispose()
            MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Sub isitextbox(ByVal x As Integer)
        Try
            TextBox2.Text = DataGridView1.Rows(x).Cells(0).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(1).Value
        Catch ex As Exception
        End Try
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Bersih()
        isitextbox(e.RowIndex)
        TextBox3.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        databaru = False
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + TextBox3.Text, vbExclamation + MsgBoxStyle.YesNo, "IDSF")
        If pesan = MsgBoxResult.No Then Exit Sub

        hapussql = "DELETE FROM tjenisuser WHERE tjenisuser.IDJenisUser ='" & TextBox2.Text & "'"
        If TextBox2.Text = "" Then Exit Sub
        jalankansql(hapussql)
        DataGridView1.Refresh()
        IsiGrid()
    End Sub

    Private Sub FormDataUser_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class