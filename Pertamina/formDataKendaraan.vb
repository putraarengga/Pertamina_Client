Public Class formDataKendaraan

    Dim databaru As Boolean
    Dim selectDataBase, vJenisUser, tmpString As String
    Dim indexSatuan, indexKategori, indexLokasi As Integer
    Dim dataSelected As Boolean
    Dim indeksKendaraan As String

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub FormDataUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        dataSelected = False
        Dim appPath As String = Application.StartupPath()
        'Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        'Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\Truck_supplier.ico")
        IsiGrid()
        TextBox2.Enabled = False
    End Sub
    Sub IsiGrid()
        selectDataBase = "SELECT tkendaraan.namaPerusahaan,tkendaraan.noPolKendaraan,tkendaraan.namaSopir,tkendaraan.namaKernet,tkendaraan.kapasitasTruk,tkendaraan.IDKendaraan,tkendaraan.callCenter " +
                        " FROM tkendaraan "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tkendaraan")
        DataGridView1.DataSource = (DS.Tables("tkendaraan"))
        DataGridView1.Enabled = True
    End Sub
    Sub Bersih()

        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""


    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Bersih()
        TextBox3.Focus()
        databaru = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT tkendaraan.namaPerusahaan,tkendaraan.noPolKendaraan,tkendaraan.namaSopir,tkendaraan.namaKernet,tkendaraan.kapasitasTruk,tkendaraan.IDKendaraan,tkendaraan.callCenter  " +
                                          "FROM tkendaraan WHERE namaPerusahaan LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tkendaraan")
            DataGridView1.DataSource = (DS.Tables("tkendaraan"))
            DataGridView1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim simpan As String
        Dim pesan As String

        If TextBox4.Text = "" Then Exit Sub
        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, "IDSF")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO tkendaraan(IDKendaraan,noPolKendaraan,namaSopir,namaKernet,kapasitasTruk,callCenter,namaPerusahaan) " +
                     "VALUES (LAST_INSERT_ID(),'" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')"
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE tkendaraan SET noPolKendaraan= '" & TextBox3.Text & "', namaSopir = '" & TextBox4.Text & "',namaKernet= '" & TextBox5.Text & "',kapasitasTruk= '" & TextBox6.Text & "',callCenter= '" & TextBox7.Text & "',namaPerusahaan= '" & TextBox8.Text & "'  WHERE IDKendaraan= '" & TextBox2.Text & "' "
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False

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
            TextBox2.Text = DataGridView1.Rows(x).Cells(5).Value
            TextBox3.Text = DataGridView1.Rows(x).Cells(1).Value
            TextBox4.Text = DataGridView1.Rows(x).Cells(2).Value
            TextBox5.Text = DataGridView1.Rows(x).Cells(3).Value
            TextBox6.Text = DataGridView1.Rows(x).Cells(4).Value
            TextBox7.Text = DataGridView1.Rows(x).Cells(6).Value
            TextBox8.Text = DataGridView1.Rows(x).Cells(0).Value

        Catch ex As Exception
        End Try
    End Sub

    Sub GetIndeks(ByVal x As Integer)
        Try
            indeksKendaraan = DataGridView1.Rows(x).Cells(1).Value.ToString
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Bersih()
        isitextbox(e.RowIndex)
        GetIndeks(e.RowIndex)
        dataSelected = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        databaru = False
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? " + TextBox4.Text, vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        If pesan = MsgBoxResult.No Then Exit Sub

        hapussql = "DELETE FROM tkendaraan WHERE tkendaraan.IDKendaraan ='" & TextBox2.Text & "'"
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

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If dataSelected = True Then
            Frm_Main.indexKendaraan = indeksKendaraan
            Frm_Main.GetDataKendaraan()
            Me.Close()
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If dataSelected = True Then
            Frm_Main.indexKendaraan = indeksKendaraan
            Frm_Main.GetDataKendaraan()
            Me.Close()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox1.Text = "" Then
            DataGridView1.Refresh()
            IsiGrid()
        Else
            DataGridView1.Refresh()
            bukaDB()
            DA = New Odbc.OdbcDataAdapter("SELECT tkendaraan.namaPerusahaan,tkendaraan.noPolKendaraan,tkendaraan.namaSopir,tkendaraan.namaKernet,tkendaraan.kapasitasTruk,tkendaraan.IDKendaraan,tkendaraan.callCenter  " +
                                          "FROM tkendaraan WHERE namaPerusahaan LIKE '%" & TextBox1.Text & "%'", konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tkendaraan")
            DataGridView1.DataSource = (DS.Tables("tkendaraan"))
            DataGridView1.Enabled = True
        End If
    End Sub
End Class