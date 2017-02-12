Imports Spire.Barcode
Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Data.Odbc
Imports System.ComponentModel
Imports System.IO

Public Class Frm_Main

    Dim connect As MySqlConnection
    Dim command As MySqlCommand
    Dim databaru As Boolean
    Dim selectDataBase As String
    Shared Property indexKendaraan As String
    Shared Property indexTujuan As String
    Dim counter As Integer
    Dim pesan, simpan, TextToPrint, keterangan As String
    Dim IDTujuan, IDUser, IDKendaraan, IDDistribusi As Integer
    Dim isLevelChecked As Boolean
    Dim statusDistribusi As String


    Private Sub Frm_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\icons\IDSF.ico")
        Me.Icon = img
        GetCounter()
        IsiGrid()
        databaru = False
        PrintDocument1.PrinterSettings.PrinterName = FormMenu.arrValue(2)
        tb_9.Focus()
        isLevelChecked = False
    End Sub

    Sub IsiGridPencarian()
        selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO,tdatatujuan.namaTujuan,tkendaraan.noPolKendaraan,tdistribusi.Keterangan,tdatauser.NamaUser,tdistribusi.dataBarcode, " +
                        " tdistribusi.tglMuat,tdistribusi.wktMuat,tdistribusi.wktSampai,tdistribusi.tglSampai,tkendaraan.callCenter,tkendaraan.kapasitasTruk, tdistribusi.tempatLoading " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatatujuan ON tdatatujuan.IDTujuan = tdistribusi.IDTujuan " +
                        " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUser WHERE tdistribusi.dataBarcode LIKE '%" & Trim(tb_9.Text) & "%'"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = DS.Tables("tdistribusi")
        DataGridView1.Sort(DataGridView1.Columns(6), ListSortDirection.Descending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Transportir"
            .Columns(1).HeaderCell.Value = "Nomor DO"
            .Columns(2).HeaderCell.Value = "Tempat Tujuan"
            .Columns(3).HeaderCell.Value = "No Pol Kendaraan"
            .Columns(4).HeaderCell.Value = "Keterangan"
            .Columns(5).HeaderCell.Value = "User Server"
            .Columns(6).HeaderCell.Value = "Barcode"
            .Columns(7).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(8).HeaderCell.Value = "Waktu Pengiriman"
            .Columns(9).HeaderCell.Value = "Tanggal Sampai"
            .Columns(10).HeaderCell.Value = "Waktu Sampai"
            .Columns(11).HeaderCell.Value = "Call Center"
            .Columns(12).HeaderCell.Value = "Liter"
            .Columns(13).HeaderCell.Value = "Tempat Loading"
        End With
        DataGridView1.Enabled = True
    End Sub
    Sub IsiGrid()
        Dim tanggalSekarang As String
        tanggalSekarang = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO,tdatatujuan.namaTujuan,tkendaraan.noPolKendaraan,tdistribusi.Keterangan,tdatauser.NamaUser,tdistribusi.dataBarcode, " +
                        " tdistribusi.tglMuat,tdistribusi.wktMuat,tdistribusi.wktSampai,tdistribusi.tglSampai,tkendaraan.callCenter,tkendaraan.kapasitasTruk, tdistribusi.tempatLoading " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatatujuan ON tdatatujuan.IDTujuan = tdistribusi.IDTujuan " +
                        " JOIN tdatauser ON tdistribusi.IDUser = tdatauser.IDUser WHERE tdistribusi.tglMuat='" + tanggalSekarang + "' AND (tdatatujuan.namaTujuan = '" + FormMenu.arrValue(3) + "' OR tdistribusi.tempatLoading = '" + FormMenu.arrValue(3) + "')"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = DS.Tables("tdistribusi")
        DataGridView1.Sort(DataGridView1.Columns(6), ListSortDirection.Descending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Transportir"
            .Columns(1).HeaderCell.Value = "Nomor DO"
            .Columns(2).HeaderCell.Value = "Tempat Tujuan"
            .Columns(3).HeaderCell.Value = "No Pol Kendaraan"
            .Columns(4).HeaderCell.Value = "Keterangan"
            .Columns(5).HeaderCell.Value = "User Server"
            .Columns(6).HeaderCell.Value = "Barcode"
            .Columns(7).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(8).HeaderCell.Value = "Waktu Pengiriman"
            .Columns(9).HeaderCell.Value = "Tanggal Sampai"
            .Columns(10).HeaderCell.Value = "Waktu Sampai"
            .Columns(11).HeaderCell.Value = "Call Center"
            .Columns(12).HeaderCell.Value = "Liter"
            .Columns(13).HeaderCell.Value = "Tempat Loading"
        End With
        DataGridView1.Enabled = True
    End Sub
    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Form2.Show()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        databaru = True
        Bersih()
        tb_9.Text = Format(DateTime.Now, "yyyyMMdd") & counter
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Bersih()
        isitextbox(e.RowIndex)
        databaru = False
        GetIDDistribusi()
        If statusDistribusi = "LEVEL CHECKED" Then
            btn_4.Enabled = True
            btn_5.Enabled = True
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles btn_4.Click
        Dim tmpWaktu, tmpTanggal As String
        tmpWaktu = DateTimePicker3.Value.ToString("hh:mm:00")
        tmpTanggal = DateTimePicker4.Value.ToString("yyyy-MM-dd")
        If tb_5.Text = "" Or tb_6.Text = "" Or tb_7.Text = "" Or tb_9.Text = "" Then
            MsgBox("Tidak bisa menyimpan data ke server. (Data Kurang Lengkap)")
            Return

        End If
        pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, "IDSF_Client")
        If pesan = MsgBoxResult.No Then
            Exit Sub
        End If
        GetIDKendaraan()
        GetIDDistribusi()
        simpan = "UPDATE tdistribusi SET IDKendaraan = '" & IDKendaraan & "'," +
                   " IDKendaraan = '" & IDKendaraan & "',IDTujuan = '" & IDTujuan & "',NoDO='" & tb_6.Text & "',wktSampai='" & tmpWaktu & "',tglSampai='" & tmpTanggal & "', " +
                   " dataBarcode='" & tb_9.Text & "' , IDUserClient='" & FormMenu.idUser & "',Keterangan= 'ACCEPTED'" +
                  " WHERE IDDistribusi = '" & IDDistribusi & "'"

        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        Bersih()

    End Sub

    Private Sub GroupBox8_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub isitextbox(ByVal x As Integer)
        Try

            tb_1.Text = DataGridView1.Rows(x).Cells(3).Value
            tb_5.Text = DataGridView1.Rows(x).Cells(0).Value
            tb_6.Text = DataGridView1.Rows(x).Cells(1).Value
            tb_9.Text = DataGridView1.Rows(x).Cells(6).Value
            tb_7.Text = DataGridView1.Rows(x).Cells(2).Value
            statusDistribusi = DataGridView1.Rows(x).Cells(4).Value

            indexKendaraan = tb_1.Text
            indexTujuan = tb_7.Text
            GetDataKendaraan()
            GetDataTujuan()
            If DataGridView1.Rows(x).Cells(7).Value.ToString = "" Then
                DateTimePicker1.CustomFormat = " "  'An empty SPACE
                DateTimePicker1.Format = DateTimePickerFormat.Custom
            Else
                DateTimePicker1.CustomFormat = "dd/MM/yyyy"
                DateTimePicker1.Value = DataGridView1.Rows(x).Cells(7).Value
                DateTimePicker1.Format = DateTimePickerFormat.Custom
                'DateTimePicker1.CustomFormat = "yyyy-MM-dd"
                'vdate = DateTimePicker1.Value.Year + "-" + DateTimePicker1.Value.Month + "-" + DateTimePicker1.Value.Day
            End If

        Catch ex As Exception
        End Try
    End Sub

    Sub Bersih()
        'tb_1.Enabled = True
        'tb_2.Enabled = True
        'tb_3.Enabled = True
        'tb_4.Enabled = True
        'tb_5.Enabled = True
        'tb_6.Enabled = True
        'tb_7.Enabled = True
        'tb_8.Enabled = True
        'tb_9.Enabled = True
        'btn_1.Enabled = True
        'btn_2.Enabled = True
        btn_3.Enabled = True
        btn_4.Enabled = True
        btn_5.Enabled = True
        DateTimePicker3.Enabled = True
        DateTimePicker4.Enabled = True

        tb_1.Text = ""
        tb_2.Text = ""
        tb_3.Text = ""
        tb_4.Text = ""
        tb_5.Text = ""
        tb_6.Text = ""
        tb_7.Text = ""
        tb_8.Text = ""
        tb_9.Text = " "
        DateTimePicker3.Value = DateTime.Now
        DateTimePicker4.Value = DateTime.Now
        tb_9.Focus()
        Button1.Enabled = False
        isLevelChecked = False


    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub


    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btn_1.Click
        formDataKendaraan.Button6.Visible = True
        formDataKendaraan.Show()
        formDataKendaraan.Focus()
    End Sub

    Sub GetDataKendaraan()

        selectDataBase = "SELECT * FROM tkendaraan WHERE noPolKendaraan='" & indexKendaraan & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            tb_1.Text = DT.Rows(0).Item("noPolKendaraan")
            tb_2.Text = DT.Rows(0).Item("namaSopir")
            tb_3.Text = DT.Rows(0).Item("namaKernet")
            tb_4.Text = DT.Rows(0).Item("kapasitasTruk")
            tb_5.Text = DT.Rows(0).Item("namaPerusahaan")
        End If
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btn_2.Click
        formDataTujuan.Button6.Visible = True
        formDataTujuan.Show()
        formDataTujuan.Focus()
        GetDataTujuan()


    End Sub

    Sub GetDataTujuan()
        selectDataBase = "SELECT * FROM tdatatujuan WHERE namaTujuan='" & indexTujuan & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            tb_7.Text = DT.Rows(0).Item("namaTujuan")
            tb_8.Text = DT.Rows(0).Item("alamatTujuan")
            IDTujuan = DT.Rows(0).Item("IDTujuan")
        End If
    End Sub


    Sub GetIDDistribusi()
        selectDataBase = "SELECT * FROM tdistribusi WHERE NoDO = '" & tb_6.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            IDDistribusi = DT.Rows(0).Item("IDDistribusi")
        End If
    End Sub

    Sub GetCounter()
        selectDataBase = "SELECT * FROM tdistribusi ORDER BY IDDistribusi DESC LIMIT 1"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            counter = DT.Rows(0).Item("IDDistribusi") + 1
        End If
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles tb_5.TextChanged

    End Sub

    Private Sub BarCodeControl1_Click(sender As Object, e As EventArgs) Handles BarCodeControl1.Click

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles tb_9.TextChanged
        Dim tmpString As String
        Dim SearchForThis As String = "2016"
        Dim FirstCharacter As Integer
        If tb_9.Text = "" Then

            tb_9.Text = " "

        Else
            tmpString = tb_9.Text
            Button1.Enabled = True
            FirstCharacter = tmpString.IndexOf(SearchForThis)
            If FirstCharacter = 0 Then
                If tb_9.TextLength > 12 Then
                    tb_9.Text = "201"
                    tb_9.Select(tb_9.Text.Length + 1, 1)
                End If

            End If
        End If
        If tb_9.Text = " " Or tb_9.Text = "" Then
            Button1.Enabled = False
            btn_4.Enabled = False
            btn_5.Enabled = False
        End If
        BarCodeControl1.Data = tb_9.Text
        BarCodeControl1.Data2D = tb_9.Text
    End Sub
    Private Sub GetIDKendaraan()
        selectDataBase = "SELECT * FROM tkendaraan WHERE noPolKendaraan='" & tb_1.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            IDKendaraan = DT.Rows(0).Item("IDKendaraan")
        End If
    End Sub
    Private Sub GetIDTujuan()
        selectDataBase = "SELECT * FROM tdatatujuan WHERE namaTujuan='" & tb_7.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            IDTujuan = DT.Rows(0).Item("IDTujuan")
        End If
    End Sub

    Private Sub btn_3_Click(sender As Object, e As EventArgs) Handles btn_3.Click
        Bersih()
        IsiGrid()

        'IsiGridPencarian()
        'isitextbox(0)
        'Dim tmpWaktu, tmpTanggal As String
        'tmpWaktu = DateTimePicker1.Value.ToString("hh:mm:00")
        'tmpTanggal = DateTimePicker2.Value.ToString("yyyy-MM-dd")
        'GetIDKendaraan()
        'GetIDTujuan()
        'If tb_5.Text = "" Or tb_6.Text = "" Or tb_7.Text = "" Or tb_9.Text = "" Then
        '    MsgBox("Tidak bisa menyimpan data ke server. (Data Kurang Lengkap)")
        '    Return

        'End If

        'If databaru Then
        '    pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, vbInformation)
        '    If pesan = MsgBoxResult.No Then
        '        Exit Sub
        '    End If
        '    simpan = "INSERT INTO tdistribusi(IDKendaraan,IDUser, " +
        '                " NamaPerusahaan,IDTujuan,NoDO,wktMuat,tglMuat, " +
        '                " dataBarcode,Keterangan) " +
        '             "VALUES ('" & IDKendaraan & "','" & FormMenu.idUser & "'" +
        '                ",'" & tb_5.Text & "','" & IDTujuan & "','" & tb_6.Text & "','" & tmpWaktu & "','" & tmpTanggal & "'" +
        '                ",'" & tb_9.Text & "','REGISTERED')"
        'Else
        '    pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, vbInformation)
        '    If pesan = MsgBoxResult.No Then
        '        Exit Sub
        '    End If
        '    simpan = "UPDATE tdistribusi SET IDKendaraan = '" & IDKendaraan & "',IDUser = '" & FormMenu.idUser & "', " +
        '               " NamaPerusahaan = '" & tb_5.Text & "',IDTujuan = '" & IDTujuan & "',NoDO='" & tb_6.Text & "',wktMuat='" & tmpWaktu & "',tglMuat='" & tmpTanggal & "', " +
        '               " dataBarcode='" & tb_9.Text & "' " +
        '              " WHERE IDDistribusi = '" & IDDistribusi & "'"

        '    'simpan = "UPDATE tdatatujuan SET namaTujuan= '" & TextBox4.Text & "', alamatTujuan = '" & TextBox5.Text & "' WHERE IDTujuan= '" & TextBox2.Text & "' "
        'End If
        'jalankansql(simpan)
        'DataGridView1.Refresh()
        'IsiGrid()
        'Bersih()
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

    Private Sub btn_5_Click(sender As Object, e As EventArgs) Handles btn_5.Click
        Dim tmpWaktu, tmpTanggal As String
        tmpWaktu = DateTimePicker3.Value.ToString("hh:mm:00")
        tmpTanggal = DateTimePicker4.Value.ToString("yyyy-MM-dd")
        If tb_5.Text = "" Or tb_6.Text = "" Or tb_7.Text = "" Or tb_9.Text = "" Then
            MsgBox("Tidak bisa menyimpan data ke server. (Data Kurang Lengkap)")
            Return

        End If
        pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, "IDSF_Client")
        If pesan = MsgBoxResult.No Then
            Exit Sub
        End If
        GetIDKendaraan()
        GetIDDistribusi()
        simpan = "UPDATE tdistribusi SET IDKendaraan = '" & IDKendaraan & "'," +
                   " IDTujuan = '" & IDTujuan & "',NoDO='" & tb_6.Text & "',wktSampai='" & tmpWaktu & "',tglSampai='" & tmpTanggal & "', " +
                   " dataBarcode='" & tb_9.Text & "' ,Keterangan= 'REJECTED'" +
                  " WHERE IDDistribusi = '" & IDDistribusi & "'"

        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        Bersih()

    End Sub

    Private Sub btn_7_Click(sender As Object, e As EventArgs)
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? ", vbExclamation + MsgBoxStyle.YesNo, "Perhatian")
        If pesan = MsgBoxResult.No Then Exit Sub

        hapussql = "DELETE FROM tdistribusi WHERE NoDO ='" & tb_6.Text & "'"
        If tb_6.Text = "" Then Exit Sub
        jalankansql(hapussql)
        DataGridView1.Refresh()
        IsiGrid()
        Bersih()

    End Sub
    Public Sub PrintHeader()
        TextToPrint &= " "
        TextToPrint &= Environment.NewLine
        TextToPrint &= Environment.NewLine


        TextToPrint &= Environment.NewLine
        Dim StringToPrint As String = "Untuk Pelanggan"
        Dim LineLen As Integer = StringToPrint.Length
        Dim spcLen1 As New String(" "c, Math.Round((17 - LineLen))) 'This line is used to center text in the middle of the receipt
        TextToPrint &= spcLen1 & StringToPrint & Environment.NewLine


        TextToPrint &= Environment.NewLine
        StringToPrint = "Standard-PLN"
        LineLen = StringToPrint.Length
        Dim spcLen2 As New String(" "c, Math.Round((14 - LineLen)))
        TextToPrint &= spcLen2 & StringToPrint & Environment.NewLine

        StringToPrint = "2601-Kota Jayapura"
        LineLen = StringToPrint.Length
        Dim spcLen3 As New String(" "c, Math.Round((20 - LineLen)))
        TextToPrint &= spcLen3 & StringToPrint & Environment.NewLine
        TextToPrint &= Environment.NewLine

        StringToPrint = "SURAT PENGANTAR PENGIRIMAN"
        LineLen = StringToPrint.Length
        Dim spcLen4 As New String(" "c, Math.Round((33 - LineLen)))
        TextToPrint &= spcLen4 & StringToPrint & Environment.NewLine

        StringToPrint = "=================================================="
        LineLen = StringToPrint.Length
        Dim spcLen10 As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen10 & StringToPrint & Environment.NewLine
    End Sub

    Public Sub ItemsToBePrinted1()

        TextToPrint &= " "
        Dim globalLengt As Integer = 0

        Dim NamaPerusahaan As String = tb_5.Text.ToString()
        Dim NomorDO As String = tb_6.Text.ToString()
        Dim Nopol As String = tb_1.Text.ToString()
        Dim NamaSopir As String = tb_2.Text.ToString()
        Dim NamaKernet As String = tb_3.Text.ToString()
        Dim Kapasitas As String = tb_4.Text.ToString()
        Dim Tujuan As String = tb_7.Text.ToString()
        Dim AlamatTujuan As String = tb_8.Text.ToString()
        Dim Barcode As String = tb_9.Text.ToString()
        Dim TanggalPengiriman As String = DateTimePicker1.Text.ToString()
        Dim WaktuPengiriman As String = DateTimePicker2.Text.ToString()

        Dim StringToPrint As String = "No.Pol.Kendaraan      :"
        Dim StringToPrint2 As String = Nopol
        Dim LineLen As String = StringToPrint.Length
        Dim LineLen2 As String = StringToPrint2.Length
        globalLengt = StringToPrint.Length
        Dim spcLen1 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen1b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen1 & StringToPrint & spcLen1b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Shippment No           :"
        StringToPrint2 = IDDistribusi
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen2 As New String(" "c, Math.Round(24 - LineLen))
        Dim spcLen2b As New String(" "c, Math.Round(1))
        TextToPrint &= spcLen2 & StringToPrint & spcLen2b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nama Pengemudi         :"
        StringToPrint2 = NamaSopir
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen3 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen3b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen3 & StringToPrint & spcLen3b & StringToPrint2 & Environment.NewLine


        StringToPrint = "Tujuan                 :"
        StringToPrint2 = Tujuan
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen4 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen4b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen4 & StringToPrint & spcLen4b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nomor DO               :"
        StringToPrint2 = NomorDO
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen5b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen5 & StringToPrint & spcLen5b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Kapasitas Tangki       :"
        StringToPrint2 = Kapasitas
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen6 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen6b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen6 & StringToPrint & spcLen6b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Tanggal Pengiriman     :"
        StringToPrint2 = TanggalPengiriman
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen7 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen7b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen7 & StringToPrint & spcLen7b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Waktu Pengiriman       :"
        StringToPrint2 = WaktuPengiriman
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen8 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen8b As New String(" "c, Math.Round((2)))
        TextToPrint &= spcLen8 & StringToPrint & spcLen8b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nomor Barcode          :"
        StringToPrint2 = Barcode
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen9 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen9b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen9 & StringToPrint & spcLen9b & StringToPrint2 & Environment.NewLine

        StringToPrint = "=================================================="
        LineLen = StringToPrint.Length
        Dim spcLen10 As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen10 & StringToPrint & Environment.NewLine

    End Sub

    Public Sub printFooter()

        TextToPrint &= Environment.NewLine & Environment.NewLine & Environment.NewLine & Environment.NewLine
        Dim globalLengt As Integer = 0

        TextToPrint &= Environment.NewLine & Environment.NewLine
        Dim StringToPrint As String = "TERIMA KASIH ATAS KEPERCAYAAN ANDA "
        Dim LineLen As String = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen5 & StringToPrint

        StringToPrint = "MENGGUNAKAN PRODUK PERTAMINA"
        LineLen = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen6 As New String(" "c, Math.Round((5)))
        TextToPrint &= Environment.NewLine & spcLen6 & StringToPrint & Environment.NewLine

    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Static currentChar As Integer
        Dim textfont As Font = New Font("Courier New", 10, FontStyle.Bold)

        Dim h, w As Integer
        Dim left, top As Integer
        With PrintDocument1.DefaultPageSettings
            h = 0
            w = 0
            left = 0
            top = 0
        End With


        Dim lines As Integer = CInt(Math.Round(h / 1))
        Dim b As New Rectangle(left, top, w, h)
        Dim format As StringFormat
        format = New StringFormat(StringFormatFlags.LineLimit)
        Dim line, chars As Integer
        Dim appPath As String = Application.StartupPath()
        Dim newImage As Image = Image.FromFile(appPath + "\logoPertamina.png")
        Dim newImage2 As Image = Image.FromFile(appPath + "\barcode.png")

        ' Create Point for upper-left corner of image.
        Dim ulCorner As New Point(200, 20)
        Dim ulCorner2 As New Point(35, 280)

        ' Draw image to screen.
        e.Graphics.DrawImage(newImage, ulCorner)
        e.Graphics.DrawImage(newImage2, ulCorner2)
        printFooter()
        e.Graphics.MeasureString(Mid(TextToPrint, currentChar + 1), textfont, New SizeF(w, h), format, chars, line)
        e.Graphics.DrawString(TextToPrint.Substring(currentChar, chars), New Font("Courier New", 9, FontStyle.Bold), Brushes.Black, b, format)


        currentChar = currentChar + chars
        If currentChar < TextToPrint.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            currentChar = 0
        End If
    End Sub



    Private Sub btn_9_Click(sender As Object, e As EventArgs) Handles btn_9.Click
        Dim tanggalAwal, tanggalAkhir As String
        bukaDB()
        DateTimePicker5.Format = DateTimePickerFormat.Custom
        DateTimePicker6.Format = DateTimePickerFormat.Custom

        DateTimePicker5.CustomFormat = "dd/MMM/yyyy"
        DateTimePicker6.CustomFormat = "dd/MMM/yyyy"

        tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-dd")
        tanggalAkhir = Format(DateTimePicker6.Value.Date, "yyyy-MM-dd")
        If tanggalAwal = tanggalAkhir Then
            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO,tdatatujuan.namaTujuan,tkendaraan.noPolKendaraan,tdistribusi.Keterangan,tdatauser.NamaUser,tdistribusi.dataBarcode, " +
                       " tdistribusi.tglMuat,tdistribusi.wktMuat,tdistribusi.wktSampai,tdistribusi.tglSampai,tkendaraan.callCenter,tkendaraan.kapasitasTruk, tdistribusi.tempatLoading " +
                       " FROM tdistribusi " +
                       " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                       " JOIN tdatatujuan ON tdatatujuan.IDTujuan = tdistribusi.IDTujuan " +
                       " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUser WHERE tdistribusi.tglMuat LIKE '%" & tanggalAkhir & "%'  AND (tdatatujuan.namaTujuan = '" + FormMenu.arrValue(3) + "' OR tdistribusi.tempatLoading = '" + FormMenu.arrValue(3) + "')"
            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        Else
            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO,tdatatujuan.namaTujuan,tkendaraan.noPolKendaraan,tdistribusi.Keterangan,tdatauser.NamaUser,tdistribusi.dataBarcode, " +
                        " tdistribusi.tglMuat,tdistribusi.wktMuat,tdistribusi.wktSampai,tdistribusi.tglSampai,tkendaraan.callCenter,tkendaraan.kapasitasTruk, tdistribusi.tempatLoading " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatatujuan ON tdatatujuan.IDTujuan = tdistribusi.IDTujuan " +
                        " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUser WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'  AND (tdatatujuan.namaTujuan = '" + FormMenu.arrValue(3) + "' OR tdistribusi.tempatLoading = '" + FormMenu.arrValue(3) + "')"
            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)

        End If
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = (DS.Tables("tdistribusi"))
        DataGridView1.Enabled = True
        DataGridView1.Sort(DataGridView1.Columns(6), ListSortDirection.Descending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Transportir"
            .Columns(1).HeaderCell.Value = "Nomor DO"
            .Columns(2).HeaderCell.Value = "Tempat Tujuan"
            .Columns(3).HeaderCell.Value = "No Pol Kendaraan"
            .Columns(4).HeaderCell.Value = "Keterangan"
            .Columns(5).HeaderCell.Value = "User Server"
            .Columns(6).HeaderCell.Value = "Barcode"
            .Columns(7).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(8).HeaderCell.Value = "Waktu Pengiriman"
            .Columns(9).HeaderCell.Value = "Tanggal Sampai"
            .Columns(10).HeaderCell.Value = "Waktu Sampai"
            .Columns(11).HeaderCell.Value = "Call Center"
            .Columns(12).HeaderCell.Value = "Liter"
            .Columns(13).HeaderCell.Value = "Tempat Loading"
        End With
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        IsiGridPencarian()
        isitextbox(0)

        If statusDistribusi = "ACCEPTED" Or statusDistribusi = "REJECTED" Then
            MsgBox("Data yang anda cari sudah di " + statusDistribusi, MsgBoxStyle.OkOnly, "IDSF_Client")
            DataGridView1.Refresh()
            IsiGrid()
            Bersih()

        ElseIf isLevelChecked = False And Not statusDistribusi = "LEVEL CHECKED" Then

            btn_4.Enabled = False
            btn_5.Enabled = False
            GetStatusDistribusi()


            'pesan = MsgBox("Apakah anda yakin level minyak telah sesuai? " + tb_4.Text + " liter", MsgBoxStyle.YesNo, "IDSF_Client")
            pesan = MsgBox("Truck pertamina masuk Dan pengecekan level minyak", MsgBoxStyle.OkCancel, "IDSF_Client")

            GetIDKendaraan()
            GetIDDistribusi()
            If pesan = MsgBoxResult.Cancel Then
                isLevelChecked = False
                IsiGrid()
                Bersih()
                Exit Sub
            Else
                simpan = "UPDATE tdistribusi SET IDKendaraan = '" & IDKendaraan & "', " +
                          "Keterangan= 'LEVEL CHECKED'" +
                         " WHERE IDDistribusi = '" & IDDistribusi & "'"
                jalankansql(simpan)
                DataGridView1.Refresh()
                IsiGrid()
                Bersih()
                isLevelChecked = True
                'update levelchecked pada tdistribusi
            End If
        Else

            btn_4.Enabled = True
            btn_5.Enabled = True
            connect = New MySqlConnection

            connect.ConnectionString = "server=" + FormMenu.arrValue(1) + ";userid=admin_idsf;password=123456;database=idsf"
            'connect.ConnectionString = "server=localhost;userid=root;password=r7pqv6s6Xc9QbZKK;database=idsf"

            Dim reader As MySqlDataReader

            connect.Open()
            Dim Query As String
            Dim idJenisUserClient As Integer = 0
            Query = String.Format("SELECT * FROM tdatauser WHERE NamaUser = '" & DataGridView1.Rows(0).Cells(5).Value.ToString & "' ")
            command = New MySqlCommand(Query, connect)
            reader = command.ExecuteReader

            While reader.Read
                idJenisUserClient = reader("IDJenisUser")

            End While
            connect.Close()

            Dim tmpWaktu, tmpTanggal As String
            tmpWaktu = DateTimePicker3.Value.ToString("hh:mm:00")
            tmpTanggal = DateTimePicker4.Value.ToString("yyyy-MM-dd")
            If tb_5.Text = "" Or tb_6.Text = "" Or tb_7.Text = "" Or tb_9.Text = "" Then
                MsgBox("Tidak bisa menyimpan data ke server. (Data Kurang Lengkap)")
                Return

            End If
            GetIDKendaraan()
            GetIDDistribusi()
            If tb_7.Text = FormMenu.arrValue(3) Or FormMenu.idJenisUser = 1 Then
                simpan = "UPDATE tdistribusi SET IDKendaraan = '" & IDKendaraan & "', " +
                           "IDTujuan = '" & IDTujuan & "',NoDO='" & tb_6.Text & "',wktSampai='" & tmpWaktu & "',tglSampai='" & tmpTanggal & "', " +
                           " dataBarcode='" & tb_9.Text & "' , IDUserClient='" & FormMenu.idUser & "',Keterangan= 'ACCEPTED',tempatLoading= '" + FormMenu.arrValue(3) + "'" +
                          " WHERE IDDistribusi = '" & IDDistribusi & "'"
                pesan = MsgBox("DATA ACCEPTED. Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, "IDSF_Client")
                If pesan = MsgBoxResult.No Then
                    Exit Sub
                End If
            Else
                simpan = "UPDATE tdistribusi SET IDKendaraan = '" & IDKendaraan & "', " +
                           " IDTujuan = '" & IDTujuan & "',NoDO='" & tb_6.Text & "',wktSampai='" & tmpWaktu & "',tglSampai='" & tmpTanggal & "', " +
                           " dataBarcode='" & tb_9.Text & "' , IDUserClient='" & FormMenu.idUser & "',Keterangan= 'REJECTED',tempatLoading= '" + FormMenu.arrValue(3) + "'" +
                          " WHERE IDDistribusi = '" & IDDistribusi & "'"
                pesan = MsgBox("DATA REJECTED. Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, "IDSF_Client")


                If pesan = MsgBoxResult.No Then
                    Exit Sub
                Else
                    keterangan = InputBox("Keterangan ", "DATA REJECTED", "", MousePosition.X, MousePosition.Y)

                End If
            End If

            jalankansql(simpan)
            DataGridView1.Refresh()
            IsiGrid()
            Bersih()
        End If

    End Sub
    Sub GetStatusDistribusi()

    End Sub

End Class
