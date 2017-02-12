Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Net

Public Class FormLogin
    Dim connect As MySqlConnection
    Dim command As MySqlCommand
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim simpan As String
    Dim tmpAbsensi, tmpDate, tmpTime As String
    Dim countAbsensi As Integer
    Dim tmpCount As Integer
    Dim selectDataBase As String
    Dim count As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        connect = New MySqlConnection
        connect.ConnectionString = "server=" + FormMenu.arrValue(1) + ";userid=admin_idsf;password=123456;database=idsf"
        'connect.ConnectionString = "server=localhost;userid=root;password=r7pqv6s6Xc9QbZKK;database=idsf"
        Dim reader As MySqlDataReader
        Dim userFound As Boolean = False
        Dim FullName As String = ""
        Dim idJenisUser As Integer

        'ProgressBar1.Visible = True
        'Timer1.Enabled = True
        'Label3.Visible = True
        FormMenu.Lbl_connection.Text = "DISCONNECTED"
        If CheckForInternetConnection() Then
            Try
                connect.Open()
                Dim Query As String

                Query = String.Format("SELECT tdatauser.IDUser, tdatauser.NamaUser,tdatauser.Password,tdatauser.NamaLengkap,tdatauser.IDJenisUser,tjenisuser.JenisUser " +
                                      "FROM tdatauser join tjenisuser on tdatauser.IDJenisUser = tjenisuser.IDJenisUser " +
                                      "WHERE tdatauser.NamaUser = '{0}' AND tdatauser.Password = '{1}' AND (tjenisuser.JenisUser = '{2}' OR tjenisuser.JenisUser = 'admin')",
                                      Me.TB_Nama.Text.Trim(), Me.TB_Password.Text.Trim(), FormMenu.arrValue(3))
                command = New MySqlCommand(Query, connect)
                reader = command.ExecuteReader

                Dim count As Integer
                FormMenu.statusLog = 0
                count = 0
                While reader.Read
                    count = count + 1
                    userFound = True
                    FullName = reader("NamaLengkap").ToString
                    FormMenu.idJenisUser = reader("IDJenisUser")
                    FormMenu.idUser = reader("IDUser")
                End While
                connect.Close()

                If count < 1 Then
                    MsgBox("Sorry, username or password not found", MsgBoxStyle.OkOnly, "Invalid Login")
                End If
                If userFound = True Then
                    TB_Nama.Clear()
                    TB_Password.Clear()
                    Hide()
                    FormMenu.Enabled = True
                    FormMenu.Show()
                    FormMenu.MenuStrip1.Enabled = True
                    FormMenu.TransaksiToolStripMenuItem.Enabled = True
                    FormMenu.LogoutToolStripMenuItem.Text = "Logout"

                    Frm_Main.Show()
                    Me.Close()
                    Frm_Main.MdiParent = FormMenu
                    Frm_Main.Show()
                    Frm_Main.Focus()
                    FormMenu.loggedIn = 1
                    FormMenu.Lbl_User.Text = FullName
                    GetIDUser()
                    GetJenisUser()
                    If Not FormMenu.idJenisUser = 1 Then
                        FormMenu.DataUserToolStripMenuItem.Enabled = False
                        FormMenu.DataJenisUserToolStripMenuItem.Enabled = False
                        FormMenu.DataKendaraanToolStripMenuItem.Enabled = False
                        FormMenu.DataTujuanToolStripMenuItem.Enabled = False
                    Else
                        FormMenu.DataUserToolStripMenuItem.Enabled = True
                        FormMenu.DataJenisUserToolStripMenuItem.Enabled = True
                        FormMenu.DataKendaraanToolStripMenuItem.Enabled = True
                    End If

                    tmpDate = Format(Date.Now, "yyyy-MM-dd")
                    tmpTime = Format(DateTime.Now, "HH:mm:ss")

                    FormMenu.dbOnline = True
                    FormMenu.Lbl_connection.Text = "CONNECTED"
                    Dim appPath As String = Application.StartupPath()
                    FormMenu.PictureBox1.ImageLocation = appPath + ("\icons\Button-Blank-Green-icon.png")


                    'If countAbsensi = 0 And Not (FormMenu.statusLog Mod 2 = 1) Then

                    '    Try
                    '        connect.Open()
                    '        Query = String.Format("SELECT * FROM tabsensi ORDER BY status DESC LIMIT 1")
                    '        command = New MySqlCommand(Query, connect)
                    '        reader = command.ExecuteReader

                    '        While reader.Read
                    '            FormMenu.statusLog = reader("status")
                    '        End While
                    '        connect.Close()
                    '    Catch ex As MySqlException
                    '        MessageBox.Show(ex.Message)
                    '    Finally
                    '        connect.Dispose()
                    '    End Try

                    '    FormMenu.statusLog = FormMenu.statusLog + 1
                    '    tmpAbsensi = Format(DateTime.Now, "yyyyMMdd") & FormMenu.idUser & FormMenu.statusLog

                    '    simpan = "INSERT INTO tabsensi(IDAbsensi, IDUser, tglMasuk, wktMasuk) " +
                    '               "VALUES ('" & tmpAbsensi & "','" & FormMenu.idUser & "','" & tmpDate & "','" & tmpTime & "')"

                    '    jalankansql(simpan)
                    'End If

                End If



            Catch ex As MySqlException
                MessageBox.Show(ex.Message)
            Finally
                connect.Dispose()
            End Try

        End If

    End Sub
    Private Sub GetIDUser()
        selectDataBase = "SELECT * FROM tdatauser WHERE NamaLengkap='" & FormMenu.Lbl_User.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            FormMenu.idUser = DT.Rows(0).Item("IDUser")
        End If
    End Sub
    Private Sub GetJenisUser()
        selectDataBase = "SELECT * FROM tjenisuser WHERE IDJenisUser='" & FormMenu.idJenisUser & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            FormMenu.Lbl_JenisUser.Text = DT.Rows(0).Item("JenisUser")
        End If
    End Sub
    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\icons\IDSF.ico")
        Me.Icon = img
        Me.BackgroundImage = System.Drawing.Image.FromFile(appPath + "\black.jpg")
        Me.BackgroundImageLayout = ImageLayout.Stretch
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
            'MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Private Sub periksasql(ByVal sQL As String)
        Dim objcmd As New System.Data.Odbc.OdbcCommand
        bukaDB()
        Try
            objcmd.Connection = konek
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQL
            countAbsensi = Convert.ToInt16(objcmd.ExecuteScalar())
            objcmd.Dispose()
            'MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub
    Public Shared Function CheckForInternetConnection() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://103.247.8.246")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        count = count + 1
        ProgressBar1.Value = count
        If count = 100 Then
            count = 0
        End If
    End Sub

    Private Sub TB_Password_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_Password.KeyDown
        If e.KeyCode = Keys.Enter Then
            btn_login.PerformClick()
        End If
    End Sub

    Private Sub TB_Nama_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_Nama.KeyDown
        If e.KeyCode = Keys.Enter Then
            btn_login.PerformClick()
        End If
    End Sub
End Class
