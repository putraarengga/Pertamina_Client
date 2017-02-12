Imports Spire.Barcode
Imports System.IO
Imports System.Net.NetworkInformation
Imports System.ComponentModel


Public Class FormMenu
    Shared Property loggedIn As Integer
    Shared Property statusLog As Integer
    Shared Property idUser As Integer
    Shared Property idJenisUser As Integer
    Shared Property dbOnline As Boolean
    Shared Property arrName As New List(Of String)
    Shared Property arrValue As New List(Of String)

    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BarcodeSettings.ApplyKey("KTWS5-S17CF-B3LKE-FXT34-DVRUH")
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        Me.Width = screenWidth
        Me.Height = screenHeight
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\icons\IDSF.ico")
        Me.Icon = img
        Me.BackgroundImage = System.Drawing.Image.FromFile(appPath + "\background.jpg")
        Me.BackgroundImageLayout = ImageLayout.Stretch
        PictureBox1.ImageLocation = appPath + ("\icons\Button-Blank-Red-icon.png")
        

        Dim sData() As String

        Using sr As New StreamReader("Setting.csv")
            While Not sr.EndOfStream
                sData = sr.ReadLine().Split(","c)

                arrName.Add(sData(0).Trim())
                arrValue.Add(sData(1).Trim())
            End While
        End Using

        ComboBox1.Text = arrValue(3)

        FormLogin.MdiParent = Me
        FormLogin.Show()
        FormLogin.Focus()
        loggedIn = 0
        TransaksiToolStripMenuItem.Enabled = False
    End Sub

    Private Sub NetworkSettingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NetworkSettingToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Form2.MdiParent = Me
            Form2.Show()
            Form2.Focus()

        End If
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        Frm_Main.Close()
        FormLogin.MdiParent = Me
        FormLogin.Show()
        FormLogin.Focus()
        Dim appPath As String = Application.StartupPath()
        PictureBox1.ImageLocation = appPath + ("\icons\Button-Blank-Red-icon.png")
        Lbl_connection.Text = "DISCONNECTED"
        Lbl_User.Text = ""
        Lbl_JenisUser.Text = ""
        LogoutToolStripMenuItem.Text = "Login"
    End Sub

    Private Sub TransaksiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransaksiToolStripMenuItem.Click

        Frm_Main.MdiParent = Me
        Frm_Main.Show()
        Frm_Main.Focus()
        FormMenu.loggedIn = 1
    End Sub

    Private Sub ServerControlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ServerControlToolStripMenuItem.Click

    End Sub

    Private Sub DataUserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataUserToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataUser.MdiParent = Me
            formDataUser.Show()
            formDataUser.Focus()

        End If
    End Sub

    Private Sub DataKendaraanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataKendaraanToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataKendaraan.MdiParent = Me
            formDataKendaraan.Show()
            formDataKendaraan.Focus()

        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Lbl_Date.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub



    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub DataJenisUserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataJenisUserToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataJenisUser.MdiParent = Me
            formDataJenisUser.Show()
            formDataJenisUser.Focus()

        End If
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub DataTujuanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataTujuanToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataTujuan.MdiParent = Me
            formDataTujuan.Show()
            formDataTujuan.Focus()

        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim host As String = "103.247.8.246" ' use any other machine name
        Dim pingreq As Ping = New Ping()

        Try
            Dim rep As PingReply = pingreq.Send(host)
            Label2.Text = "Ping = " + rep.RoundtripTime.ToString
        Catch ex As Exception
            Label2.Text = "Server Connection Problem!"
        End Try



    End Sub
End Class