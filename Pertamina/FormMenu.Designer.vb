<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ServerControlToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataUserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataJenisUserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataKendaraanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataTujuanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransaksiToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NetworkSettingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogoutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Lbl_Date = New System.Windows.Forms.Label()
        Me.Lbl_User = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Lbl_connection = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Lbl_JenisUser = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ServerControlToolStripMenuItem, Me.TransaksiToolStripMenuItem, Me.NetworkSettingToolStripMenuItem, Me.LogoutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1008, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ServerControlToolStripMenuItem
        '
        Me.ServerControlToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DataUserToolStripMenuItem, Me.DataJenisUserToolStripMenuItem, Me.DataKendaraanToolStripMenuItem, Me.DataTujuanToolStripMenuItem})
        Me.ServerControlToolStripMenuItem.Name = "ServerControlToolStripMenuItem"
        Me.ServerControlToolStripMenuItem.Size = New System.Drawing.Size(82, 20)
        Me.ServerControlToolStripMenuItem.Text = "Data Master"
        '
        'DataUserToolStripMenuItem
        '
        Me.DataUserToolStripMenuItem.Name = "DataUserToolStripMenuItem"
        Me.DataUserToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.DataUserToolStripMenuItem.Text = "Data User"
        '
        'DataJenisUserToolStripMenuItem
        '
        Me.DataJenisUserToolStripMenuItem.Name = "DataJenisUserToolStripMenuItem"
        Me.DataJenisUserToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.DataJenisUserToolStripMenuItem.Text = "Data Jenis User"
        '
        'DataKendaraanToolStripMenuItem
        '
        Me.DataKendaraanToolStripMenuItem.Name = "DataKendaraanToolStripMenuItem"
        Me.DataKendaraanToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.DataKendaraanToolStripMenuItem.Text = "Data Transportir"
        '
        'DataTujuanToolStripMenuItem
        '
        Me.DataTujuanToolStripMenuItem.Name = "DataTujuanToolStripMenuItem"
        Me.DataTujuanToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.DataTujuanToolStripMenuItem.Text = "Data Tujuan"
        '
        'TransaksiToolStripMenuItem
        '
        Me.TransaksiToolStripMenuItem.Name = "TransaksiToolStripMenuItem"
        Me.TransaksiToolStripMenuItem.Size = New System.Drawing.Size(95, 20)
        Me.TransaksiToolStripMenuItem.Text = "Data Distribusi"
        '
        'NetworkSettingToolStripMenuItem
        '
        Me.NetworkSettingToolStripMenuItem.Name = "NetworkSettingToolStripMenuItem"
        Me.NetworkSettingToolStripMenuItem.Size = New System.Drawing.Size(56, 20)
        Me.NetworkSettingToolStripMenuItem.Text = "Setting"
        Me.NetworkSettingToolStripMenuItem.Visible = False
        '
        'LogoutToolStripMenuItem
        '
        Me.LogoutToolStripMenuItem.Name = "LogoutToolStripMenuItem"
        Me.LogoutToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
        Me.LogoutToolStripMenuItem.Text = "Login"
        '
        'Lbl_Date
        '
        Me.Lbl_Date.AutoSize = True
        Me.Lbl_Date.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Lbl_Date.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Date.Location = New System.Drawing.Point(300, 3)
        Me.Lbl_Date.Name = "Lbl_Date"
        Me.Lbl_Date.Size = New System.Drawing.Size(0, 15)
        Me.Lbl_Date.TabIndex = 7
        '
        'Lbl_User
        '
        Me.Lbl_User.AutoSize = True
        Me.Lbl_User.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Lbl_User.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_User.Location = New System.Drawing.Point(541, 3)
        Me.Lbl_User.Name = "Lbl_User"
        Me.Lbl_User.Size = New System.Drawing.Size(0, 15)
        Me.Lbl_User.TabIndex = 8
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Lbl_connection
        '
        Me.Lbl_connection.AutoSize = True
        Me.Lbl_connection.Location = New System.Drawing.Point(24, 5)
        Me.Lbl_connection.Name = "Lbl_connection"
        Me.Lbl_connection.Size = New System.Drawing.Size(92, 13)
        Me.Lbl_connection.TabIndex = 10
        Me.Lbl_connection.Text = "DISCONNECTED"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Lbl_JenisUser)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.Lbl_User)
        Me.Panel1.Controls.Add(Me.Lbl_connection)
        Me.Panel1.Controls.Add(Me.Lbl_Date)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 709)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1008, 21)
        Me.Panel1.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(186, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Label2"
        '
        'Lbl_JenisUser
        '
        Me.Lbl_JenisUser.AutoSize = True
        Me.Lbl_JenisUser.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Lbl_JenisUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_JenisUser.Location = New System.Drawing.Point(671, 3)
        Me.Lbl_JenisUser.Name = "Lbl_JenisUser"
        Me.Lbl_JenisUser.Size = New System.Drawing.Size(0, 15)
        Me.Lbl_JenisUser.TabIndex = 12
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(15, 15)
        Me.PictureBox1.TabIndex = 11
        Me.PictureBox1.TabStop = False
        '
        'ComboBox1
        '
        Me.ComboBox1.DisplayMember = "PLTD WAENA"
        Me.ComboBox1.Enabled = False
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"PLTD WAENA", "PLTD SENTAN TAMA", "PLTD SENTANI", "PLTD ARSO", "PLTD YARMOKH"})
        Me.ComboBox1.Location = New System.Drawing.Point(468, 2)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 15
        Me.ComboBox1.Text = "PLTD WAENA"
        Me.ComboBox1.ValueMember = "PLTD WAENA"
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 4000
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1008, 730)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormMenu"
        Me.Text = "Client  || IDSF (Information Distribution System of Fuel oil)"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents NetworkSettingToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LogoutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransaksiToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ServerControlToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DataUserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DataKendaraanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Lbl_Date As System.Windows.Forms.Label
    Friend WithEvents Lbl_User As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Lbl_connection As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents DataJenisUserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Lbl_JenisUser As System.Windows.Forms.Label
    Friend WithEvents DataTujuanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
