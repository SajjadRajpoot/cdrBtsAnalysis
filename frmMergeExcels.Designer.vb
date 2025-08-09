<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMergeExcels
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMergeExcels))
        Me.btnMobilink = New System.Windows.Forms.Button()
        Me.btnMerge = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lbMobilinkFileCount = New System.Windows.Forms.Label()
        Me.lbRecCount = New System.Windows.Forms.Label()
        Me.lbMobilink = New System.Windows.Forms.Label()
        Me.prbMbilink = New System.Windows.Forms.ProgressBar()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lbUfoneFileCount = New System.Windows.Forms.Label()
        Me.lbUfoneCount = New System.Windows.Forms.Label()
        Me.lbUfone = New System.Windows.Forms.Label()
        Me.prbUfone = New System.Windows.Forms.ProgressBar()
        Me.btnUfone = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lbZongFileCount = New System.Windows.Forms.Label()
        Me.lbZongCount = New System.Windows.Forms.Label()
        Me.lbZong = New System.Windows.Forms.Label()
        Me.prbZong = New System.Windows.Forms.ProgressBar()
        Me.btnZong = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.lbTeleNorFileCount = New System.Windows.Forms.Label()
        Me.lbTeleNorCount = New System.Windows.Forms.Label()
        Me.lbTeleNor = New System.Windows.Forms.Label()
        Me.prbTeleNor = New System.Windows.Forms.ProgressBar()
        Me.btnTeleNor = New System.Windows.Forms.Button()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnSetting = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.txtFolderName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.lbAllIDsFiles = New System.Windows.Forms.Label()
        Me.lbAllIDsCount = New System.Windows.Forms.Label()
        Me.lbAllIDs = New System.Windows.Forms.Label()
        Me.pbrAllIDs = New System.Windows.Forms.ProgressBar()
        Me.btnAllIDs = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnMobilink
        '
        Me.btnMobilink.Location = New System.Drawing.Point(469, 17)
        Me.btnMobilink.Name = "btnMobilink"
        Me.btnMobilink.Size = New System.Drawing.Size(75, 23)
        Me.btnMobilink.TabIndex = 2
        Me.btnMobilink.Text = "Browse"
        Me.btnMobilink.UseVisualStyleBackColor = True
        '
        'btnMerge
        '
        Me.btnMerge.Location = New System.Drawing.Point(319, 492)
        Me.btnMerge.Name = "btnMerge"
        Me.btnMerge.Size = New System.Drawing.Size(75, 23)
        Me.btnMerge.TabIndex = 12
        Me.btnMerge.Text = "&Merge"
        Me.btnMerge.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(481, 492)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lbMobilinkFileCount)
        Me.GroupBox1.Controls.Add(Me.lbRecCount)
        Me.GroupBox1.Controls.Add(Me.lbMobilink)
        Me.GroupBox1.Controls.Add(Me.prbMbilink)
        Me.GroupBox1.Controls.Add(Me.btnMobilink)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 50)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(566, 80)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Mobilink"
        '
        'lbMobilinkFileCount
        '
        Me.lbMobilinkFileCount.AutoSize = True
        Me.lbMobilinkFileCount.Location = New System.Drawing.Point(6, 56)
        Me.lbMobilinkFileCount.Name = "lbMobilinkFileCount"
        Me.lbMobilinkFileCount.Size = New System.Drawing.Size(39, 13)
        Me.lbMobilinkFileCount.TabIndex = 6
        Me.lbMobilinkFileCount.Text = "Label1"
        Me.lbMobilinkFileCount.Visible = False
        '
        'lbRecCount
        '
        Me.lbRecCount.AutoSize = True
        Me.lbRecCount.Location = New System.Drawing.Point(442, 56)
        Me.lbRecCount.Name = "lbRecCount"
        Me.lbRecCount.Size = New System.Drawing.Size(39, 13)
        Me.lbRecCount.TabIndex = 5
        Me.lbRecCount.Text = "Label1"
        Me.lbRecCount.Visible = False
        '
        'lbMobilink
        '
        Me.lbMobilink.BackColor = System.Drawing.Color.White
        Me.lbMobilink.Location = New System.Drawing.Point(25, 17)
        Me.lbMobilink.Name = "lbMobilink"
        Me.lbMobilink.Size = New System.Drawing.Size(438, 23)
        Me.lbMobilink.TabIndex = 4
        Me.lbMobilink.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'prbMbilink
        '
        Me.prbMbilink.Location = New System.Drawing.Point(85, 51)
        Me.prbMbilink.Name = "prbMbilink"
        Me.prbMbilink.Size = New System.Drawing.Size(340, 23)
        Me.prbMbilink.TabIndex = 2
        Me.prbMbilink.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lbUfoneFileCount)
        Me.GroupBox2.Controls.Add(Me.lbUfoneCount)
        Me.GroupBox2.Controls.Add(Me.lbUfone)
        Me.GroupBox2.Controls.Add(Me.prbUfone)
        Me.GroupBox2.Controls.Add(Me.btnUfone)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 136)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(566, 80)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Ufone"
        '
        'lbUfoneFileCount
        '
        Me.lbUfoneFileCount.AutoSize = True
        Me.lbUfoneFileCount.Location = New System.Drawing.Point(6, 56)
        Me.lbUfoneFileCount.Name = "lbUfoneFileCount"
        Me.lbUfoneFileCount.Size = New System.Drawing.Size(39, 13)
        Me.lbUfoneFileCount.TabIndex = 9
        Me.lbUfoneFileCount.Text = "Label1"
        Me.lbUfoneFileCount.Visible = False
        '
        'lbUfoneCount
        '
        Me.lbUfoneCount.AutoSize = True
        Me.lbUfoneCount.BackColor = System.Drawing.Color.Transparent
        Me.lbUfoneCount.Location = New System.Drawing.Point(442, 56)
        Me.lbUfoneCount.Name = "lbUfoneCount"
        Me.lbUfoneCount.Size = New System.Drawing.Size(39, 13)
        Me.lbUfoneCount.TabIndex = 8
        Me.lbUfoneCount.Text = "Label1"
        Me.lbUfoneCount.Visible = False
        '
        'lbUfone
        '
        Me.lbUfone.BackColor = System.Drawing.Color.White
        Me.lbUfone.Location = New System.Drawing.Point(25, 17)
        Me.lbUfone.Name = "lbUfone"
        Me.lbUfone.Size = New System.Drawing.Size(438, 23)
        Me.lbUfone.TabIndex = 7
        Me.lbUfone.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'prbUfone
        '
        Me.prbUfone.Location = New System.Drawing.Point(85, 51)
        Me.prbUfone.Name = "prbUfone"
        Me.prbUfone.Size = New System.Drawing.Size(340, 23)
        Me.prbUfone.TabIndex = 2
        Me.prbUfone.Visible = False
        '
        'btnUfone
        '
        Me.btnUfone.Location = New System.Drawing.Point(469, 17)
        Me.btnUfone.Name = "btnUfone"
        Me.btnUfone.Size = New System.Drawing.Size(75, 23)
        Me.btnUfone.TabIndex = 5
        Me.btnUfone.Text = "Browse"
        Me.btnUfone.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lbZongFileCount)
        Me.GroupBox3.Controls.Add(Me.lbZongCount)
        Me.GroupBox3.Controls.Add(Me.lbZong)
        Me.GroupBox3.Controls.Add(Me.prbZong)
        Me.GroupBox3.Controls.Add(Me.btnZong)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 222)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(566, 80)
        Me.GroupBox3.TabIndex = 6
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Zong"
        '
        'lbZongFileCount
        '
        Me.lbZongFileCount.AutoSize = True
        Me.lbZongFileCount.Location = New System.Drawing.Point(6, 56)
        Me.lbZongFileCount.Name = "lbZongFileCount"
        Me.lbZongFileCount.Size = New System.Drawing.Size(39, 13)
        Me.lbZongFileCount.TabIndex = 12
        Me.lbZongFileCount.Text = "Label1"
        Me.lbZongFileCount.Visible = False
        '
        'lbZongCount
        '
        Me.lbZongCount.AutoSize = True
        Me.lbZongCount.Location = New System.Drawing.Point(445, 56)
        Me.lbZongCount.Name = "lbZongCount"
        Me.lbZongCount.Size = New System.Drawing.Size(39, 13)
        Me.lbZongCount.TabIndex = 11
        Me.lbZongCount.Text = "Label1"
        Me.lbZongCount.Visible = False
        '
        'lbZong
        '
        Me.lbZong.BackColor = System.Drawing.Color.White
        Me.lbZong.Location = New System.Drawing.Point(25, 17)
        Me.lbZong.Name = "lbZong"
        Me.lbZong.Size = New System.Drawing.Size(438, 23)
        Me.lbZong.TabIndex = 10
        Me.lbZong.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'prbZong
        '
        Me.prbZong.Location = New System.Drawing.Point(85, 51)
        Me.prbZong.Name = "prbZong"
        Me.prbZong.Size = New System.Drawing.Size(340, 23)
        Me.prbZong.TabIndex = 2
        Me.prbZong.Visible = False
        '
        'btnZong
        '
        Me.btnZong.Location = New System.Drawing.Point(469, 17)
        Me.btnZong.Name = "btnZong"
        Me.btnZong.Size = New System.Drawing.Size(75, 23)
        Me.btnZong.TabIndex = 8
        Me.btnZong.Text = "Browse"
        Me.btnZong.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.lbTeleNorFileCount)
        Me.GroupBox4.Controls.Add(Me.lbTeleNorCount)
        Me.GroupBox4.Controls.Add(Me.lbTeleNor)
        Me.GroupBox4.Controls.Add(Me.prbTeleNor)
        Me.GroupBox4.Controls.Add(Me.btnTeleNor)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 306)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(566, 80)
        Me.GroupBox4.TabIndex = 9
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "TeleNor"
        '
        'lbTeleNorFileCount
        '
        Me.lbTeleNorFileCount.AutoSize = True
        Me.lbTeleNorFileCount.Location = New System.Drawing.Point(6, 56)
        Me.lbTeleNorFileCount.Name = "lbTeleNorFileCount"
        Me.lbTeleNorFileCount.Size = New System.Drawing.Size(39, 13)
        Me.lbTeleNorFileCount.TabIndex = 15
        Me.lbTeleNorFileCount.Text = "Label1"
        Me.lbTeleNorFileCount.Visible = False
        '
        'lbTeleNorCount
        '
        Me.lbTeleNorCount.AutoSize = True
        Me.lbTeleNorCount.Location = New System.Drawing.Point(448, 56)
        Me.lbTeleNorCount.Name = "lbTeleNorCount"
        Me.lbTeleNorCount.Size = New System.Drawing.Size(39, 13)
        Me.lbTeleNorCount.TabIndex = 14
        Me.lbTeleNorCount.Text = "Label2"
        Me.lbTeleNorCount.Visible = False
        '
        'lbTeleNor
        '
        Me.lbTeleNor.BackColor = System.Drawing.Color.White
        Me.lbTeleNor.Location = New System.Drawing.Point(25, 17)
        Me.lbTeleNor.Name = "lbTeleNor"
        Me.lbTeleNor.Size = New System.Drawing.Size(438, 23)
        Me.lbTeleNor.TabIndex = 13
        Me.lbTeleNor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'prbTeleNor
        '
        Me.prbTeleNor.Location = New System.Drawing.Point(85, 51)
        Me.prbTeleNor.Name = "prbTeleNor"
        Me.prbTeleNor.Size = New System.Drawing.Size(340, 23)
        Me.prbTeleNor.TabIndex = 2
        Me.prbTeleNor.Visible = False
        '
        'btnTeleNor
        '
        Me.btnTeleNor.Location = New System.Drawing.Point(469, 17)
        Me.btnTeleNor.Name = "btnTeleNor"
        Me.btnTeleNor.Size = New System.Drawing.Size(75, 23)
        Me.btnTeleNor.TabIndex = 11
        Me.btnTeleNor.Text = "Browse"
        Me.btnTeleNor.UseVisualStyleBackColor = True
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(400, 492)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnRefresh.TabIndex = 14
        Me.btnRefresh.Text = "&Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'btnSetting
        '
        Me.btnSetting.AutoSize = True
        Me.btnSetting.Image = CType(resources.GetObject("btnSetting.Image"), System.Drawing.Image)
        Me.btnSetting.Location = New System.Drawing.Point(547, 8)
        Me.btnSetting.Name = "btnSetting"
        Me.btnSetting.Size = New System.Drawing.Size(31, 31)
        Me.btnSetting.TabIndex = 15
        Me.btnSetting.UseVisualStyleBackColor = True
        '
        'txtFolderName
        '
        Me.txtFolderName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFolderName.Location = New System.Drawing.Point(128, 17)
        Me.txtFolderName.Name = "txtFolderName"
        Me.txtFolderName.Size = New System.Drawing.Size(347, 22)
        Me.txtFolderName.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(52, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Folder Name:"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.lbAllIDsFiles)
        Me.GroupBox5.Controls.Add(Me.lbAllIDsCount)
        Me.GroupBox5.Controls.Add(Me.lbAllIDs)
        Me.GroupBox5.Controls.Add(Me.pbrAllIDs)
        Me.GroupBox5.Controls.Add(Me.btnAllIDs)
        Me.GroupBox5.Location = New System.Drawing.Point(12, 392)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(566, 82)
        Me.GroupBox5.TabIndex = 18
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Combined"
        '
        'lbAllIDsFiles
        '
        Me.lbAllIDsFiles.AutoSize = True
        Me.lbAllIDsFiles.Location = New System.Drawing.Point(14, 55)
        Me.lbAllIDsFiles.Name = "lbAllIDsFiles"
        Me.lbAllIDsFiles.Size = New System.Drawing.Size(39, 13)
        Me.lbAllIDsFiles.TabIndex = 20
        Me.lbAllIDsFiles.Text = "Label1"
        Me.lbAllIDsFiles.Visible = False
        '
        'lbAllIDsCount
        '
        Me.lbAllIDsCount.AutoSize = True
        Me.lbAllIDsCount.Location = New System.Drawing.Point(456, 55)
        Me.lbAllIDsCount.Name = "lbAllIDsCount"
        Me.lbAllIDsCount.Size = New System.Drawing.Size(39, 13)
        Me.lbAllIDsCount.TabIndex = 19
        Me.lbAllIDsCount.Text = "Label2"
        Me.lbAllIDsCount.Visible = False
        '
        'lbAllIDs
        '
        Me.lbAllIDs.BackColor = System.Drawing.Color.White
        Me.lbAllIDs.Location = New System.Drawing.Point(33, 16)
        Me.lbAllIDs.Name = "lbAllIDs"
        Me.lbAllIDs.Size = New System.Drawing.Size(438, 23)
        Me.lbAllIDs.TabIndex = 18
        Me.lbAllIDs.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pbrAllIDs
        '
        Me.pbrAllIDs.Location = New System.Drawing.Point(93, 50)
        Me.pbrAllIDs.Name = "pbrAllIDs"
        Me.pbrAllIDs.Size = New System.Drawing.Size(340, 23)
        Me.pbrAllIDs.TabIndex = 16
        Me.pbrAllIDs.Visible = False
        '
        'btnAllIDs
        '
        Me.btnAllIDs.Location = New System.Drawing.Point(477, 16)
        Me.btnAllIDs.Name = "btnAllIDs"
        Me.btnAllIDs.Size = New System.Drawing.Size(75, 23)
        Me.btnAllIDs.TabIndex = 17
        Me.btnAllIDs.Text = "Browse"
        Me.btnAllIDs.UseVisualStyleBackColor = True
        '
        'frmMergeExcels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(590, 527)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFolderName)
        Me.Controls.Add(Me.btnSetting)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnMerge)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMergeExcels"
        Me.Text = "Merge Excels"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnMobilink As System.Windows.Forms.Button
    Friend WithEvents btnMerge As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents prbMbilink As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents prbUfone As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUfone As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents prbZong As System.Windows.Forms.ProgressBar
    Friend WithEvents btnZong As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents prbTeleNor As System.Windows.Forms.ProgressBar
    Friend WithEvents btnTeleNor As System.Windows.Forms.Button
    Friend WithEvents lbUfone As System.Windows.Forms.Label
    Friend WithEvents lbZong As System.Windows.Forms.Label
    Friend WithEvents lbTeleNor As System.Windows.Forms.Label
    Friend WithEvents lbRecCount As System.Windows.Forms.Label
    Friend WithEvents lbUfoneCount As System.Windows.Forms.Label
    Friend WithEvents lbMobilink As System.Windows.Forms.Label
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents lbZongCount As System.Windows.Forms.Label
    Friend WithEvents lbTeleNorCount As System.Windows.Forms.Label
    Friend WithEvents lbMobilinkFileCount As System.Windows.Forms.Label
    Friend WithEvents lbUfoneFileCount As System.Windows.Forms.Label
    Friend WithEvents lbZongFileCount As System.Windows.Forms.Label
    Friend WithEvents lbTeleNorFileCount As System.Windows.Forms.Label
    Friend WithEvents btnSetting As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents txtFolderName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents lbAllIDsFiles As System.Windows.Forms.Label
    Friend WithEvents lbAllIDsCount As System.Windows.Forms.Label
    Friend WithEvents lbAllIDs As System.Windows.Forms.Label
    Friend WithEvents pbrAllIDs As System.Windows.Forms.ProgressBar
    Friend WithEvents btnAllIDs As System.Windows.Forms.Button
End Class
