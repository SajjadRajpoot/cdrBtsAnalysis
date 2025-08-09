<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVerisys
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
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtVerisysPath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnAddNew = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.lbInsertedVerisys = New System.Windows.Forms.Label()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.pbVerisys = New System.Windows.Forms.ProgressBar()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtTargetPath = New System.Windows.Forms.TextBox()
        Me.btnTargetFolder = New System.Windows.Forms.Button()
        Me.cbSameAddress = New System.Windows.Forms.CheckBox()
        Me.lbSelectedFiles = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(573, 57)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(75, 23)
        Me.btnBrowse.TabIndex = 0
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtVerisysPath
        '
        Me.txtVerisysPath.Location = New System.Drawing.Point(56, 57)
        Me.txtVerisysPath.Name = "txtVerisysPath"
        Me.txtVerisysPath.Size = New System.Drawing.Size(511, 20)
        Me.txtVerisysPath.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(56, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 18)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Current Path:"
        '
        'btnAddNew
        '
        Me.btnAddNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNew.Location = New System.Drawing.Point(411, 375)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(75, 23)
        Me.btnAddNew.TabIndex = 3
        Me.btnAddNew.Text = "&Save"
        Me.btnAddNew.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(573, 375)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'lbInsertedVerisys
        '
        Me.lbInsertedVerisys.AutoSize = True
        Me.lbInsertedVerisys.Location = New System.Drawing.Point(56, 250)
        Me.lbInsertedVerisys.Name = "lbInsertedVerisys"
        Me.lbInsertedVerisys.Size = New System.Drawing.Size(39, 13)
        Me.lbInsertedVerisys.TabIndex = 8
        Me.lbInsertedVerisys.Text = "Label5"
        Me.lbInsertedVerisys.Visible = False
        '
        'btnUpdate
        '
        Me.btnUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.Location = New System.Drawing.Point(492, 375)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 9
        Me.btnUpdate.Text = "&Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'pbVerisys
        '
        Me.pbVerisys.Location = New System.Drawing.Point(56, 266)
        Me.pbVerisys.Name = "pbVerisys"
        Me.pbVerisys.Size = New System.Drawing.Size(592, 15)
        Me.pbVerisys.TabIndex = 10
        Me.pbVerisys.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(56, 105)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 18)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Target Path:"
        '
        'txtTargetPath
        '
        Me.txtTargetPath.Location = New System.Drawing.Point(56, 125)
        Me.txtTargetPath.Name = "txtTargetPath"
        Me.txtTargetPath.Size = New System.Drawing.Size(511, 20)
        Me.txtTargetPath.TabIndex = 11
        '
        'btnTargetFolder
        '
        Me.btnTargetFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTargetFolder.Location = New System.Drawing.Point(573, 122)
        Me.btnTargetFolder.Name = "btnTargetFolder"
        Me.btnTargetFolder.Size = New System.Drawing.Size(75, 23)
        Me.btnTargetFolder.TabIndex = 13
        Me.btnTargetFolder.Text = "Browse"
        Me.btnTargetFolder.UseVisualStyleBackColor = True
        '
        'cbSameAddress
        '
        Me.cbSameAddress.AutoSize = True
        Me.cbSameAddress.Location = New System.Drawing.Point(573, 95)
        Me.cbSameAddress.Name = "cbSameAddress"
        Me.cbSameAddress.Size = New System.Drawing.Size(78, 17)
        Me.cbSameAddress.TabIndex = 14
        Me.cbSameAddress.Text = "Both Same"
        Me.cbSameAddress.UseVisualStyleBackColor = True
        '
        'lbSelectedFiles
        '
        Me.lbSelectedFiles.AutoSize = True
        Me.lbSelectedFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbSelectedFiles.Location = New System.Drawing.Point(284, 181)
        Me.lbSelectedFiles.Name = "lbSelectedFiles"
        Me.lbSelectedFiles.Size = New System.Drawing.Size(28, 18)
        Me.lbSelectedFiles.TabIndex = 15
        Me.lbSelectedFiles.Text = "----"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(174, 181)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 18)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Selected Files:"
        '
        'frmVerisys
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(705, 439)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lbSelectedFiles)
        Me.Controls.Add(Me.cbSameAddress)
        Me.Controls.Add(Me.btnTargetFolder)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtTargetPath)
        Me.Controls.Add(Me.pbVerisys)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.lbInsertedVerisys)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtVerisysPath)
        Me.Controls.Add(Me.btnBrowse)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmVerisys"
        Me.Text = "Verisys"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtVerisysPath As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lbInsertedVerisys As System.Windows.Forms.Label
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents pbVerisys As System.Windows.Forms.ProgressBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtTargetPath As System.Windows.Forms.TextBox
    Friend WithEvents btnTargetFolder As System.Windows.Forms.Button
    Friend WithEvents cbSameAddress As System.Windows.Forms.CheckBox
    Friend WithEvents lbSelectedFiles As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
