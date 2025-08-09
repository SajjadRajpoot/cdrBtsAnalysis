<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.btn_CropBTS = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lbfilepath = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lbCheckedNos = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lbFound = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lbTotalNos = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkLimitedActivity = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lbCheckedIMEI = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lbFoundIMEIs = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lbTotalIMEIs = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btnIMEIsComparison = New System.Windows.Forms.Button()
        Me.btnBrowseBTSs = New System.Windows.Forms.Button()
        Me.lbIMEIsNames = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.cbkWithVisio = New System.Windows.Forms.CheckBox()
        Me.lbtotalCallsa = New System.Windows.Forms.Label()
        Me.lbCheckedaPraty = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lbNoOfAnagram = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lbTotalNoOfaParty = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.btnBuildAnagram = New System.Windows.Forms.Button()
        Me.btnBrowseGroupSheet = New System.Windows.Forms.Button()
        Me.lbGroupSheetPath = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lbBT_CRP_Checked = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cbFind_b_in_a = New System.Windows.Forms.CheckBox()
        Me.lbNoOfRecords = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_BrowsBTS = New System.Windows.Forms.Button()
        Me.txt_BTS_Path = New System.Windows.Forms.TextBox()
        Me.dtp_Ending_Time = New System.Windows.Forms.DateTimePicker()
        Me.dtp_Initial_Time = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkCNIC = New System.Windows.Forms.CheckBox()
        Me.cbProblemFile = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_CropBTS
        '
        Me.btn_CropBTS.Location = New System.Drawing.Point(135, 123)
        Me.btn_CropBTS.Name = "btn_CropBTS"
        Me.btn_CropBTS.Size = New System.Drawing.Size(75, 23)
        Me.btn_CropBTS.TabIndex = 1
        Me.btn_CropBTS.Text = "Analyze"
        Me.btn_CropBTS.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(135, 506)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lbfilepath
        '
        Me.lbfilepath.BackColor = System.Drawing.Color.White
        Me.lbfilepath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbfilepath.Location = New System.Drawing.Point(6, 9)
        Me.lbfilepath.Name = "lbfilepath"
        Me.lbfilepath.Size = New System.Drawing.Size(261, 19)
        Me.lbfilepath.TabIndex = 0
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(271, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Browse"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(127, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Checked:"
        '
        'lbCheckedNos
        '
        Me.lbCheckedNos.AutoSize = True
        Me.lbCheckedNos.Location = New System.Drawing.Point(177, 42)
        Me.lbCheckedNos.Name = "lbCheckedNos"
        Me.lbCheckedNos.Size = New System.Drawing.Size(13, 13)
        Me.lbCheckedNos.TabIndex = 14
        Me.lbCheckedNos.Text = "0"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(259, 41)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Found:"
        '
        'lbFound
        '
        Me.lbFound.AutoSize = True
        Me.lbFound.Location = New System.Drawing.Point(298, 41)
        Me.lbFound.Name = "lbFound"
        Me.lbFound.Size = New System.Drawing.Size(13, 13)
        Me.lbFound.TabIndex = 16
        Me.lbFound.Text = "0"
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(140, 83)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 17
        Me.Button3.Text = "Find Groups"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lbTotalNos)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.lbFound)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lbCheckedNos)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.lbfilepath)
        Me.Panel1.Location = New System.Drawing.Point(6, 174)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(348, 68)
        Me.Panel1.TabIndex = 11
        '
        'lbTotalNos
        '
        Me.lbTotalNos.AutoSize = True
        Me.lbTotalNos.Location = New System.Drawing.Point(43, 41)
        Me.lbTotalNos.Name = "lbTotalNos"
        Me.lbTotalNos.Size = New System.Drawing.Size(13, 13)
        Me.lbTotalNos.TabIndex = 23
        Me.lbTotalNos.Text = "0"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 41)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(43, 13)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Total  : "
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkLimitedActivity)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(3, 158)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(354, 113)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Groups"
        '
        'chkLimitedActivity
        '
        Me.chkLimitedActivity.AutoSize = True
        Me.chkLimitedActivity.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLimitedActivity.Location = New System.Drawing.Point(7, 84)
        Me.chkLimitedActivity.Name = "chkLimitedActivity"
        Me.chkLimitedActivity.Size = New System.Drawing.Size(127, 17)
        Me.chkLimitedActivity.TabIndex = 18
        Me.chkLimitedActivity.Text = "Create LimitedActivity"
        Me.chkLimitedActivity.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lbCheckedIMEI)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.lbFoundIMEIs)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.lbTotalIMEIs)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.btnIMEIsComparison)
        Me.GroupBox3.Controls.Add(Me.btnBrowseBTSs)
        Me.GroupBox3.Controls.Add(Me.lbIMEIsNames)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(3, 386)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(354, 114)
        Me.GroupBox3.TabIndex = 14
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "BTS IMEI Comparison"
        '
        'lbCheckedIMEI
        '
        Me.lbCheckedIMEI.AutoSize = True
        Me.lbCheckedIMEI.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbCheckedIMEI.Location = New System.Drawing.Point(180, 54)
        Me.lbCheckedIMEI.Name = "lbCheckedIMEI"
        Me.lbCheckedIMEI.Size = New System.Drawing.Size(13, 13)
        Me.lbCheckedIMEI.TabIndex = 26
        Me.lbCheckedIMEI.Text = "0"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(130, 54)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 13)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Checked:"
        '
        'lbFoundIMEIs
        '
        Me.lbFoundIMEIs.AutoSize = True
        Me.lbFoundIMEIs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbFoundIMEIs.Location = New System.Drawing.Point(301, 54)
        Me.lbFoundIMEIs.Name = "lbFoundIMEIs"
        Me.lbFoundIMEIs.Size = New System.Drawing.Size(13, 13)
        Me.lbFoundIMEIs.TabIndex = 23
        Me.lbFoundIMEIs.Text = "0"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(262, 54)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 13)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Found:"
        '
        'lbTotalIMEIs
        '
        Me.lbTotalIMEIs.AutoSize = True
        Me.lbTotalIMEIs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbTotalIMEIs.Location = New System.Drawing.Point(46, 55)
        Me.lbTotalIMEIs.Name = "lbTotalIMEIs"
        Me.lbTotalIMEIs.Size = New System.Drawing.Size(13, 13)
        Me.lbTotalIMEIs.TabIndex = 21
        Me.lbTotalIMEIs.Text = "0"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 55)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 13)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "Total  : "
        '
        'btnIMEIsComparison
        '
        Me.btnIMEIsComparison.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnIMEIsComparison.Location = New System.Drawing.Point(132, 82)
        Me.btnIMEIsComparison.Name = "btnIMEIsComparison"
        Me.btnIMEIsComparison.Size = New System.Drawing.Size(75, 23)
        Me.btnIMEIsComparison.TabIndex = 24
        Me.btnIMEIsComparison.Text = "Comparison"
        Me.btnIMEIsComparison.UseVisualStyleBackColor = True
        '
        'btnBrowseBTSs
        '
        Me.btnBrowseBTSs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowseBTSs.Location = New System.Drawing.Point(274, 19)
        Me.btnBrowseBTSs.Name = "btnBrowseBTSs"
        Me.btnBrowseBTSs.Size = New System.Drawing.Size(75, 23)
        Me.btnBrowseBTSs.TabIndex = 19
        Me.btnBrowseBTSs.Text = "Browse"
        Me.btnBrowseBTSs.UseVisualStyleBackColor = True
        '
        'lbIMEIsNames
        '
        Me.lbIMEIsNames.BackColor = System.Drawing.Color.White
        Me.lbIMEIsNames.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbIMEIsNames.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbIMEIsNames.Location = New System.Drawing.Point(9, 21)
        Me.lbIMEIsNames.Name = "lbIMEIsNames"
        Me.lbIMEIsNames.Size = New System.Drawing.Size(261, 19)
        Me.lbIMEIsNames.TabIndex = 18
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cbkWithVisio)
        Me.GroupBox4.Controls.Add(Me.lbtotalCallsa)
        Me.GroupBox4.Controls.Add(Me.lbCheckedaPraty)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Controls.Add(Me.lbNoOfAnagram)
        Me.GroupBox4.Controls.Add(Me.Label13)
        Me.GroupBox4.Controls.Add(Me.lbTotalNoOfaParty)
        Me.GroupBox4.Controls.Add(Me.Label15)
        Me.GroupBox4.Controls.Add(Me.btnBuildAnagram)
        Me.GroupBox4.Controls.Add(Me.btnBrowseGroupSheet)
        Me.GroupBox4.Controls.Add(Me.lbGroupSheetPath)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(3, 274)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(354, 107)
        Me.GroupBox4.TabIndex = 27
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Group Anagram"
        '
        'cbkWithVisio
        '
        Me.cbkWithVisio.AutoSize = True
        Me.cbkWithVisio.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbkWithVisio.Location = New System.Drawing.Point(9, 80)
        Me.cbkWithVisio.Name = "cbkWithVisio"
        Me.cbkWithVisio.Size = New System.Drawing.Size(73, 17)
        Me.cbkWithVisio.TabIndex = 27
        Me.cbkWithVisio.Text = "With Visio"
        Me.cbkWithVisio.UseVisualStyleBackColor = True
        '
        'lbtotalCallsa
        '
        Me.lbtotalCallsa.AutoSize = True
        Me.lbtotalCallsa.Location = New System.Drawing.Point(213, 81)
        Me.lbtotalCallsa.Name = "lbtotalCallsa"
        Me.lbtotalCallsa.Size = New System.Drawing.Size(14, 13)
        Me.lbtotalCallsa.TabIndex = 27
        Me.lbtotalCallsa.Text = "0"
        Me.lbtotalCallsa.Visible = False
        '
        'lbCheckedaPraty
        '
        Me.lbCheckedaPraty.AutoSize = True
        Me.lbCheckedaPraty.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbCheckedaPraty.Location = New System.Drawing.Point(180, 53)
        Me.lbCheckedaPraty.Name = "lbCheckedaPraty"
        Me.lbCheckedaPraty.Size = New System.Drawing.Size(13, 13)
        Me.lbCheckedaPraty.TabIndex = 26
        Me.lbCheckedaPraty.Text = "0"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(130, 53)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(53, 13)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Checked:"
        '
        'lbNoOfAnagram
        '
        Me.lbNoOfAnagram.AutoSize = True
        Me.lbNoOfAnagram.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbNoOfAnagram.Location = New System.Drawing.Point(301, 53)
        Me.lbNoOfAnagram.Name = "lbNoOfAnagram"
        Me.lbNoOfAnagram.Size = New System.Drawing.Size(13, 13)
        Me.lbNoOfAnagram.TabIndex = 23
        Me.lbNoOfAnagram.Text = "0"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(262, 53)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 13)
        Me.Label13.TabIndex = 22
        Me.Label13.Text = "Found:"
        '
        'lbTotalNoOfaParty
        '
        Me.lbTotalNoOfaParty.AutoSize = True
        Me.lbTotalNoOfaParty.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbTotalNoOfaParty.Location = New System.Drawing.Point(46, 54)
        Me.lbTotalNoOfaParty.Name = "lbTotalNoOfaParty"
        Me.lbTotalNoOfaParty.Size = New System.Drawing.Size(13, 13)
        Me.lbTotalNoOfaParty.TabIndex = 21
        Me.lbTotalNoOfaParty.Text = "0"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(6, 54)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(43, 13)
        Me.Label15.TabIndex = 20
        Me.Label15.Text = "Total  : "
        '
        'btnBuildAnagram
        '
        Me.btnBuildAnagram.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBuildAnagram.Location = New System.Drawing.Point(132, 76)
        Me.btnBuildAnagram.Name = "btnBuildAnagram"
        Me.btnBuildAnagram.Size = New System.Drawing.Size(75, 23)
        Me.btnBuildAnagram.TabIndex = 24
        Me.btnBuildAnagram.Text = "Anagram"
        Me.btnBuildAnagram.UseVisualStyleBackColor = True
        '
        'btnBrowseGroupSheet
        '
        Me.btnBrowseGroupSheet.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowseGroupSheet.Location = New System.Drawing.Point(274, 18)
        Me.btnBrowseGroupSheet.Name = "btnBrowseGroupSheet"
        Me.btnBrowseGroupSheet.Size = New System.Drawing.Size(75, 23)
        Me.btnBrowseGroupSheet.TabIndex = 19
        Me.btnBrowseGroupSheet.Text = "Browse"
        Me.btnBrowseGroupSheet.UseVisualStyleBackColor = True
        '
        'lbGroupSheetPath
        '
        Me.lbGroupSheetPath.BackColor = System.Drawing.Color.White
        Me.lbGroupSheetPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbGroupSheetPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbGroupSheetPath.Location = New System.Drawing.Point(9, 20)
        Me.lbGroupSheetPath.Name = "lbGroupSheetPath"
        Me.lbGroupSheetPath.Size = New System.Drawing.Size(261, 19)
        Me.lbGroupSheetPath.TabIndex = 18
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.lbBT_CRP_Checked)
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Controls.Add(Me.cbFind_b_in_a)
        Me.Panel2.Controls.Add(Me.lbNoOfRecords)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.btn_BrowsBTS)
        Me.Panel2.Controls.Add(Me.txt_BTS_Path)
        Me.Panel2.Controls.Add(Me.dtp_Ending_Time)
        Me.Panel2.Controls.Add(Me.dtp_Initial_Time)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Location = New System.Drawing.Point(3, 6)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(354, 146)
        Me.Panel2.TabIndex = 12
        '
        'lbBT_CRP_Checked
        '
        Me.lbBT_CRP_Checked.AutoSize = True
        Me.lbBT_CRP_Checked.Location = New System.Drawing.Point(57, 122)
        Me.lbBT_CRP_Checked.Name = "lbBT_CRP_Checked"
        Me.lbBT_CRP_Checked.Size = New System.Drawing.Size(13, 13)
        Me.lbBT_CRP_Checked.TabIndex = 15
        Me.lbBT_CRP_Checked.Text = "0"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(5, 122)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(46, 13)
        Me.Label12.TabIndex = 14
        Me.Label12.Text = "Croping:"
        '
        'cbFind_b_in_a
        '
        Me.cbFind_b_in_a.AutoSize = True
        Me.cbFind_b_in_a.Location = New System.Drawing.Point(262, 92)
        Me.cbFind_b_in_a.Name = "cbFind_b_in_a"
        Me.cbFind_b_in_a.Size = New System.Drawing.Size(83, 17)
        Me.cbFind_b_in_a.TabIndex = 10
        Me.cbFind_b_in_a.Text = "Find Groups"
        Me.cbFind_b_in_a.UseVisualStyleBackColor = True
        '
        'lbNoOfRecords
        '
        Me.lbNoOfRecords.AutoSize = True
        Me.lbNoOfRecords.Location = New System.Drawing.Point(46, 92)
        Me.lbNoOfRecords.Name = "lbNoOfRecords"
        Me.lbNoOfRecords.Size = New System.Drawing.Size(13, 13)
        Me.lbNoOfRecords.TabIndex = 9
        Me.lbNoOfRecords.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Found:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(180, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "End Date and Time"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Start Date and Time"
        '
        'btn_BrowsBTS
        '
        Me.btn_BrowsBTS.Location = New System.Drawing.Point(273, 14)
        Me.btn_BrowsBTS.Name = "btn_BrowsBTS"
        Me.btn_BrowsBTS.Size = New System.Drawing.Size(75, 23)
        Me.btn_BrowsBTS.TabIndex = 4
        Me.btn_BrowsBTS.Text = "Browse"
        Me.btn_BrowsBTS.UseVisualStyleBackColor = True
        '
        'txt_BTS_Path
        '
        Me.txt_BTS_Path.Location = New System.Drawing.Point(6, 16)
        Me.txt_BTS_Path.Name = "txt_BTS_Path"
        Me.txt_BTS_Path.Size = New System.Drawing.Size(261, 20)
        Me.txt_BTS_Path.TabIndex = 3
        '
        'dtp_Ending_Time
        '
        Me.dtp_Ending_Time.CustomFormat = "dd/MM/yyyy  h:mm:ss tt"
        Me.dtp_Ending_Time.DropDownAlign = System.Windows.Forms.LeftRightAlignment.Right
        Me.dtp_Ending_Time.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_Ending_Time.Location = New System.Drawing.Point(183, 57)
        Me.dtp_Ending_Time.Name = "dtp_Ending_Time"
        Me.dtp_Ending_Time.Size = New System.Drawing.Size(162, 20)
        Me.dtp_Ending_Time.TabIndex = 2
        '
        'dtp_Initial_Time
        '
        Me.dtp_Initial_Time.CustomFormat = "dd/MM/yyyy  h:mm:ss tt"
        Me.dtp_Initial_Time.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_Initial_Time.Location = New System.Drawing.Point(6, 57)
        Me.dtp_Initial_Time.Name = "dtp_Initial_Time"
        Me.dtp_Initial_Time.Size = New System.Drawing.Size(162, 20)
        Me.dtp_Initial_Time.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbProblemFile)
        Me.GroupBox1.Controls.Add(Me.chkCNIC)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(0, -1)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(354, 147)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Crop BTS"
        '
        'chkCNIC
        '
        Me.chkCNIC.AutoSize = True
        Me.chkCNIC.Checked = True
        Me.chkCNIC.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCNIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCNIC.Location = New System.Drawing.Point(291, 119)
        Me.chkCNIC.Name = "chkCNIC"
        Me.chkCNIC.Size = New System.Drawing.Size(51, 17)
        Me.chkCNIC.TabIndex = 14
        Me.chkCNIC.Text = "CNIC"
        Me.chkCNIC.UseVisualStyleBackColor = True
        Me.chkCNIC.Visible = False
        '
        'cbProblemFile
        '
        Me.cbProblemFile.AutoSize = True
        Me.cbProblemFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbProblemFile.Location = New System.Drawing.Point(261, 118)
        Me.cbProblemFile.Name = "cbProblemFile"
        Me.cbProblemFile.Size = New System.Drawing.Size(83, 17)
        Me.cbProblemFile.TabIndex = 15
        Me.cbProblemFile.Text = "Problem File"
        Me.cbProblemFile.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(360, 536)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btn_CropBTS)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BTS Analysis"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btn_CropBTS As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lbfilepath As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lbCheckedNos As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lbFound As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lbFoundIMEIs As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lbTotalIMEIs As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnIMEIsComparison As System.Windows.Forms.Button
    Friend WithEvents btnBrowseBTSs As System.Windows.Forms.Button
    Friend WithEvents lbIMEIsNames As System.Windows.Forms.Label
    Friend WithEvents lbCheckedIMEI As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lbTotalNos As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents lbCheckedaPraty As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lbNoOfAnagram As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lbTotalNoOfaParty As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents btnBuildAnagram As System.Windows.Forms.Button
    Friend WithEvents btnBrowseGroupSheet As System.Windows.Forms.Button
    Friend WithEvents lbGroupSheetPath As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cbFind_b_in_a As System.Windows.Forms.CheckBox
    Friend WithEvents lbNoOfRecords As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_BrowsBTS As System.Windows.Forms.Button
    Friend WithEvents txt_BTS_Path As System.Windows.Forms.TextBox
    Friend WithEvents dtp_Ending_Time As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtp_Initial_Time As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkCNIC As System.Windows.Forms.CheckBox
    Friend WithEvents chkLimitedActivity As System.Windows.Forms.CheckBox
    Friend WithEvents lbtotalCallsa As System.Windows.Forms.Label
    Friend WithEvents cbkWithVisio As System.Windows.Forms.CheckBox
    Friend WithEvents lbBT_CRP_Checked As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cbProblemFile As System.Windows.Forms.CheckBox

End Class
