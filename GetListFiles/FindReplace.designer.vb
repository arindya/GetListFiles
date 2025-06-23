<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FindReplace
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FindReplace))
        Me.btnProcess = New System.Windows.Forms.Button
        Me.txFind = New System.Windows.Forms.TextBox
        Me.txReplace = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.txBrowse = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.btnReplace = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.lbl_record = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.cbFindMod = New System.Windows.Forms.CheckBox
        Me.cbCaseStv = New System.Windows.Forms.CheckBox
        Me.MS = New System.Windows.Forms.MenuStrip
        Me.GetListFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.GetListFilesToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.FindReplaceFillesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RenameFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MS.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnProcess
        '
        Me.btnProcess.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProcess.Location = New System.Drawing.Point(333, 107)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(111, 26)
        Me.btnProcess.TabIndex = 0
        Me.btnProcess.Text = "Find"
        Me.btnProcess.UseVisualStyleBackColor = True
        '
        'txFind
        '
        Me.txFind.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txFind.Location = New System.Drawing.Point(158, 82)
        Me.txFind.Name = "txFind"
        Me.txFind.Size = New System.Drawing.Size(168, 22)
        Me.txFind.TabIndex = 1
        '
        'txReplace
        '
        Me.txReplace.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txReplace.Location = New System.Drawing.Point(158, 111)
        Me.txReplace.Name = "txReplace"
        Me.txReplace.Size = New System.Drawing.Size(168, 22)
        Me.txReplace.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(12, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 22)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Find File Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(12, 108)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(108, 22)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Replace With"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(137, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(14, 22)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = ":"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(137, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(14, 22)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = ":"
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(450, 55)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(111, 24)
        Me.btnBrowse.TabIndex = 7
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txBrowse
        '
        Me.txBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txBrowse.Location = New System.Drawing.Point(123, 55)
        Me.txBrowse.Name = "txBrowse"
        Me.txBrowse.Size = New System.Drawing.Size(321, 22)
        Me.txBrowse.TabIndex = 8
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(12, 177)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(549, 328)
        Me.TextBox1.TabIndex = 9
        '
        'btnReplace
        '
        Me.btnReplace.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReplace.Location = New System.Drawing.Point(450, 109)
        Me.btnReplace.Name = "btnReplace"
        Me.btnReplace.Size = New System.Drawing.Size(111, 26)
        Me.btnReplace.TabIndex = 10
        Me.btnReplace.Text = "Replace"
        Me.btnReplace.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(102, 52)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(14, 22)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = ":"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(12, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(81, 22)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Directory"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 138)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(405, 23)
        Me.ProgressBar1.TabIndex = 13
        '
        'lbl_record
        '
        Me.lbl_record.AutoSize = True
        Me.lbl_record.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lbl_record.Location = New System.Drawing.Point(423, 139)
        Me.lbl_record.Name = "lbl_record"
        Me.lbl_record.Size = New System.Drawing.Size(87, 22)
        Me.lbl_record.TabIndex = 14
        Me.lbl_record.Text = "0 of 0 Files"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label7.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(456, 484)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 17)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "by armayndo"
        '
        'cbFindMod
        '
        Me.cbFindMod.AutoSize = True
        Me.cbFindMod.Location = New System.Drawing.Point(16, 32)
        Me.cbFindMod.Name = "cbFindMod"
        Me.cbFindMod.Size = New System.Drawing.Size(120, 17)
        Me.cbFindMod.TabIndex = 16
        Me.cbFindMod.Text = "Include Sub Folders"
        Me.cbFindMod.UseVisualStyleBackColor = True
        '
        'cbCaseStv
        '
        Me.cbCaseStv.AutoSize = True
        Me.cbCaseStv.Location = New System.Drawing.Point(132, 32)
        Me.cbCaseStv.Name = "cbCaseStv"
        Me.cbCaseStv.Size = New System.Drawing.Size(139, 17)
        Me.cbCaseStv.TabIndex = 17
        Me.cbCaseStv.Text = "Case Sensitive Replace"
        Me.cbCaseStv.UseVisualStyleBackColor = True
        '
        'MS
        '
        Me.MS.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GetListFilesToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MS.Location = New System.Drawing.Point(0, 0)
        Me.MS.Name = "MS"
        Me.MS.Size = New System.Drawing.Size(573, 24)
        Me.MS.TabIndex = 41
        Me.MS.Text = "MenuStrip1"
        '
        'GetListFilesToolStripMenuItem
        '
        Me.GetListFilesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GetListFilesToolStripMenuItem1, Me.FindReplaceFillesToolStripMenuItem, Me.RenameFilesToolStripMenuItem})
        Me.GetListFilesToolStripMenuItem.Name = "GetListFilesToolStripMenuItem"
        Me.GetListFilesToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.GetListFilesToolStripMenuItem.Text = "List Apps"
        '
        'GetListFilesToolStripMenuItem1
        '
        Me.GetListFilesToolStripMenuItem1.Name = "GetListFilesToolStripMenuItem1"
        Me.GetListFilesToolStripMenuItem1.Size = New System.Drawing.Size(167, 22)
        Me.GetListFilesToolStripMenuItem1.Text = "Get List Files"
        '
        'FindReplaceFillesToolStripMenuItem
        '
        Me.FindReplaceFillesToolStripMenuItem.Name = "FindReplaceFillesToolStripMenuItem"
        Me.FindReplaceFillesToolStripMenuItem.Size = New System.Drawing.Size(167, 22)
        Me.FindReplaceFillesToolStripMenuItem.Text = "FindReplace Filles"
        '
        'RenameFilesToolStripMenuItem
        '
        Me.RenameFilesToolStripMenuItem.Name = "RenameFilesToolStripMenuItem"
        Me.RenameFilesToolStripMenuItem.Size = New System.Drawing.Size(167, 22)
        Me.RenameFilesToolStripMenuItem.Text = "Rename Files"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'FindReplace
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(573, 520)
        Me.Controls.Add(Me.MS)
        Me.Controls.Add(Me.cbCaseStv)
        Me.Controls.Add(Me.cbFindMod)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lbl_record)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnReplace)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.txBrowse)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txReplace)
        Me.Controls.Add(Me.txFind)
        Me.Controls.Add(Me.btnProcess)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(579, 549)
        Me.MinimumSize = New System.Drawing.Size(579, 549)
        Me.Name = "FindReplace"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FindReplace"
        Me.MS.ResumeLayout(False)
        Me.MS.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnProcess As System.Windows.Forms.Button
    Friend WithEvents txFind As System.Windows.Forms.TextBox
    Friend WithEvents txReplace As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txBrowse As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents btnReplace As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lbl_record As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbFindMod As System.Windows.Forms.CheckBox
    Friend WithEvents cbCaseStv As System.Windows.Forms.CheckBox
    Friend WithEvents MS As System.Windows.Forms.MenuStrip
    Friend WithEvents GetListFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GetListFilesToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FindReplaceFillesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RenameFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
