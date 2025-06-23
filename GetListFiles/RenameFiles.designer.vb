<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RenameFiles
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RenameFiles))
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.txBrowse = New System.Windows.Forms.TextBox
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Label1 = New System.Windows.Forms.Label
        Me.txRange = New System.Windows.Forms.TextBox
        Me.btnProses = New System.Windows.Forms.Button
        Me.lbl_record = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Label7 = New System.Windows.Forms.Label
        Me.txSheet = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.txFolder = New System.Windows.Forms.TextBox
        Me.btnBrowse2 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.cbDelOri = New System.Windows.Forms.CheckBox
        Me.cbData = New System.Windows.Forms.CheckBox
        Me.MS = New System.Windows.Forms.MenuStrip
        Me.GetListFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.GetListFilesToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.FindReplaceFillesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RenameFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MS.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.btnBrowse.Location = New System.Drawing.Point(418, 57)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(75, 22)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Load"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txBrowse
        '
        Me.txBrowse.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.txBrowse.Location = New System.Drawing.Point(12, 57)
        Me.txBrowse.Name = "txBrowse"
        Me.txBrowse.Size = New System.Drawing.Size(400, 22)
        Me.txBrowse.TabIndex = 1
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(147, 90)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Range Data :"
        '
        'txRange
        '
        Me.txRange.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.txRange.Location = New System.Drawing.Point(257, 87)
        Me.txRange.Name = "txRange"
        Me.txRange.Size = New System.Drawing.Size(72, 22)
        Me.txRange.TabIndex = 4
        '
        'btnProses
        '
        Me.btnProses.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.btnProses.Location = New System.Drawing.Point(418, 145)
        Me.btnProses.Name = "btnProses"
        Me.btnProses.Size = New System.Drawing.Size(75, 23)
        Me.btnProses.TabIndex = 6
        Me.btnProses.Text = "Process"
        Me.btnProses.UseVisualStyleBackColor = True
        '
        'lbl_record
        '
        Me.lbl_record.AutoSize = True
        Me.lbl_record.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.lbl_record.Location = New System.Drawing.Point(12, 142)
        Me.lbl_record.Name = "lbl_record"
        Me.lbl_record.Size = New System.Drawing.Size(104, 16)
        Me.lbl_record.TabIndex = 16
        Me.lbl_record.Text = "0 of 0 Files"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 116)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(481, 23)
        Me.ProgressBar1.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label7.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(412, 457)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 17)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "by armayndo"
        '
        'txSheet
        '
        Me.txSheet.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.txSheet.Location = New System.Drawing.Point(78, 87)
        Me.txSheet.Name = "txSheet"
        Me.txSheet.Size = New System.Drawing.Size(68, 22)
        Me.txSheet.TabIndex = 3
        Me.txSheet.Text = "Report"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 90)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Sheet :"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.Button1.Location = New System.Drawing.Point(337, 145)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "View"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 174)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(481, 280)
        Me.DataGridView1.TabIndex = 20
        '
        'txFolder
        '
        Me.txFolder.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.txFolder.Location = New System.Drawing.Point(90, 29)
        Me.txFolder.Name = "txFolder"
        Me.txFolder.Size = New System.Drawing.Size(322, 22)
        Me.txFolder.TabIndex = 7
        '
        'btnBrowse2
        '
        Me.btnBrowse2.Font = New System.Drawing.Font("Courier New", 9.75!)
        Me.btnBrowse2.Location = New System.Drawing.Point(418, 29)
        Me.btnBrowse2.Name = "btnBrowse2"
        Me.btnBrowse2.Size = New System.Drawing.Size(75, 22)
        Me.btnBrowse2.TabIndex = 0
        Me.btnBrowse2.Text = "Browse"
        Me.btnBrowse2.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Folder :"
        '
        'cbDelOri
        '
        Me.cbDelOri.AutoSize = True
        Me.cbDelOri.Location = New System.Drawing.Point(418, 90)
        Me.cbDelOri.Name = "cbDelOri"
        Me.cbDelOri.Size = New System.Drawing.Size(80, 17)
        Me.cbDelOri.TabIndex = 24
        Me.cbDelOri.Text = "Del. Ori File"
        Me.cbDelOri.UseVisualStyleBackColor = True
        '
        'cbData
        '
        Me.cbData.AutoSize = True
        Me.cbData.Location = New System.Drawing.Point(335, 90)
        Me.cbData.Name = "cbData"
        Me.cbData.Size = New System.Drawing.Size(79, 17)
        Me.cbData.TabIndex = 25
        Me.cbData.Text = "Using Data"
        Me.cbData.UseVisualStyleBackColor = True
        '
        'MS
        '
        Me.MS.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GetListFilesToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MS.Location = New System.Drawing.Point(0, 0)
        Me.MS.Name = "MS"
        Me.MS.Size = New System.Drawing.Size(507, 24)
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
        'RenameFiles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(507, 482)
        Me.Controls.Add(Me.MS)
        Me.Controls.Add(Me.cbData)
        Me.Controls.Add(Me.cbDelOri)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txFolder)
        Me.Controls.Add(Me.btnBrowse2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txSheet)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lbl_record)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnProses)
        Me.Controls.Add(Me.txRange)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txBrowse)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(513, 511)
        Me.MinimumSize = New System.Drawing.Size(513, 511)
        Me.Name = "RenameFiles"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "RenameFiles"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MS.ResumeLayout(False)
        Me.MS.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txBrowse As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txRange As System.Windows.Forms.TextBox
    Friend WithEvents btnProses As System.Windows.Forms.Button
    Friend WithEvents lbl_record As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txSheet As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents txFolder As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowse2 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbDelOri As System.Windows.Forms.CheckBox
    Friend WithEvents cbData As System.Windows.Forms.CheckBox
    Friend WithEvents MS As System.Windows.Forms.MenuStrip
    Friend WithEvents GetListFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GetListFilesToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FindReplaceFillesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RenameFilesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
