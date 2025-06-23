<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Label6 = New System.Windows.Forms.Label
        Me.txBrowse = New System.Windows.Forms.TextBox
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.txDB = New System.Windows.Forms.TextBox
        Me.txPass = New System.Windows.Forms.TextBox
        Me.txID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnProcess = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txFind = New System.Windows.Forms.TextBox
        Me.txMove = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txType = New System.Windows.Forms.TextBox
        Me.txDigDat = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(12, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(85, 22)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Directory:"
        '
        'txBrowse
        '
        Me.txBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txBrowse.Location = New System.Drawing.Point(96, 46)
        Me.txBrowse.Name = "txBrowse"
        Me.txBrowse.Size = New System.Drawing.Size(459, 22)
        Me.txBrowse.TabIndex = 14
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(444, 8)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(111, 32)
        Me.btnBrowse.TabIndex = 13
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txDB
        '
        Me.txDB.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txDB.Location = New System.Drawing.Point(319, 9)
        Me.txDB.Name = "txDB"
        Me.txDB.Size = New System.Drawing.Size(96, 26)
        Me.txDB.TabIndex = 21
        Me.txDB.Text = "elnusa"
        '
        'txPass
        '
        Me.txPass.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txPass.Location = New System.Drawing.Point(202, 9)
        Me.txPass.Name = "txPass"
        Me.txPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(94)
        Me.txPass.Size = New System.Drawing.Size(77, 26)
        Me.txPass.TabIndex = 20
        Me.txPass.Text = "elnusa"
        '
        'txID
        '
        Me.txID.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txID.Location = New System.Drawing.Point(76, 9)
        Me.txID.Name = "txID"
        Me.txID.Size = New System.Drawing.Size(76, 26)
        Me.txID.TabIndex = 18
        Me.txID.Text = "elnusa"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(285, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 22)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "DB:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(158, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 22)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Pwd:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 22)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "UserID:"
        '
        'btnProcess
        '
        Me.btnProcess.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProcess.Location = New System.Drawing.Point(444, 143)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(111, 32)
        Me.btnProcess.TabIndex = 23
        Me.btnProcess.Text = "Process"
        Me.btnProcess.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(6, 189)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(549, 345)
        Me.TextBox1.TabIndex = 25
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(12, 142)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 22)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "Filter        :"
        '
        'txFind
        '
        Me.txFind.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txFind.Location = New System.Drawing.Point(96, 143)
        Me.txFind.Name = "txFind"
        Me.txFind.Size = New System.Drawing.Size(67, 22)
        Me.txFind.TabIndex = 26
        Me.txFind.Text = "*"
        '
        'txMove
        '
        Me.txMove.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.txMove.Location = New System.Drawing.Point(243, 143)
        Me.txMove.Name = "txMove"
        Me.txMove.Size = New System.Drawing.Size(121, 22)
        Me.txMove.TabIndex = 33
        Me.txMove.Text = "E:\[Junk]"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(169, 142)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 22)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Move to:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(12, 96)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(87, 22)
        Me.Label7.TabIndex = 35
        Me.Label7.Text = "File Type :"
        '
        'txType
        '
        Me.txType.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txType.Location = New System.Drawing.Point(96, 97)
        Me.txType.Multiline = True
        Me.txType.Name = "txType"
        Me.txType.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txType.Size = New System.Drawing.Size(459, 40)
        Me.txType.TabIndex = 36
        Me.txType.Text = ".BMP;.CGM;.DWG;.GIF;.JPG;.JPEG;.LAS;.PDF;.PDS;.PPT;.PPT;.PPTX;.TIF;.TIFF;.XLS;.XL" & _
            "SX"
        '
        'txDigDat
        '
        Me.txDigDat.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txDigDat.Location = New System.Drawing.Point(96, 71)
        Me.txDigDat.Name = "txDigDat"
        Me.txDigDat.Size = New System.Drawing.Size(459, 22)
        Me.txDigDat.TabIndex = 37
        Me.txDigDat.Text = "E:\DigDat"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(12, 68)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 22)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "DigDat    :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label9.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(453, 513)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(81, 17)
        Me.Label9.TabIndex = 39
        Me.Label9.Text = "by armayndo"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(567, 546)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txDigDat)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txType)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txMove)
        Me.Controls.Add(Me.txFind)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.btnProcess)
        Me.Controls.Add(Me.txDB)
        Me.Controls.Add(Me.txPass)
        Me.Controls.Add(Me.txID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txBrowse)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label5)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(575, 580)
        Me.MinimumSize = New System.Drawing.Size(575, 580)
        Me.Name = "Form1"
        Me.Text = "JunkFile Finder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txBrowse As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents txDB As System.Windows.Forms.TextBox
    Friend WithEvents txPass As System.Windows.Forms.TextBox
    Friend WithEvents txID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnProcess As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txFind As System.Windows.Forms.TextBox
    Friend WithEvents txMove As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txType As System.Windows.Forms.TextBox
    Friend WithEvents txDigDat As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label

End Class
