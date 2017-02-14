<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportPO
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportPO))
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.tp = New System.Windows.Forms.TabControl()
        Me.st1 = New System.Windows.Forms.TabPage()
        Me.grData = New System.Windows.Forms.DataGridView()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.btnBrowseFile = New System.Windows.Forms.Button()
        Me.cbSheet = New System.Windows.Forms.ComboBox()
        Me.st2 = New System.Windows.Forms.TabPage()
        Me.grPOLine = New System.Windows.Forms.DataGridView()
        Me.grPOHeader = New System.Windows.Forms.DataGridView()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.btnPost = New System.Windows.Forms.Button()
        Me.btnSimulate = New System.Windows.Forms.Button()
        Me.Panel2.SuspendLayout()
        Me.tp.SuspendLayout()
        Me.st1.SuspendLayout()
        CType(Me.grData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.st2.SuspendLayout()
        CType(Me.grPOLine, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grPOHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.tp)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(857, 449)
        Me.Panel2.TabIndex = 22
        '
        'tp
        '
        Me.tp.Controls.Add(Me.st1)
        Me.tp.Controls.Add(Me.st2)
        Me.tp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tp.Location = New System.Drawing.Point(0, 0)
        Me.tp.Name = "tp"
        Me.tp.SelectedIndex = 0
        Me.tp.Size = New System.Drawing.Size(857, 449)
        Me.tp.TabIndex = 19
        '
        'st1
        '
        Me.st1.Controls.Add(Me.grData)
        Me.st1.Controls.Add(Me.Panel3)
        Me.st1.Controls.Add(Me.Panel1)
        Me.st1.Location = New System.Drawing.Point(4, 23)
        Me.st1.Name = "st1"
        Me.st1.Padding = New System.Windows.Forms.Padding(3)
        Me.st1.Size = New System.Drawing.Size(849, 422)
        Me.st1.TabIndex = 0
        Me.st1.Text = "Step 1: Excel File"
        Me.st1.UseVisualStyleBackColor = True
        '
        'grData
        '
        Me.grData.AllowUserToAddRows = False
        Me.grData.AllowUserToDeleteRows = False
        Me.grData.AllowUserToOrderColumns = True
        Me.grData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grData.Location = New System.Drawing.Point(3, 76)
        Me.grData.Name = "grData"
        Me.grData.Size = New System.Drawing.Size(843, 294)
        Me.grData.TabIndex = 25
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.btnGenerate)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(3, 370)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(843, 49)
        Me.Panel3.TabIndex = 24
        '
        'btnGenerate
        '
        Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerate.Location = New System.Drawing.Point(763, 6)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(75, 34)
        Me.btnGenerate.TabIndex = 2
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtFileName)
        Me.Panel1.Controls.Add(Me.btnBrowseFile)
        Me.Panel1.Controls.Add(Me.cbSheet)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(843, 73)
        Me.Panel1.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 14)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Sheet Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 14)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "File Name"
        '
        'txtFileName
        '
        Me.txtFileName.Enabled = False
        Me.txtFileName.Location = New System.Drawing.Point(74, 11)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(456, 20)
        Me.txtFileName.TabIndex = 16
        '
        'btnBrowseFile
        '
        Me.btnBrowseFile.Location = New System.Drawing.Point(536, 8)
        Me.btnBrowseFile.Name = "btnBrowseFile"
        Me.btnBrowseFile.Size = New System.Drawing.Size(75, 23)
        Me.btnBrowseFile.TabIndex = 0
        Me.btnBrowseFile.Text = "Browser"
        Me.btnBrowseFile.UseVisualStyleBackColor = True
        '
        'cbSheet
        '
        Me.cbSheet.FormattingEnabled = True
        Me.cbSheet.Location = New System.Drawing.Point(74, 38)
        Me.cbSheet.Name = "cbSheet"
        Me.cbSheet.Size = New System.Drawing.Size(172, 22)
        Me.cbSheet.TabIndex = 1
        '
        'st2
        '
        Me.st2.Controls.Add(Me.grPOLine)
        Me.st2.Controls.Add(Me.grPOHeader)
        Me.st2.Controls.Add(Me.Panel4)
        Me.st2.Location = New System.Drawing.Point(4, 23)
        Me.st2.Name = "st2"
        Me.st2.Padding = New System.Windows.Forms.Padding(3)
        Me.st2.Size = New System.Drawing.Size(849, 422)
        Me.st2.TabIndex = 1
        Me.st2.Text = "Step 2: Generate PO(s)"
        Me.st2.UseVisualStyleBackColor = True
        '
        'grPOLine
        '
        Me.grPOLine.AllowUserToAddRows = False
        Me.grPOLine.AllowUserToDeleteRows = False
        Me.grPOLine.AllowUserToOrderColumns = True
        Me.grPOLine.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grPOLine.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grPOLine.Location = New System.Drawing.Point(3, 190)
        Me.grPOLine.Name = "grPOLine"
        Me.grPOLine.Size = New System.Drawing.Size(843, 180)
        Me.grPOLine.TabIndex = 27
        '
        'grPOHeader
        '
        Me.grPOHeader.AllowUserToAddRows = False
        Me.grPOHeader.AllowUserToDeleteRows = False
        Me.grPOHeader.AllowUserToOrderColumns = True
        Me.grPOHeader.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grPOHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.grPOHeader.Location = New System.Drawing.Point(3, 3)
        Me.grPOHeader.Name = "grPOHeader"
        Me.grPOHeader.Size = New System.Drawing.Size(843, 187)
        Me.grPOHeader.TabIndex = 26
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.btnPost)
        Me.Panel4.Controls.Add(Me.btnSimulate)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel4.Location = New System.Drawing.Point(3, 370)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(843, 49)
        Me.Panel4.TabIndex = 25
        '
        'btnPost
        '
        Me.btnPost.Location = New System.Drawing.Point(755, 6)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(75, 34)
        Me.btnPost.TabIndex = 3
        Me.btnPost.Text = "Post PO(s)"
        Me.btnPost.UseVisualStyleBackColor = True
        '
        'btnSimulate
        '
        Me.btnSimulate.Location = New System.Drawing.Point(674, 6)
        Me.btnSimulate.Name = "btnSimulate"
        Me.btnSimulate.Size = New System.Drawing.Size(75, 34)
        Me.btnSimulate.TabIndex = 2
        Me.btnSimulate.Text = "Simulate"
        Me.btnSimulate.UseVisualStyleBackColor = True
        '
        'frmImportPO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(857, 449)
        Me.Controls.Add(Me.Panel2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmImportPO"
        Me.Text = "Import Purchase Order (Ver. 20130204)"
        Me.Panel2.ResumeLayout(False)
        Me.tp.ResumeLayout(False)
        Me.st1.ResumeLayout(False)
        CType(Me.grData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.st2.ResumeLayout(False)
        CType(Me.grPOLine, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grPOHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents tp As System.Windows.Forms.TabControl
    Friend WithEvents st1 As System.Windows.Forms.TabPage
    Friend WithEvents st2 As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseFile As System.Windows.Forms.Button
    Friend WithEvents cbSheet As System.Windows.Forms.ComboBox
    Friend WithEvents grData As System.Windows.Forms.DataGridView
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents grPOHeader As System.Windows.Forms.DataGridView
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents btnSimulate As System.Windows.Forms.Button
    Friend WithEvents grPOLine As System.Windows.Forms.DataGridView
    Friend WithEvents btnPost As System.Windows.Forms.Button

End Class
