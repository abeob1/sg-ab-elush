<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMornitor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMornitor))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblRunning = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnUpload = New System.Windows.Forms.Button()
        Me.btnUnReg = New System.Windows.Forms.Button()
        Me.btnReg = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.cbSendDate = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbResult = New System.Windows.Forms.ComboBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ckAutoRef = New System.Windows.Forms.CheckBox()
        Me.cbFilter = New System.Windows.Forms.ComboBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.grMonitor = New System.Windows.Forms.DataGridView()
        Me.grDetail = New System.Windows.Forms.DataGridView()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnLog = New System.Windows.Forms.Button()
        Me.btnRetryAll = New System.Windows.Forms.Button()
        Me.btnRetry = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.grMonitor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.lblRunning)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.btnUpload)
        Me.Panel2.Controls.Add(Me.btnUnReg)
        Me.Panel2.Controls.Add(Me.btnReg)
        Me.Panel2.Controls.Add(Me.btnStop)
        Me.Panel2.Controls.Add(Me.btnStart)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 433)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(955, 39)
        Me.Panel2.TabIndex = 1
        '
        'lblRunning
        '
        Me.lblRunning.AutoSize = True
        Me.lblRunning.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRunning.Location = New System.Drawing.Point(516, 21)
        Me.lblRunning.Name = "lblRunning"
        Me.lblRunning.Size = New System.Drawing.Size(0, 14)
        Me.lblRunning.TabIndex = 12
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(427, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 14)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Running Status:"
        '
        'btnUpload
        '
        Me.btnUpload.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUpload.Location = New System.Drawing.Point(327, 6)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(94, 29)
        Me.btnUpload.TabIndex = 10
        Me.btnUpload.Text = "Upload Wincor"
        Me.btnUpload.UseVisualStyleBackColor = True
        '
        'btnUnReg
        '
        Me.btnUnReg.Location = New System.Drawing.Point(246, 5)
        Me.btnUnReg.Name = "btnUnReg"
        Me.btnUnReg.Size = New System.Drawing.Size(75, 30)
        Me.btnUnReg.TabIndex = 3
        Me.btnUnReg.Text = "Un-Register"
        Me.btnUnReg.UseVisualStyleBackColor = True
        '
        'btnReg
        '
        Me.btnReg.Location = New System.Drawing.Point(165, 5)
        Me.btnReg.Name = "btnReg"
        Me.btnReg.Size = New System.Drawing.Size(75, 30)
        Me.btnReg.TabIndex = 2
        Me.btnReg.Text = "Register"
        Me.btnReg.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(84, 5)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(75, 30)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(3, 5)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 30)
        Me.btnStart.TabIndex = 0
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(7, 352)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 30)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.cbSendDate)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.cbResult)
        Me.Panel3.Controls.Add(Me.btnRefresh)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.ckAutoRef)
        Me.Panel3.Controls.Add(Me.cbFilter)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(955, 45)
        Me.Panel3.TabIndex = 3
        '
        'cbSendDate
        '
        Me.cbSendDate.Checked = False
        Me.cbSendDate.CustomFormat = "dd/MM/yyyy"
        Me.cbSendDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.cbSendDate.Location = New System.Drawing.Point(635, 15)
        Me.cbSendDate.Name = "cbSendDate"
        Me.cbSendDate.ShowCheckBox = True
        Me.cbSendDate.Size = New System.Drawing.Size(118, 20)
        Me.cbSendDate.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(568, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 14)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Send Date"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(307, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 14)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Result Filter"
        '
        'cbResult
        '
        Me.cbResult.FormattingEnabled = True
        Me.cbResult.Items.AddRange(New Object() {"All", "Pending", "Successfull", "Failed"})
        Me.cbResult.Location = New System.Drawing.Point(394, 12)
        Me.cbResult.Name = "cbResult"
        Me.cbResult.Size = New System.Drawing.Size(142, 22)
        Me.cbResult.TabIndex = 6
        Me.cbResult.Text = "Pending"
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(759, 7)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 30)
        Me.btnRefresh.TabIndex = 5
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 14)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Document Filter"
        '
        'ckAutoRef
        '
        Me.ckAutoRef.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ckAutoRef.AutoSize = True
        Me.ckAutoRef.Location = New System.Drawing.Point(856, 19)
        Me.ckAutoRef.Name = "ckAutoRef"
        Me.ckAutoRef.Size = New System.Drawing.Size(87, 18)
        Me.ckAutoRef.TabIndex = 1
        Me.ckAutoRef.Text = "Auto Refesh"
        Me.ckAutoRef.UseVisualStyleBackColor = True
        '
        'cbFilter
        '
        Me.cbFilter.FormattingEnabled = True
        Me.cbFilter.Items.AddRange(New Object() {"Item", "Business Partner", "Purchase Order", "GRPO", "Goods Return", "Inventory Transfer", "Invoice", "Goods Receipt", "Goods Issue", "Stock Take", "Send Email", "Wincor Sales"})
        Me.cbFilter.Location = New System.Drawing.Point(99, 12)
        Me.cbFilter.Name = "cbFilter"
        Me.cbFilter.Size = New System.Drawing.Size(192, 22)
        Me.cbFilter.TabIndex = 0
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.grMonitor)
        Me.Panel4.Controls.Add(Me.grDetail)
        Me.Panel4.Controls.Add(Me.Panel1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(0, 45)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(955, 388)
        Me.Panel4.TabIndex = 4
        '
        'grMonitor
        '
        Me.grMonitor.AllowUserToAddRows = False
        Me.grMonitor.AllowUserToDeleteRows = False
        Me.grMonitor.AllowUserToOrderColumns = True
        Me.grMonitor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grMonitor.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grMonitor.Location = New System.Drawing.Point(0, 0)
        Me.grMonitor.Name = "grMonitor"
        Me.grMonitor.Size = New System.Drawing.Size(868, 252)
        Me.grMonitor.TabIndex = 6
        '
        'grDetail
        '
        Me.grDetail.AllowUserToAddRows = False
        Me.grDetail.AllowUserToDeleteRows = False
        Me.grDetail.AllowUserToOrderColumns = True
        Me.grDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grDetail.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grDetail.Location = New System.Drawing.Point(0, 252)
        Me.grDetail.Name = "grDetail"
        Me.grDetail.Size = New System.Drawing.Size(868, 136)
        Me.grDetail.TabIndex = 5
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.btnLog)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnRetryAll)
        Me.Panel1.Controls.Add(Me.btnRetry)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel1.Location = New System.Drawing.Point(868, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(87, 388)
        Me.Panel1.TabIndex = 3
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(6, 186)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 53)
        Me.Button3.TabIndex = 11
        Me.Button3.Text = "Retry Invoice with Flag"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(7, 150)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 30)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "XML string"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(7, 114)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 30)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Mapping"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnLog
        '
        Me.btnLog.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLog.Location = New System.Drawing.Point(7, 78)
        Me.btnLog.Name = "btnLog"
        Me.btnLog.Size = New System.Drawing.Size(75, 30)
        Me.btnLog.TabIndex = 8
        Me.btnLog.Text = "Log File"
        Me.btnLog.UseVisualStyleBackColor = True
        '
        'btnRetryAll
        '
        Me.btnRetryAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRetryAll.Location = New System.Drawing.Point(7, 42)
        Me.btnRetryAll.Name = "btnRetryAll"
        Me.btnRetryAll.Size = New System.Drawing.Size(75, 30)
        Me.btnRetryAll.TabIndex = 7
        Me.btnRetryAll.Text = "Retry All"
        Me.btnRetryAll.UseVisualStyleBackColor = True
        '
        'btnRetry
        '
        Me.btnRetry.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRetry.Location = New System.Drawing.Point(7, 6)
        Me.btnRetry.Name = "btnRetry"
        Me.btnRetry.Size = New System.Drawing.Size(75, 30)
        Me.btnRetry.TabIndex = 6
        Me.btnRetry.Text = "Retry"
        Me.btnRetry.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 10000
        '
        'frmMornitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(955, 472)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMornitor"
        Me.Text = "Integration Mornitor (V 2013.06.13)"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        CType(Me.grMonitor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUnReg As System.Windows.Forms.Button
    Friend WithEvents btnReg As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ckAutoRef As System.Windows.Forms.CheckBox
    Friend WithEvents cbFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbResult As System.Windows.Forms.ComboBox
    Friend WithEvents btnRetry As System.Windows.Forms.Button
    Friend WithEvents btnRetryAll As System.Windows.Forms.Button
    Friend WithEvents btnLog As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnUpload As System.Windows.Forms.Button
    Friend WithEvents grMonitor As System.Windows.Forms.DataGridView
    Friend WithEvents grDetail As System.Windows.Forms.DataGridView
    Friend WithEvents cbSendDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblRunning As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
End Class
