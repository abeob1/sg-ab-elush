<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEmailToVendor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEmailToVendor))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ckOpenOnly = New System.Windows.Forms.CheckBox()
        Me.cbToDate = New System.Windows.Forms.DateTimePicker()
        Me.cbFromDate = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnView = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSendEmail = New System.Windows.Forms.Button()
        Me.grData = New System.Windows.Forms.DataGridView()
        Me.Check = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.WhsCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DocEntry = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DocNum = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DocDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DocDueDate = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CardCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CardName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DocTotal = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.E_Mail = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ToEmailList = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.grData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ckOpenOnly)
        Me.Panel1.Controls.Add(Me.cbToDate)
        Me.Panel1.Controls.Add(Me.cbFromDate)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnView)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1037, 73)
        Me.Panel1.TabIndex = 23
        '
        'ckOpenOnly
        '
        Me.ckOpenOnly.AutoSize = True
        Me.ckOpenOnly.Checked = True
        Me.ckOpenOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckOpenOnly.Location = New System.Drawing.Point(198, 42)
        Me.ckOpenOnly.Name = "ckOpenOnly"
        Me.ckOpenOnly.Size = New System.Drawing.Size(77, 18)
        Me.ckOpenOnly.TabIndex = 21
        Me.ckOpenOnly.Text = "Open Only"
        Me.ckOpenOnly.UseVisualStyleBackColor = True
        '
        'cbToDate
        '
        Me.cbToDate.CustomFormat = "dd/MM/yyyy"
        Me.cbToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.cbToDate.Location = New System.Drawing.Point(77, 41)
        Me.cbToDate.Name = "cbToDate"
        Me.cbToDate.Size = New System.Drawing.Size(105, 20)
        Me.cbToDate.TabIndex = 20
        '
        'cbFromDate
        '
        Me.cbFromDate.CustomFormat = "dd/MM/yyyy"
        Me.cbFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.cbFromDate.Location = New System.Drawing.Point(77, 12)
        Me.cbFromDate.Name = "cbFromDate"
        Me.cbFromDate.Size = New System.Drawing.Size(105, 20)
        Me.cbFromDate.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 14)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "To Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 14)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "From Date"
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(307, 24)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 36)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.btnSendEmail)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 432)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1037, 49)
        Me.Panel3.TabIndex = 25
        '
        'btnSendEmail
        '
        Me.btnSendEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSendEmail.Location = New System.Drawing.Point(956, 6)
        Me.btnSendEmail.Name = "btnSendEmail"
        Me.btnSendEmail.Size = New System.Drawing.Size(75, 34)
        Me.btnSendEmail.TabIndex = 2
        Me.btnSendEmail.Text = "Send Email"
        Me.btnSendEmail.UseVisualStyleBackColor = True
        '
        'grData
        '
        Me.grData.AllowUserToAddRows = False
        Me.grData.AllowUserToDeleteRows = False
        Me.grData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grData.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Check, Me.WhsCode, Me.DocEntry, Me.DocNum, Me.DocDate, Me.DocDueDate, Me.CardCode, Me.CardName, Me.DocTotal, Me.E_Mail, Me.ToEmailList})
        Me.grData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grData.Location = New System.Drawing.Point(0, 73)
        Me.grData.Name = "grData"
        Me.grData.Size = New System.Drawing.Size(1037, 359)
        Me.grData.TabIndex = 26
        '
        'Check
        '
        Me.Check.HeaderText = "Select"
        Me.Check.Name = "Check"
        Me.Check.Width = 50
        '
        'WhsCode
        '
        Me.WhsCode.DataPropertyName = "WhsCode"
        Me.WhsCode.HeaderText = "Store"
        Me.WhsCode.Name = "WhsCode"
        '
        'DocEntry
        '
        Me.DocEntry.DataPropertyName = "DocEntry"
        Me.DocEntry.HeaderText = "Doc Entry"
        Me.DocEntry.Name = "DocEntry"
        Me.DocEntry.ReadOnly = True
        '
        'DocNum
        '
        Me.DocNum.DataPropertyName = "DocNum"
        Me.DocNum.HeaderText = "Doc. Number"
        Me.DocNum.Name = "DocNum"
        Me.DocNum.ReadOnly = True
        '
        'DocDate
        '
        Me.DocDate.DataPropertyName = "DocDate"
        Me.DocDate.HeaderText = "Document Date"
        Me.DocDate.Name = "DocDate"
        Me.DocDate.ReadOnly = True
        Me.DocDate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DocDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'DocDueDate
        '
        Me.DocDueDate.DataPropertyName = "DocDueDate"
        Me.DocDueDate.HeaderText = "Due Date"
        Me.DocDueDate.Name = "DocDueDate"
        Me.DocDueDate.ReadOnly = True
        Me.DocDueDate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DocDueDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'CardCode
        '
        Me.CardCode.DataPropertyName = "CardCode"
        Me.CardCode.HeaderText = "Vendor Code"
        Me.CardCode.Name = "CardCode"
        Me.CardCode.ReadOnly = True
        Me.CardCode.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CardCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'CardName
        '
        Me.CardName.DataPropertyName = "CardName"
        Me.CardName.HeaderText = "Vendor Name"
        Me.CardName.Name = "CardName"
        Me.CardName.ReadOnly = True
        Me.CardName.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CardName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CardName.Width = 200
        '
        'DocTotal
        '
        Me.DocTotal.DataPropertyName = "DocTotal"
        Me.DocTotal.HeaderText = "Total"
        Me.DocTotal.Name = "DocTotal"
        Me.DocTotal.ReadOnly = True
        Me.DocTotal.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DocTotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'E_Mail
        '
        Me.E_Mail.DataPropertyName = "E_Mail"
        Me.E_Mail.HeaderText = "Email"
        Me.E_Mail.Name = "E_Mail"
        Me.E_Mail.ReadOnly = True
        '
        'ToEmailList
        '
        Me.ToEmailList.DataPropertyName = "ToEmailList"
        Me.ToEmailList.HeaderText = "Email(s) CC"
        Me.ToEmailList.Name = "ToEmailList"
        Me.ToEmailList.ReadOnly = True
        Me.ToEmailList.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ToEmailList.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ToEmailList.Width = 200
        '
        'frmEmailToVendor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1037, 481)
        Me.Controls.Add(Me.grData)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEmailToVendor"
        Me.Text = "Email To Vendor (Ver. 20130221)"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        CType(Me.grData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnSendEmail As System.Windows.Forms.Button
    Friend WithEvents grData As System.Windows.Forms.DataGridView
    Friend WithEvents ckOpenOnly As System.Windows.Forms.CheckBox
    Friend WithEvents cbToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Check As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents WhsCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DocEntry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DocNum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DocDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DocDueDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CardCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CardName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DocTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents E_Mail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ToEmailList As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
