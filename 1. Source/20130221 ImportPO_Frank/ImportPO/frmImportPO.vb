Public Class frmImportPO
    Dim dt As DataTable
    Dim de As ExcelDrilling
    Private Sub cbSheet_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbSheet.SelectedIndexChanged
        LoadExcelData()
    End Sub
    Private Sub LoadExcelData()
        Try
            Dim sheet As String = cbSheet.SelectedValue.ToString
            dt = de.GetDataSQL("Select * from [" + sheet + "]")
            grData.DataSource = dt
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs)
        Dim sheet As String = cbSheet.SelectedValue.ToString
        dt = de.GetDataSQL("Select * from [" + sheet + "]")
        Dim gen As New oGeneratePO
        gen.GeneratePO(dt)
    End Sub

    Private Sub btnBrowseFile_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowseFile.Click
        Try
            Dim filename As String
            OpenFileDialog1.Title = "Select Excel File"
            OpenFileDialog1.InitialDirectory = "C:\"
            OpenFileDialog1.Filter = "Excel File | *.xls;*.xlsx"

            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                filename = OpenFileDialog1.FileName
                txtFileName.Text = filename
                de = New ExcelDrilling(filename)
                Dim dt As DataTable = de.GetSheets
                cbSheet.DataSource = dt
                cbSheet.DisplayMember = "TABLE_NAME"
                cbSheet.ValueMember = "TABLE_NAME"
                If dt.Rows.Count > 0 Then
                    cbSheet.SelectedIndex = 0
                    LoadExcelData()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnGenerate_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerate.Click
        PublicVariable.Token = Functions.CreateToken
        If PublicVariable.Token <> "" Then
            Dim sheet As String = cbSheet.SelectedValue.ToString
            dt = de.GetDataSQL("Select * from [" + sheet + "]")
            Dim gen As New oGeneratePO
            Dim st As String = gen.GeneratePO(dt)
            If st <> "" Then
                MessageBox.Show(st)
            Else
                Dim dt As DataTable
                Dim fn As New Functions
                dt = fn.DoQueryReturnDT("select * from AIPOImpH where Token='" + PublicVariable.Token + "'")
                grPOHeader.DataSource = dt
                MessageBox.Show("PO(s) generated!")
                btnPost.Enabled = True
            End If
        End If
    End Sub


    Private Sub grPOHeader_SelectionChanged(sender As System.Object, e As System.EventArgs) Handles grPOHeader.SelectionChanged
        If grPOHeader.SelectedRows.Count > 0 Then
            Dim dt As DataTable
            Dim fn As New Functions
            dt = fn.DoQueryReturnDT("select * from AIPOImpL where HeaderID=" + grPOHeader.SelectedRows.Item(0).Cells("ID").Value.ToString)
            grPOLine.DataSource = dt
        End If


    End Sub

    Private Sub btnSimulate_Click(sender As System.Object, e As System.EventArgs) Handles btnSimulate.Click
        If MessageBox.Show("During simulation process, SAP will be hold. Do you want to continue?", "Warning", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Me.Cursor = Cursors.WaitCursor
            Dim po As New oPO
            po.CreateAllPO(True)
            Dim dt As DataTable
            Dim fn As New Functions
            dt = fn.DoQueryReturnDT("select * from AIPOImpH where Token='" + PublicVariable.Token + "'")
            grPOHeader.DataSource = dt
            MessageBox.Show("Complete!")
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub btnPost_Click(sender As System.Object, e As System.EventArgs) Handles btnPost.Click
        Me.Cursor = Cursors.WaitCursor
        btnPost.Enabled = False
        Dim po As New oPO
        po.CreateAllPO(False)
        Dim dt As DataTable
        Dim fn As New Functions
        dt = fn.DoQueryReturnDT("select * from AIPOImpH where Token='" + PublicVariable.Token + "'")
        grPOHeader.DataSource = dt
        MessageBox.Show("Complete!")
        Me.Cursor = Cursors.Default
    End Sub
End Class
