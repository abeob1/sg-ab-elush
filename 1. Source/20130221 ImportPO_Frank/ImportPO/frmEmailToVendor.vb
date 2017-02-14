Public Class frmEmailToVendor

    Private Sub btnView_Click(sender As System.Object, e As System.EventArgs) Handles btnView.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim fn As New Functions
            Dim strquery As String = "exec sp_AI_LoadPOforEmail '" + cbFromDate.Value.Date.ToString("MM/dd/yyyy") + "','" + cbToDate.Value.Date.ToString("MM/dd/yyyy") + "',"
            If ckOpenOnly.Checked Then
                strquery = strquery + "'O'"
            Else
                strquery = strquery + "''"
            End If

            Dim dt As DataTable = fn.DoQueryReturnDT(strquery)
            If Not IsNothing(dt) Then
                grData.DataSource = dt.DefaultView
            Else
                MessageBox.Show("There's no data!")
            End If

            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub frmEmailToVendor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        grData.AutoGenerateColumns = False
    End Sub

    Private Sub btnSendEmail_Click(sender As System.Object, e As System.EventArgs) Handles btnSendEmail.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim om As New oEmailPO2VendorWeb
            Dim dt As New DataTable

            dt.Columns.Add("DocEntry")
            dt.Columns.Add("E_Mail")
            dt.Columns.Add("ToEmailList")
            dt.Columns.Add("CardCode")
            dt.Columns.Add("WhsCode")
            dt.Columns.Add("DocNum")
            dt.Columns.Add("CardName")
            dt.Columns.Add("DocDate")
            dt.Columns.Add("GrandTotal")
            dt.Columns.Add("SubTotal")
            dt.Columns.Add("GST")

            For i As Integer = 0 To grData.RowCount - 1
                'grData.Rows(i).Selected = True
                If grData.Rows(i).Cells("Check").Value = True Then
                    Dim drNew As DataRow = dt.NewRow
                    drNew("DocEntry") = grData.Rows(i).Cells("DocEntry").Value.ToString
                    drNew("E_Mail") = grData.Rows(i).Cells("E_Mail").Value.ToString
                    drNew("ToEmailList") = grData.Rows(i).Cells("ToEmailList").Value.ToString
                    drNew("CardCode") = grData.Rows(i).Cells("CardCode").Value.ToString
                    drNew("WhsCode") = grData.Rows(i).Cells("WhsCode").Value.ToString
                    drNew("DocNum") = grData.Rows(i).Cells("DocNum").Value.ToString

                    drNew("CardName") = grData.Rows(i).Cells("CardName").Value.ToString
                    drNew("DocDate") = grData.Rows(i).Cells("DocDate").Value.ToString
                    drNew("GrandTotal") = grData.Rows(i).Cells("DocTotal").Value.ToString
                    'TODO UPDATE subtotal and GST
                    drNew("SubTotal") = grData.Rows(i).Cells("DocTotal").Value.ToString
                    drNew("GST") = grData.Rows(i).Cells("DocTotal").Value.ToString
                    dt.Rows.Add(drNew)
                End If
            Next
            'Dim fn As New Functions
            'fn.LoadParameters() 'get email information from SAP_Integration DB

            Dim ret As String = om.SendPOEmail(dt)
            If ret = "" Then ret = "Email(s) have been sent!"
            MessageBox.Show(ret)
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class