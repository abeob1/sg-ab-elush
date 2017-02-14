Public Class oGeneratePO
    'Seperate PO:
    ' - Per warehouse
    ' - Per Vendor
    ' - Per 20 lines
    
    Public Function GeneratePO(dt As DataTable) As String
        Try
            ' -------------GET ALL RECORDS BY VENDOR------------------
            Dim dtDistinctVendor As DataTable
            dtDistinctVendor = dt.DefaultView.ToTable("VendorCode", True, "VendorCode").Copy
            For Each drVendor As DataRow In dtDistinctVendor.Rows
                Dim dtByVendor As DataTable = dt.Clone
                For Each dr As DataRow In dt.Select("VendorCode='" + drVendor("VendorCode").ToString() + "'")
                    dtByVendor.ImportRow(dr)
                Next
                ' -------------GET ALL RECORDS BY WAREHOUSE------------------
                Dim dtDistinctWhs As DataTable
                dtDistinctWhs = dtByVendor.DefaultView.ToTable("Warehouse", True, "Warehouse")
                For Each drWhs As DataRow In dtDistinctWhs.Rows
                    Dim dtByWarehouse As DataTable = dt.Clone
                    Dim strFiler As String = "Warehouse='" + drWhs("Warehouse").ToString() + "'"
                    For Each dr As DataRow In dtByVendor.Select(strFiler)
                        dtByWarehouse.ImportRow(dr)
                    Next
                    '--------------INSERT INTO PO------------------
                    Dim i As Integer = 0
                    Dim HeaderID As Integer

                    HeaderID = InsertPOHeader(dtByWarehouse.Rows(0))
                    For Each dr As DataRow In dtByWarehouse.Rows
                        InsertPORow(dr, HeaderID)
                        i = i + 1
                        '--------------20 Lines per PO------------------
                        If i > 20 Then
                            HeaderID = InsertPOHeader(dtByWarehouse.Rows(0))
                            i = 0
                        End If
                    Next
                Next
            Next
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
    Private Function InsertPOHeader(dr As DataRow) As Integer

        Dim ors As SAPbobsCOM.Recordset
        Dim str As String = "exec sp_AI_InsertH "
        str = str + "'" + PublicVariable.Token + "',"
        str = str + "'" + dr("VendorCode").ToString + "',"
        str = str + "'" + dr("VendorReferenceNo").ToString + "',"
        str = str + "'" + dr("Remarks").ToString + "'"

        ors = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ors.DoQuery(str)
        ors.MoveFirst()
        If ors.RecordCount > 0 Then
            Return ors.Fields.Item("ID").Value
            MessageBox.Show(ors.Fields.Item("ID").Value)
        Else
            Return 0
        End If

    End Function
    Private Sub InsertPORow(dr As DataRow, ID As Integer)
        Dim ors As SAPbobsCOM.Recordset
        Dim str As String = "exec sp_AI_InsertL "
        str = str + ID.ToString + ","
        str = str + "'" + dr("ItemCode").ToString + "',"
        str = str + dr("Quantity").ToString + ","
        str = str + "'" + dr("Warehouse").ToString + "'"

        ors = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ors.DoQuery(str)

    End Sub
End Class
