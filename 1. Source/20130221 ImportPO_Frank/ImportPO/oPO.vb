Public Class oPO
    'Default:
    ' - Posting Date: today
    ' - Delivery Date: next day
    ' - Price: price list of warehouse
    ' - PO Type: OR
    'Send email
    'Validation: if PO less than MOQ or MOA, can't create PO
#Region "Build Table Structure"
    Private Function BuildTableOPOR() As DataTable
        Dim dt As New DataTable("OPOR")
        dt.Columns.Add("CardCode")
        dt.Columns.Add("DocDate")
        dt.Columns.Add("DocDueDate")
        dt.Columns.Add("NumAtCard")
        dt.Columns.Add("Comments")
        dt.Columns.Add("U_Type")
        dt.Columns.Add("U_POSTxNo")
        Return dt
    End Function
    Private Function BuildTablePOR1() As DataTable
        Dim dt As New DataTable("POR1")
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("WhsCode")
        dt.Columns.Add("Quantity", Type.GetType("System.Double"))
        dt.Columns.Add("PriceBefDi", Type.GetType("System.Double"))
        Return dt
    End Function
    Private Function BuildTableSimulateLog() As DataTable
        Dim dt As New DataTable("POR1")
        dt.Columns.Add("ID")
        dt.Columns.Add("ErrMsg")
        Return dt
    End Function
#End Region
#Region "Insert into Table"
    Private Function InsertIntoOPOR(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("CardCode") = dr("VendorCode")
        drNew("NumAtCard") = dr("VendorReferenceNo")
        drNew("Comments") = dr("Remarks")
        drNew("DocDate") = Now.Date.ToString("yyyyMMdd")
        drNew("DocDueDate") = Now.Date.AddDays(1).ToString("yyyyMMdd")
        drNew("U_Type") = "OR"
        drNew("U_POSTxNo") = PublicVariable.Token
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoPOR1(dt As DataTable, dr As DataRow) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("ItemCode") = dr("ItemCode")
        drNew("WhsCode") = dr("Warehouse")
        drNew("Quantity") = dr("Quantity")
        drNew("PriceBefDi") = GetPriceByWarehouse(dr("ItemCode"), dr("Warehouse"))
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoSimulateLog(dt As DataTable, ID As Integer, ret As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("ID") = ID
        drNew("ErrMsg") = ret
        dt.Rows.Add(drNew)
        Return dt
    End Function
#End Region
    Private Function POValidation(dtPOLine As DataTable, VendorCode As String) As Boolean
        Dim totalQuantity As Double = dtPOLine.Compute("Sum(Quantity)", "")
        Dim totalAmount As Double = 0 'dtPOLine.Compute("Sum(""Quantity*PriceBefDi"")", "")
        For Each dr As DataRow In dtPOLine.Rows
            totalAmount = totalAmount + dr("Quantity") * dr("PriceBefDi")
        Next

        Dim ors As SAPbobsCOM.Recordset
        ors = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ors.DoQuery("select isnull(U_POMOQ,0) POMOQ ,isnull(U_POMOA,0) POMOA from OCRD where cardcode='" + VendorCode + "'")
        If ors.RecordCount > 0 Then
            If totalAmount < ors.Fields.Item("POMOA").Value Or totalQuantity < ors.Fields.Item("POMOQ").Value Then
                Return False
            End If
        End If
        Return True
    End Function
    Private Function GetPriceByWarehouse(ItemCode As String, WhsCode As String) As Double
        Dim ors As SAPbobsCOM.Recordset
        ors = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ors.DoQuery("select Price from ITM1 T0 join OWHS T1 on T0.PriceList=T1.U_POPriceList where T0.ItemCode='" + ItemCode + "' and T1.WhsCode='" + WhsCode + "'")
        If ors.RecordCount > 0 Then
            Return ors.Fields.Item("Price").Value
        Else
            Return 0
        End If
    End Function
    Public Sub CreateAllPO(Simulate As Boolean)
        'Simulate: alway rollback
        Try
            Dim ofn As New Functions
            Dim dt As DataTable = ofn.DoQueryReturnDT("select * from AIPOImpH where Token='" + PublicVariable.Token + "'")
            If dt.Rows.Count = 0 Then Return

            Dim dtErrorLog As DataTable = BuildTableSimulateLog()
            

            For Each dr As DataRow In dt.Rows
                If Simulate Then
                    If Not PublicVariable.oCompany.InTransaction Then
                        PublicVariable.oCompany.StartTransaction()
                    End If
                End If
                
                Dim dtPOHeader As DataTable = BuildTableOPOR()
                Dim dtPOLine As DataTable = BuildTablePOR1()

                Dim dt1 As DataTable
                Dim fn As New Functions
                dt1 = fn.DoQueryReturnDT("select * from AIPOImpL where HeaderID=" + dr("ID").ToString)
                Dim ret As String
                dtPOHeader = InsertIntoOPOR(dtPOHeader, dr)
                For Each dr1 As DataRow In dt1.Rows
                    dtPOLine = InsertIntoPOR1(dtPOLine, dr1)
                Next
                If POValidation(dtPOLine, dr("VendorCode").ToString) Then
                    Dim ds As New DataSet
                    ds.Tables.Add(dtPOHeader.Copy)
                    ds.Tables.Add(dtPOLine.Copy)

                    Dim ox As New oXML
                    Dim str As String = ox.ToXMLStringFromDS("22", ds)
                    ret = ox.CreateMarketingDocument(str, "22")
                    If ret = "" Then ret = "Sucessful!"
                Else
                    ret = "Not pass PO Validation"
                End If
                
                dtErrorLog = InsertIntoSimulateLog(dtErrorLog, dr("ID"), ret)
            Next
            If PublicVariable.oCompany.InTransaction Then
                    PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            UpdateErrorLog(dtErrorLog)
            SendEmail()
        Catch ex As Exception
            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub UpdateErrorLog(dtSimulateLog As DataTable)
        For Each dr As DataRow In dtSimulateLog.Rows
            Dim ors As SAPbobsCOM.Recordset
            ors = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors.DoQuery("exec sp_AI_UpdateLog " + dr("ID").ToString + ",'" + dr("ErrMsg").ToString.Replace("'", "") + "'")
        Next
    End Sub
    Private Sub SendEmail()
        Dim fn As New Functions
        Dim dt As DataTable = fn.DoQueryReturnDT("exec sp_AI_ImpPOEmail '" + PublicVariable.Token + "'")
        If IsNothing(dt) Then Return
        If dt.Rows.Count > 0 Then
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)

            Dim oe As New oEmailWeb
            Dim st As String = ""
            'fn.LoadParameters() 'get email information from SAP_Integration DB

            st = oe.SendMailByDS(ds, PublicVariable.ToEmail, PublicVariable.EmailSub, Application.StartupPath + "\EmailTemplate.htm")
            If st <> "" Then
                MessageBox.Show("Send email error: " + st)
            End If
        End If
    End Sub
End Class
