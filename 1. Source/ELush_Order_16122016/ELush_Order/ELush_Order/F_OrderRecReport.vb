Imports Excel = Microsoft.Office.Interop.Excel
'Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
Imports System.IO


Public Class F_OrderRecReport
    'Public Enumeration XlFileFormat
    Dim WithEvents SBO_Application As SAPbouiCOM.Application
    Dim Ocompany As SAPbobsCOM.Company
    Sub New(ByVal ocompany1 As SAPbobsCOM.Company, ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        Ocompany = ocompany1
    End Sub

    Public Sub Order_bind(ByVal oform As SAPbouiCOM.Form)
        Try
            ' oform.Freeze(True)
            CFL_BP_Supplier(oform, SBO_Application)
            oform.DataSources.UserDataSources.Add("oedit1", SAPbouiCOM.BoDataType.dt_DATE)
            oform.DataSources.UserDataSources.Add("oedit2", SAPbouiCOM.BoDataType.dt_DATE)
            oform.DataSources.UserDataSources.Add("oedit3", SAPbouiCOM.BoDataType.dt_DATE)
            oform.DataSources.UserDataSources.Add("oedit4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("oedit5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("oedit6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("oedit7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("V_0", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("V_1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("V_2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("V_3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEdit = oform.Items.Item("4").Specific
            oEdit.DataBind.SetBound(True, "", "oedit1")
            '  oEdit.String = "31.12.12" 'Format(Now.Date, "dd/MM/yy")
            oEdit = oform.Items.Item("6").Specific
            oEdit.DataBind.SetBound(True, "", "oedit2")
            '  oEdit.String = "31.12.12" 'Format(Now.Date, "dd/MM/yy")
            oEdit = oform.Items.Item("8").Specific
            oEdit.DataBind.SetBound(True, "", "oedit3")
            ' oEdit.String = "01.01.13" 'Format(Now.Date, "dd/MM/yy")
            oEdit = oform.Items.Item("10").Specific
            oEdit.DataBind.SetBound(True, "", "oedit4")
            oEdit.ChooseFromListUID = "CFLBPV"
            oEdit.ChooseFromListAlias = "CardCode"
            oCombo = oform.Items.Item("12").Specific
            oCombo.DataBind.SetBound(True, "", "oedit5")
            ComboLoad_Group(oform, oCombo, Ocompany)
            oCombo = oform.Items.Item("14").Specific
            oCombo.DataBind.SetBound(True, "", "oedit6")
            ComboLoad_Brand(oform, oCombo, Ocompany)
            oCombo = oform.Items.Item("16").Specific
            oCombo.DataBind.SetBound(True, "", "oedit7")
            ComboLoad_Status(oform, oCombo, Ocompany)
            oMatrix = oform.Items.Item("17").Specific
            oColumns = oMatrix.Columns
            oColumn = oColumns.Item("V_0")
            oColumn.DataBind.SetBound(True, "", "V_0")
            oColumn = oColumns.Item("V_1")
            oColumn.DataBind.SetBound(True, "", "V_1")
            oColumn = oColumns.Item("V_2")
            oColumn.DataBind.SetBound(True, "", "V_2")
            oColumn = oColumns.Item("V_3")
            oColumn.DataBind.SetBound(True, "", "V_3")
            MatrixLoad()
            ' oform.Freeze(False)
            oform.DataSources.DataTables.Add("OWHS1")
        Catch ex As Exception
            'oform.Freeze(False)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub MatrixLoad()
        Dim sqlstr As String = "SELECT T1.[Location], T0.[WhsCode], T0.[WhsName] FROM OWHS T0 left join OLCT T1 on T1.[Code]= T0.[Location] WHERE T0.[Nettable] ='Y'"
        oForm.DataSources.DataTables.Add("OWHS")
        oForm.DataSources.DataTables.Item("OWHS").ExecuteQuery(sqlstr)
        oMatrix = oForm.Items.Item("17").Specific
        oMatrix.Clear()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        oColumns = oMatrix.Columns
        oForm.Items.Item("17").Specific.Columns.item("V_3").DataBind.Bind("OWHS", "Location")
        oForm.Items.Item("17").Specific.Columns.item("V_1").DataBind.Bind("OWHS", "WhsCode")
        oForm.Items.Item("17").Specific.Columns.item("V_0").DataBind.Bind("OWHS", "WhsName")

        oForm.Items.Item("17").Specific.Clear()
        oForm.Items.Item("17").Specific.LoadFromDataSource()
        oForm.Items.Item("17").Specific.AutoResizeColumns()

    End Sub
    Private Sub MatrixLoad_Yes()
        Dim sqlstr As String = "SELECT 'Y' as Yes,T1.[Location], T0.[WhsCode], T0.[WhsName] FROM OWHS T0 left join OLCT T1 on T1.[Code]= T0.[Location] WHERE T0.[Nettable] ='Y'"
        oForm.DataSources.DataTables.Item("OWHS1").ExecuteQuery(sqlstr)
        oMatrix = oForm.Items.Item("17").Specific
        oMatrix.Clear()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        oColumns = oMatrix.Columns
        oForm.Items.Item("17").Specific.Columns.item("V_2").DataBind.Bind("OWHS1", "Yes")
       
        oForm.Items.Item("17").Specific.Clear()
        oForm.Items.Item("17").Specific.LoadFromDataSource()
        oForm.Items.Item("17").Specific.AutoResizeColumns()

    End Sub
    Private Sub MatrixLoad_No()
        Dim sqlstr As String = "SELECT 'N' as Yes,T1.[Location], T0.[WhsCode], T0.[WhsName] FROM OWHS T0 left join OLCT T1 on T1.[Code]= T0.[Location] WHERE T0.[Nettable] ='Y'"
        oForm.DataSources.DataTables.Item("OWHS1").ExecuteQuery(sqlstr)
        oMatrix = oForm.Items.Item("17").Specific
        oMatrix.Clear()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        oColumns = oMatrix.Columns
        oForm.Items.Item("17").Specific.Columns.item("V_2").DataBind.Bind("OWHS1", "Yes")

        oForm.Items.Item("17").Specific.Clear()
        oForm.Items.Item("17").Specific.LoadFromDataSource()
        oForm.Items.Item("17").Specific.AutoResizeColumns()

    End Sub

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Then
            '  SBO_Application.StatusBar.SetText("Shuting Down K9 addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Windows.Forms.Application.Exit()
        End If

        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            ' SBO_Application.StatusBar.SetText("Shuting Down K9 addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Windows.Forms.Application.Exit()
        End If

        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then
            'SBO_Application.StatusBar.SetText("Shuting Down K9 addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Windows.Forms.Application.Exit()
        End If
    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try

            Try

                If pVal.FormUID = "Order_Replenishment_Report" Then
                    oForm = SBO_Application.Forms.Item("Order_Replenishment_Report")
                    If pVal.ItemUID = "19" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Before_Action = False And pVal.InnerEvent = False Then

                        Dim J As Integer
                        Count = Count + 1
                        If Bol = False Then
                            MatrixLoad_Yes()
                            'For J = 1 To oMatrix.RowCount
                            '    oCheck = oMatrix.Columns.Item("V_2").Cells.Item(J).Specific
                            '    oCheck.Checked = True
                            'Next
                            Bol = True
                        Else
                            MatrixLoad_No()
                            'For J = 1 To oMatrix.RowCount
                            '    oCheck = oMatrix.Columns.Item("V_2").Cells.Item(J).Specific
                            '    oCheck.Checked = False
                            'Next
                            Bol = False
                        End If

                    End If

                  
                End If

            Catch ex As Exception

            End Try

            If pVal.FormUID = "Order_Replenishment_Report" Then
                Try
                    oForm = SBO_Application.Forms.Item("Order_Replenishment_Report")
                    '---CFL
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Try
                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                If pVal.ItemUID = "10" Then
                                    Try
                                        oEdit = oForm.Items.Item("10").Specific
                                        oEdit.String = oDataTable.GetValue("CardCode", 0)
                                    Catch ex As Exception
                                    End Try
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    '-----End CFL
                    '----------Print in to Excel--
                    If pVal.ItemUID = "Print" And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Try

                            'Dim OD1 As Decimal = 5.59
                            'Dim ood1 As String = OD1.ToString
                            'Dim OF1 As Integer = OD1
                            'If OD1 > OF1 Then
                            '    OF1 = OF1 + 1
                            'End If
                            'Dim oF2 As Decimal = OD1
                            'Exit Sub

                            Dim SuppCode As String = "%"
                            Dim Group As String = "%"
                            Dim Brand As String = "%"
                            Dim Status As String = "%"
                            oEdit = oForm.Items.Item("4").Specific
                            If oEdit.String = "" Then
                                SBO_Application.StatusBar.SetText("Plz Enter SOH Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim SOHdt As String = oEdit.Value
                            oEdit = oForm.Items.Item("6").Specific
                            If oEdit.String = "" Then
                                SBO_Application.StatusBar.SetText("Plz Enter ST from Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim STfdt As String = oEdit.Value

                            oEdit = oForm.Items.Item("8").Specific
                            If oEdit.String = "" Then
                                SBO_Application.StatusBar.SetText("Plz Enter ST To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim STtdt As String = oEdit.Value


                            oEdit = oForm.Items.Item("10").Specific
                            SuppCode = oEdit.String
                            If SuppCode = "" Then
                                SuppCode = "%"
                            End If
                            oCombo = oForm.Items.Item("12").Specific
                            Try
                                Group = oCombo.Selected.Value
                                If Group = "" Then
                                    oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@GROUP]  T0")
                                    Dim M As Integer = 0
                                    For M = 1 To oRecordSet.RecordCount
                                        If Group = "" Then
                                            Group = "'" & oRecordSet.Fields.Item("Code").Value & "'"
                                        Else
                                            Group = "'" & oRecordSet.Fields.Item("Code").Value & "'" & "," & Group
                                        End If
                                    Next
                                Else
                                    Group = "'" & Group & "'"
                                End If
                            Catch ex As Exception
                                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@GROUP]  T0")
                                Dim M As Integer = 0
                                For M = 1 To oRecordSet.RecordCount
                                    If Group = "" Then
                                        Group = "'" & oRecordSet.Fields.Item("Code").Value & "'"
                                    Else
                                        Group = "'" & oRecordSet.Fields.Item("Code").Value & "'" & "," & Group
                                    End If
                                Next
                            End Try


                            oCombo = oForm.Items.Item("14").Specific
                            Try
                                Brand = oCombo.Selected.Value
                                If Brand = "" Then
                                    Brand = "%"
                                End If
                            Catch ex As Exception
                                Brand = "%"
                            End Try
                            oCombo = oForm.Items.Item("16").Specific
                            Try
                                Status = oCombo.Selected.Value
                                If Status = "" Then
                                    Status = "%"
                                End If
                            Catch ex As Exception
                                Status = "%"
                            End Try
                            Dim J As Integer = 0
                            Dim WhscSQLCode As String = ""
                            Dim WhscCount As Integer = 0
                            For J = 1 To oMatrix.RowCount
                                oCheck = oMatrix.Columns.Item("V_2").Cells.Item(J).Specific
                                If oCheck.Checked = True Then
                                    WhscCount = WhscCount + 1
                                    oEdit = oMatrix.Columns.Item("V_1").Cells.Item(J).Specific
                                    If WhscSQLCode = "" Then
                                        WhscSQLCode = "'" & oEdit.String & "'"
                                    Else
                                        WhscSQLCode = "'" & oEdit.String & "'" & "," & WhscSQLCode
                                    End If
                                End If
                            Next
                            If WhscSQLCode = "" Then
                                SBO_Application.StatusBar.SetText("Plz Select The WareHouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim CardCode As String = ""
                            Dim CardName As String = ""
                            Dim U_Group As String = ""
                            Dim U_Family As String = ""
                            Dim U_Model As String = ""
                            Dim U_Screensize As String = ""
                            Dim U_HDCapacity As String = ""
                            Dim U_Generation As String = ""
                            Dim U_Network As String = ""
                            Dim U_Kind As String = ""
                            Dim U_ForModel As String = ""
                            Dim U_Watt As String = ""
                            Dim U_type As String = ""
                            Dim U_Specification As String = ""
                            Dim U_Fits As String = ""
                            Dim U_Color As String = ""
                            Dim U_Brand As String = ""
                            Dim U_UPC As String = ""
                            Dim U_AltCode1 As String = ""
                            Dim U_AltCode2 As String = ""
                            Dim U_AltCode3 As String = ""
                            Dim SuppCatNum As String = ""
                            Dim ItemName As String = ""
                            Dim U_Status As String = ""
                            Dim CreateDate As String = ""
                            Dim U_Processor As String = ""
                            Dim U_RAMSize As String = ""
                            Dim MinStock As String = ""
                            Dim MinOrdrQty As Integer = 0
                            Dim Currency As String = ""
                            Dim COST As String = ""
                            Dim Retail_price_with_GST As String = ""
                            Dim Retail_price_with_OUT_GST As String = ""
                            Dim Margin_Amt As String = ""
                            Dim Margin_Per As String = ""
                            Dim U_MarginRank As String = ""
                            Dim ItemCode As String = ""
                            Dim U_Rebate As String = ""

                            '  oForm.DataSources.DataTables.Add("Elush")

                            Dim customerOrders As DataSet = New DataSet("CustomerOrders")


                            Dim table1 As DataTable
                            table1 = New DataTable("Elush")
                            table1.Columns.Add("CardCode")
                            table1.Columns.Add("CardName")
                            table1.Columns.Add("Group")
                            table1.Columns.Add("Family")
                            table1.Columns.Add("Model")
                            table1.Columns.Add("Screensize")
                            table1.Columns.Add("Processor")
                            table1.Columns.Add("RAMSize")
                            table1.Columns.Add("HDCapacity")
                            table1.Columns.Add("Generation")
                            table1.Columns.Add("Network")
                            table1.Columns.Add("Kind")
                            table1.Columns.Add("ForModel")
                            table1.Columns.Add("Watt")
                            table1.Columns.Add("Type")
                            table1.Columns.Add("Specification")
                            table1.Columns.Add("Fits")
                            table1.Columns.Add("Color")
                            table1.Columns.Add("Brand")
                            ' table1.Columns.Add("UPC")

                            table1.Columns.Add("ItemCode")
                            table1.Columns.Add("AltCode1")
                            table1.Columns.Add("AltCode2")
                            table1.Columns.Add("AltCode3")
                            table1.Columns.Add("SuppCatNum")
                            table1.Columns.Add("ItemName")
                            table1.Columns.Add("Status")
                            table1.Columns.Add("CreateDate")

                            table1.Columns.Add("MinStock")
                            table1.Columns.Add("MinOrdrQty")
                            table1.Columns.Add("Currency")
                            table1.Columns.Add("Cost")
                            table1.Columns.Add("Retail_Price_With_GST")
                            table1.Columns.Add("Retail_Price_Without_GST")
                            table1.Columns.Add("Margin_Amt")
                            table1.Columns.Add("Margin_Per")
                            table1.Columns.Add("MarginRank")
                            table1.Columns.Add("Rebate")
                            'table1.Columns.Add("AltCode1")
                            ' table1.Columns.Add("ItemCode")

                            oRecordSet_WHSC = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet_WHSC.DoQuery("SELECT T0.[WhsCode] FROM OWHS T0 WHERE T0.[Nettable] ='Y' and T0.[WhsCode] in(" & WhscSQLCode & ")  ORDER BY T0.[WhsCode]")
                            Dim k As Integer = 0

                            Dim WhscCode(oRecordSet_WHSC.RecordCount - 1) As String
                            Dim soh(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim st(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim PO(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim SO(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim SO1(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim OD(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim OD1(oRecordSet_WHSC.RecordCount - 1) As Integer
                            Dim DOI(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim DOI1(oRecordSet_WHSC.RecordCount - 1) As Integer
                            For k = 1 To oRecordSet_WHSC.RecordCount
                                table1.Columns.Add("WHS_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("SOH_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("ST_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("PO_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("SO_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("OD_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("DOI_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                oRecordSet_WHSC.MoveNext()
                            Next

                            table1.Columns.Add("TOTSOH")
                            table1.Columns.Add("TOTST")
                            table1.Columns.Add("TOTPO")
                            table1.Columns.Add("TOTSO")
                            ''

                            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim str As String = "SELECT distinct (T0.[CardCode]), (select CardName from OCRD where CardCode=T0.CardCode) as CardName,  T0.[U_Group], T0.[U_Family], T0.[U_Model], T0.[U_Screensize]," & _
"T0.[U_HDCapacity], T0.[U_Generation], T0.[U_Network], T0.[U_Kind], T0.[U_ForModel], T0.[U_Watt], T0.[U_type], " & _
"T0.[U_Specification], T0.[U_Fits], T0.[U_Color], T0.[U_Brand],T0.[U_AltCode1], T0.[U_AltCode2]," & _
"T0.[U_AltCode3], T0.[SuppCatNum], Replace(T0.[ItemName],',',''), T0.[U_Status], T0.[CreateDate], T0.[U_Processor]," & _
"T0.[U_RAMSize], T2.[MinStock], T0.[MinOrdrQty], (select [Currency] from OCRD where CardCode=T0.CardCode) as 'Currency',(selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='2') AS COST,(selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') AS 'Retail price with GST' ," & _
"(selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') AS 'Retail price with OUT GST' ," & _
"((selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') -(selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='2')) AS 'Margin $'," & _
" CASE when (selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') =0" & _
" then (Select '0')" & _
" else ((1-((selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='2')/(selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1')))) " & _
" end AS 'Margin %', T0.U_MarginRank, T0.ItemCode , T8.WhsCode," & _
"(SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode) as 'SOH'," & _
"(SELECT SUM(OUTQTY - INQTY ) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype in('13','14')) As 'ST'," & _
"(SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode) as 'PO'," & _
"(CASE  when cast ((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') AS Int)=0 Then (Select '0' as 'SO')" & _
"else " & _
"(((SELECT SUM(OUTQTY - INQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and '" & STtdt & "' AND Warehouse=T8.WhsCode and transtype in ('13','14'))/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')) * (SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode))-(SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode)-(SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode)" & _
"End) As 'SO'," & _
" case when cast ((select  cast(datediff(dd, '" & STfdt & "', '" & STtdt & "') as varchar) as 'M') AS Int)=0 or ((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and '" & STtdt & "' AND Warehouse=T8.WhsCode) /(select  cast(datediff(dd, '" & STfdt & "', '" & STtdt & "') as varchar) as 'M'))=0 then " & _
" (select '0' as 'DOI') else " & _
" ((SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode)  + " & _
" (CASE  when cast ((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') AS Int)=0 Then (Select '0' as 'SO') " & _
" else " & _
" (((SELECT SUM(OUTQTY-INQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and '" & STtdt & "' AND Warehouse=T8.WhsCode and transtype in ('13','14'))/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')) * (SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode))-(SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode)-(SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode) " & _
" End) + " & _
" (SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode))/ (((SELECT SUM(OUTQTY-INQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype in ('13','14')))/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')) " & _
" End as 'DOI', " & _
"T0.U_Rebate FROM [dbo].[OITM]  T0 , [dbo].[OITW]  T2 , [OWHS] T8" & _
 " WHERE T0.ItemCode = T2.ItemCode AND T2.WhsCode='HO' and T0.InvntItem='Y' and " & _
" isnull(T0.CardCode,0) like '" & SuppCode & "' and isnull(T0.U_Group,0) in(" & Group & ") AND  isnull(T0.U_Brand,0)  LIKE '" & Brand & "' AND isnull(T0.U_Status,0)  LIKE '" & Status & "' and T8.WhsCode in (" & WhscSQLCode & ")  ORDER BY T0.itemcode, T8.[WhsCode]"
                            ' and  (t0.ItemCode='I0001' or t0.ItemCode='I0002') 
                            '"(((SELECT T10.[OnHand] FROM OITW T10 WHERE T10.[ItemCode] =T0.[ItemCode] AND  T10.[WhsCode] =T8.WhsCode)+" & _
                            '" (CASE when cast ((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') AS Int)=0" & _
                            '" Then (Select '0' as 'SO') else CASE" & _
                            '" when cast(((((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13')/((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')))*(SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode))) AS int) > T0.[MinOrdrQty] " & _
                            '" then" & _
                            '" ((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13')/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M'))*(SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode)" & _
                            '" else" & _
                            '" T0.[MinOrdrQty] End " & _
                            '" End)" & _
                            '" )" & _
                            '" /((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13') /(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M'))" & _
                            '" )" & _
                            '  " as 'DOI'" & _

                            oRecordSet.DoQuery(str)
                            Dim i As Integer = 1
                            Dim k1 As Integer = oRecordSet.RecordCount / oRecordSet_WHSC.RecordCount
                            For i = 1 To k1
                                Dim Icode As String = oRecordSet.Fields.Item(0).Value.ToString
                                SBO_Application.StatusBar.SetText("" & k1 & "  -of  " & i & "-Row Processing.Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                CardCode = oRecordSet.Fields.Item(0).Value.ToString
                                CardName = oRecordSet.Fields.Item(1).Value.ToString
                                U_Group = oRecordSet.Fields.Item(2).Value.ToString
                                U_Family = oRecordSet.Fields.Item(3).Value.ToString
                                U_Model = oRecordSet.Fields.Item(4).Value.ToString
                                U_Screensize = oRecordSet.Fields.Item(5).Value.ToString
                                U_HDCapacity = oRecordSet.Fields.Item(6).Value.ToString
                                U_Generation = oRecordSet.Fields.Item(7).Value.ToString
                                U_Network = oRecordSet.Fields.Item(8).Value.ToString
                                U_Kind = oRecordSet.Fields.Item(9).Value.ToString
                                U_ForModel = oRecordSet.Fields.Item(10).Value.ToString
                                U_Watt = oRecordSet.Fields.Item(11).Value.ToString
                                U_type = oRecordSet.Fields.Item(12).Value.ToString
                                U_Specification = oRecordSet.Fields.Item(13).Value.ToString
                                U_Fits = oRecordSet.Fields.Item(14).Value.ToString
                                U_Color = oRecordSet.Fields.Item(15).Value.ToString
                                U_Brand = oRecordSet.Fields.Item(16).Value.ToString
                                ' U_UPC = oRecordSet.Fields.Item(17).Value.ToString
                                U_AltCode1 = oRecordSet.Fields.Item(17).Value.ToString
                                U_AltCode2 = oRecordSet.Fields.Item(18).Value.ToString
                                U_AltCode3 = oRecordSet.Fields.Item(19).Value.ToString
                                SuppCatNum = oRecordSet.Fields.Item(20).Value.ToString
                                ItemName = oRecordSet.Fields.Item(21).Value.ToString
                                U_Status = oRecordSet.Fields.Item(22).Value.ToString
                                CreateDate = Format(oRecordSet.Fields.Item(23).Value, "dd/MM/yyyy")
                                U_Processor = oRecordSet.Fields.Item(24).Value.ToString
                                U_RAMSize = oRecordSet.Fields.Item(25).Value.ToString
                                MinStock = oRecordSet.Fields.Item(26).Value.ToString
                                MinOrdrQty = oRecordSet.Fields.Item(27).Value
                                Currency = oRecordSet.Fields.Item(28).Value.ToString
                                COST = oRecordSet.Fields.Item(29).Value.ToString
                                Retail_price_with_GST = oRecordSet.Fields.Item(30).Value.ToString
                                Retail_price_with_OUT_GST = oRecordSet.Fields.Item(31).Value.ToString
                                Margin_Amt = oRecordSet.Fields.Item(32).Value.ToString
                                Margin_Per = oRecordSet.Fields.Item(33).Value.ToString
                                U_MarginRank = oRecordSet.Fields.Item(34).Value.ToString
                                ItemCode = oRecordSet.Fields.Item(35).Value.ToString
                                U_Rebate = oRecordSet.Fields.Item(42).Value.ToString
                                'oRecordSet_WHSC = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'oRecordSet_WHSC.DoQuery("SELECT T0.[WhsCode] FROM OWHS T0 WHERE T0.[Nettable] ='Y' ORDER BY T0.[WhsCode]")
                                If oRecordSet_WHSC.RecordCount > 0 Then
                                    oRecordSet_WHSC.MoveFirst()
                                    For J = 1 To oRecordSet_WHSC.RecordCount
                                        soh(J - 1) = oRecordSet.Fields.Item(37).Value.ToString
                                        st(J - 1) = oRecordSet.Fields.Item(38).Value.ToString
                                        PO(J - 1) = oRecordSet.Fields.Item(39).Value.ToString
                                        SO(J - 1) = oRecordSet.Fields.Item(40).Value.ToString
                                        If SO(J - 1) < 0 Then
                                            SO1(J - 1) = 0
                                        Else
                                            SO1(J - 1) = SO(J - 1)
                                        End If
                                        OD(J - 1) = oRecordSet.Fields.Item(40).Value.ToString
                                        OD1(J - 1) = OD(J - 1)
                                        If OD(J - 1) > OD1(J - 1) Then
                                            OD1(J - 1) = OD1(J - 1) + 1
                                        End If
                                        If OD1(J - 1) < 0 Then
                                            OD1(J - 1) = 0
                                        End If
                                        DOI(J - 1) = oRecordSet.Fields.Item(41).Value.ToString
                                        DOI1(J - 1) = DOI(J - 1)
                                        If DOI(J - 1) > DOI1(J - 1) Then
                                            DOI1(J - 1) = DOI1(J - 1) + 1
                                        End If
                                        If DOI1(J - 1) < 0 Then
                                            DOI1(J - 1) = 0
                                        End If
                                        WhscCode(J - 1) = oRecordSet.Fields.Item(36).Value.ToString
                                        oRecordSet.MoveNext()
                                    Next
                                End If

                                Dim dr As DataRow
                                dr = table1.NewRow()
                                dr.Item(0) = CardCode
                                dr.Item(1) = CardName
                                dr.Item(2) = U_Group
                                dr.Item(3) = U_Family
                                dr.Item(4) = U_Model
                                dr.Item(5) = U_Screensize
                                dr.Item(6) = U_Processor
                                dr.Item(7) = U_RAMSize
                                dr.Item(8) = U_HDCapacity
                                dr.Item(9) = U_Generation
                                dr.Item(10) = U_Network
                                dr.Item(11) = U_Kind
                                dr.Item(12) = U_ForModel
                                dr.Item(13) = U_Watt
                                dr.Item(14) = U_type
                                dr.Item(15) = U_Specification
                                dr.Item(16) = U_Fits
                                dr.Item(17) = U_Color
                                dr.Item(18) = U_Brand
                                ' dr.Item(17) = U_UPC

                                dr.Item(19) = ItemCode
                                dr.Item(20) = U_AltCode1
                                dr.Item(21) = U_AltCode2
                                dr.Item(22) = U_AltCode3
                                dr.Item(23) = SuppCatNum
                                dr.Item(24) = ItemName
                                dr.Item(25) = U_Status
                                dr.Item(26) = CreateDate

                                dr.Item(27) = MinStock
                                dr.Item(28) = MinOrdrQty
                                dr.Item(29) = Currency
                                dr.Item(30) = COST
                                dr.Item(31) = Retail_price_with_GST
                                dr.Item(32) = Retail_price_with_OUT_GST
                                dr.Item(33) = Margin_Amt
                                dr.Item(34) = Margin_Per
                                dr.Item(35) = U_MarginRank
                                dr.Item(36) = U_Rebate
                                Dim v As Integer = 0
                                Dim m As Integer = 0
                                m = 1
                                Dim TOTSOH As Integer = 0
                                Dim TOTST As Integer = 0
                                Dim TOTPO As Integer = 0
                                Dim TOTSO As Decimal = 0

                                For v = 1 To oRecordSet_WHSC.RecordCount
                                    dr.Item(36 + m) = WhscCode(v - 1)
                                    dr.Item(36 + m + 1) = soh(v - 1)
                                    TOTSOH = TOTSOH + soh(v - 1)
                                    dr.Item(36 + m + 2) = st(v - 1)
                                    TOTST = TOTST + st(v - 1)
                                    dr.Item(36 + m + 3) = PO(v - 1)
                                    TOTPO = TOTPO + PO(v - 1)
                                    dr.Item(36 + m + 4) = SO1(v - 1)
                                    TOTSO = TOTSO + SO(v - 1)
                                    dr.Item(36 + m + 5) = OD1(v - 1)
                                    dr.Item(36 + m + 6) = DOI1(v - 1)
                                    m = m + 7
                                Next
                                Dim whcount As Integer = oRecordSet_WHSC.RecordCount
                                Dim rno As Integer = 37 + (whcount * 7)
                                dr.Item(rno) = TOTSOH
                                dr.Item(rno + 1) = TOTST
                                dr.Item(rno + 2) = TOTPO
                                dr.Item(rno + 3) = TOTSO
                                table1.Rows.Add(dr)

                            Next

                            Try

                                ExportToExcel(table1, WhscCount)


                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                ' SBO_Application.MessageBox(ex.Message)
                            End Try
                            'Try



                            '    Dim Excel As New Microsoft.Office.Interop.Excel.Application
                            '    Dim WorkBooks1 As Microsoft.Office.Interop.Excel.Workbooks
                            '    Dim WorkBook1 As Microsoft.Office.Interop.Excel.Workbook
                            '    WorkBooks1 = Excel.Workbooks

                            '    'Usage
                            '    ' Dim instance As XlFileFormat
                            '    WorkBook1 = WorkBooks1.Open("C:\ElushOrder" & Format(Now.Date, "yyyyMMdd") & ".txt")
                            '    'Dim XlFileFormat As String = "xlsx"
                            '    ' XlFileFormat.xlWorkbookDefault = 51 - this is correct, although a lot of forums recommend
                            '    WorkBook1.SaveAs(Filename:="C:\ElushOrder" & Format(Now.Date, "yyyyMMdd") & "", FileFormat:=XlFileFormat.xlExcel8)

                            '    'CLEANUP

                            '    WorkBook1.Save() 'trying to find a way to get this to close without a prompt
                            '    WorkBook1.Close(False)
                            '    WorkBooks1.Close()
                            '    Excel.Quit()
                            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(WorkBook1)
                            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(WorkBooks1)
                            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
                            '    WorkBook1 = Nothing
                            '    WorkBooks1 = Nothing
                            '    Excel = Nothing

                            'Catch ex As Exception

                            'End Try
                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                    End If

                    '----------End Print in to Excel--
                    '------------Print into CSV----------
                    If pVal.ItemUID = "20" And pVal.Before_Action = True And pVal.InnerEvent = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Try

                            Dim SuppCode As String = "%"
                            Dim Group As String = "%"
                            Dim Brand As String = "%"
                            Dim Status As String = "%"
                            oEdit = oForm.Items.Item("4").Specific
                            If oEdit.String = "" Then
                                SBO_Application.StatusBar.SetText("Plz Enter SOH Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim SOHdt As String = oEdit.Value
                            oEdit = oForm.Items.Item("6").Specific
                            If oEdit.String = "" Then
                                SBO_Application.StatusBar.SetText("Plz Enter ST from Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim STfdt As String = oEdit.Value

                            oEdit = oForm.Items.Item("8").Specific
                            If oEdit.String = "" Then
                                SBO_Application.StatusBar.SetText("Plz Enter ST To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim STtdt As String = oEdit.Value


                            oEdit = oForm.Items.Item("10").Specific
                            SuppCode = oEdit.String
                            If SuppCode = "" Then
                                SuppCode = "%"
                            End If
                            oCombo = oForm.Items.Item("12").Specific
                            Try
                                Group = oCombo.Selected.Value
                                If Group = "" Then
                                    oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@GROUP]  T0")
                                    Dim M As Integer = 0
                                    For M = 1 To oRecordSet.RecordCount
                                        If Group = "" Then
                                            Group = "'" & oRecordSet.Fields.Item("Code").Value & "'"
                                        Else
                                            Group = "'" & oRecordSet.Fields.Item("Code").Value & "'" & "," & Group
                                        End If
                                    Next
                                Else
                                    Group = "'" & Group & "'"
                                End If
                            Catch ex As Exception
                                oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@GROUP]  T0")
                                Dim M As Integer = 0
                                For M = 1 To oRecordSet.RecordCount
                                    If Group = "" Then
                                        Group = "'" & oRecordSet.Fields.Item("Code").Value & "'"
                                    Else
                                        Group = "'" & oRecordSet.Fields.Item("Code").Value & "'" & "," & Group
                                    End If
                                Next
                            End Try


                            oCombo = oForm.Items.Item("14").Specific
                            Try
                                Brand = oCombo.Selected.Value
                                If Brand = "" Then
                                    Brand = "%"
                                End If
                            Catch ex As Exception
                                Brand = "%"
                            End Try
                            oCombo = oForm.Items.Item("16").Specific
                            Try
                                Status = oCombo.Selected.Value
                                If Status = "" Then
                                    Status = "%"
                                End If
                            Catch ex As Exception
                                Status = "%"
                            End Try
                            Dim J As Integer = 0
                            Dim WhscSQLCode As String = ""
                            Dim WhscCount As Integer = 0
                            For J = 1 To oMatrix.RowCount
                                oCheck = oMatrix.Columns.Item("V_2").Cells.Item(J).Specific
                                If oCheck.Checked = True Then
                                    WhscCount = WhscCount + 1
                                    oEdit = oMatrix.Columns.Item("V_1").Cells.Item(J).Specific
                                    If WhscSQLCode = "" Then
                                        WhscSQLCode = "'" & oEdit.String & "'"
                                    Else
                                        WhscSQLCode = "'" & oEdit.String & "'" & "," & WhscSQLCode
                                    End If
                                End If
                            Next
                            If WhscSQLCode = "" Then
                                SBO_Application.StatusBar.SetText("Plz Select The WareHouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            Dim CardCode As String = ""
                            Dim CardName As String = ""
                            Dim U_Group As String = ""
                            Dim U_Family As String = ""
                            Dim U_Model As String = ""
                            Dim U_Screensize As String = ""
                            Dim U_HDCapacity As String = ""
                            Dim U_Generation As String = ""
                            Dim U_Network As String = ""
                            Dim U_Kind As String = ""
                            Dim U_ForModel As String = ""
                            Dim U_Watt As String = ""
                            Dim U_type As String = ""
                            Dim U_Specification As String = ""
                            Dim U_Fits As String = ""
                            Dim U_Color As String = ""
                            Dim U_Brand As String = ""
                            Dim U_UPC As String = ""
                            Dim U_AltCode1 As String = ""
                            Dim U_AltCode2 As String = ""
                            Dim U_AltCode3 As String = ""
                            Dim SuppCatNum As String = ""
                            Dim ItemName As String = ""
                            Dim U_Status As String = ""
                            Dim CreateDate As String = ""
                            Dim U_Processor As String = ""
                            Dim U_RAMSize As String = ""
                            Dim MinStock As String = ""
                            Dim MinOrdrQty As Integer = 0
                            Dim Currency As String = ""
                            Dim COST As String = ""
                            Dim Retail_price_with_GST As String = ""
                            Dim Retail_price_with_OUT_GST As String = ""
                            Dim Margin_Amt As String = ""
                            Dim Margin_Per As String = ""
                            Dim U_MarginRank As String = ""
                            Dim ItemCode As String = ""
                            Dim U_Rebate As String = ""

                            '  oForm.DataSources.DataTables.Add("Elush")

                            Dim customerOrders As DataSet = New DataSet("CustomerOrders")


                            Dim table1 As DataTable
                            table1 = New DataTable("Elush")
                            table1.Columns.Add("CardCode")
                            table1.Columns.Add("CardName")
                            table1.Columns.Add("Group")
                            table1.Columns.Add("Family")
                            table1.Columns.Add("Model")
                            table1.Columns.Add("Screensize")
                            table1.Columns.Add("Processor")
                            table1.Columns.Add("RAMSize")
                            table1.Columns.Add("HDCapacity")
                            table1.Columns.Add("Generation")
                            table1.Columns.Add("Network")
                            table1.Columns.Add("Kind")
                            table1.Columns.Add("ForModel")
                            table1.Columns.Add("Watt")
                            table1.Columns.Add("Type")
                            table1.Columns.Add("Specification")
                            table1.Columns.Add("Fits")
                            table1.Columns.Add("Color")
                            table1.Columns.Add("Brand")
                            ' table1.Columns.Add("UPC")
                            table1.Columns.Add("ItemCode")
                            table1.Columns.Add("AltCode1")
                            table1.Columns.Add("AltCode2")
                            table1.Columns.Add("AltCode3")
                            table1.Columns.Add("SuppCatNum")
                            table1.Columns.Add("ItemName")
                            table1.Columns.Add("Status")
                            table1.Columns.Add("CreateDate")

                            table1.Columns.Add("MinStock")
                            table1.Columns.Add("MinOrdrQty")
                            table1.Columns.Add("Currency")
                            table1.Columns.Add("Cost")
                            table1.Columns.Add("Retail_Price_With_GST")
                            table1.Columns.Add("Retail_Price_Without_GST")
                            table1.Columns.Add("Margin_Amt")
                            table1.Columns.Add("Margin_Per")
                            table1.Columns.Add("MarginRank")
                            table1.Columns.Add("Rebate")
                            '
                            'table1.Columns.Add("ItemCode")

                            oRecordSet_WHSC = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet_WHSC.DoQuery("SELECT T0.[WhsCode] FROM OWHS T0 WHERE T0.[Nettable] ='Y' and T0.[WhsCode] in(" & WhscSQLCode & ")  ORDER BY T0.[WhsCode]")
                            Dim k As Integer = 0

                            Dim WhscCode(oRecordSet_WHSC.RecordCount - 1) As String
                            Dim soh(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim st(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim PO(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim SO(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim SO1(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim OD(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim OD1(oRecordSet_WHSC.RecordCount - 1) As Integer
                            Dim DOI(oRecordSet_WHSC.RecordCount - 1) As Decimal
                            Dim DOI1(oRecordSet_WHSC.RecordCount - 1) As Integer
                            Dim DOIUDF As Decimal = 0.0
                            For k = 1 To oRecordSet_WHSC.RecordCount
                                table1.Columns.Add("WHS_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("SOH_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("ST_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("PO_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("SO_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("OD_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                table1.Columns.Add("DOI_" & oRecordSet_WHSC.Fields.Item(0).Value & "")
                                oRecordSet_WHSC.MoveNext()
                            Next
                            table1.Columns.Add("TOTSOH")
                            table1.Columns.Add("TOTST")
                            table1.Columns.Add("TOTPO")
                            table1.Columns.Add("TOTSO")
                            table1.Columns.Add("TOTDOI")

                            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            '=============================

                            Dim str As String = ""
                            str = "SELECT SUM(T00.INQTY - T00.OUTQTY) As SOH,T00.Warehouse,T00.ItemCode INTO #SOH  " & _
"FROM OINM T00 with (nolock) " & _
"inner Join OITM T0 with (nolock) on T0.ItemCode=T00.ItemCode " & _
"WHERE T00.DOCDATE <= '" & SOHdt & "'  " & _
"and   isnull(T0.CardCode,0) like '" & SuppCode & "' and isnull(T0.U_Group,0) in(" & Group & ") AND  isnull(T0.U_Brand,0) " & _
"LIKE '" & Brand & "' AND isnull(T0.U_Status,0)  LIKE '" & Status & "' and T00.Warehouse in  (" & WhscSQLCode & ")  " & _
"Group By T00.Warehouse,T00.ItemCode; " & _
"SELECT ABS(SUM(T00.INQTY - T00.OUTQTY)) As ST,T00.Warehouse,T00.ItemCode INTO #ST " & _
"FROM OINM T00 with (nolock) " & _
"inner Join OITM T0 with (nolock) on T0.ItemCode=T00.ItemCode " & _
"WHERE T00.DOCDATE between  '" & STfdt & "' and '" & STtdt & "' and T00.transtype in ('13','14') and " & _
"isnull(T0.CardCode,0) like '" & SuppCode & "' and isnull(T0.U_Group,0) in(" & Group & ") AND  isnull(T0.U_Brand,0)   " & _
"LIKE '" & Brand & "' AND isnull(T0.U_Status,0)  LIKE '" & Status & "' and T00.Warehouse in  (" & WhscSQLCode & ")   " & _
"Group By T00.Warehouse,T00.ItemCode; " & _
"SELECT SUM(T00.OnOrder) As PO,T00.WhsCode,T00.ItemCode INTO #PO " & _
"FROM OITW T00 with (nolock) " & _
"inner Join OITM T0 with (nolock) on T0.ItemCode=T00.ItemCode " & _
"WHERE isnull(T0.CardCode,0) like '" & SuppCode & "' and isnull(T0.U_Group,0) in(" & Group & ") AND  isnull(T0.U_Brand,0) " & _
"LIKE '" & Brand & "' AND isnull(T0.U_Status,0)  LIKE '" & Status & "' and T00.WhsCode in  (" & WhscSQLCode & ")   " & _
"Group By T00.WhsCode,T00.ItemCode; " & _
"SELECT distinct (T0.[CardCode]), OC.CardName as CardName,  " & _
"T0.[U_Group], T0.[U_Family], T0.[U_Model], T0.[U_Screensize],T0.[U_HDCapacity], T0.[U_Generation],  " & _
"T0.[U_Network], T0.[U_Kind], T0.[U_ForModel], T0.[U_Watt], T0.[U_type], T0.[U_Specification], T0.[U_Fits], " & _
"T0.[U_Color], T0.[U_Brand],T0.[U_AltCode1], T0.[U_AltCode2],T0.[U_AltCode3], T0.[SuppCatNum],  " & _
"Replace(T0.[ItemName],',',''), T0.[U_Status], T0.[CreateDate], T0.[U_Processor],T0.[U_RAMSize],  " & _
"T2.[MinStock], T0.[MinOrdrQty],OC.Currency as 'Currency',T4.Price AS COST,T41.PRICE  AS 'Retail price with GST' ,(T41.PRICE/1.07) AS 'Retail price with OUT GST' ,((T41.PRICE/1.07) - T4.Price) As 'Margin $', " & _
"CASE when (T41.PRICE/1.07) =0 then   0  else  (1-(T4.Price/(T41.PRICE/1.07))) end  AS 'Margin %', T0.U_MarginRank, T0.ItemCode , T8.WhsCode, " & _
"T43.SOH as 'SOH',T44.ST as 'ST',T45.PO as 'PO', " & _
"(CASE  when cast ((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') AS Int)=0 Then (Select '0' as 'SO')   " & _
"else   " & _
"(((T44.ST)/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')) *   " & _
"(SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode))-  " & _
"(Isnull(T43.SOH,0))-(Isnull(T45.PO,0))   " & _
"End) As 'SO',  " & _
" (select ''), " & _
"T0.U_Rebate,(SELECT top 1 T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group]) as DOI1, " & _
"(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')  " & _
"FROM [dbo].[OITM]  T0  " & _
"Left Join OCRD OC on T0.CardCode=OC.CardCode " & _
"LEFT Join ITM1 T4 ON T4.ItemCode=T0.ItemCode and T4.PriceList='2' " & _
"LEFT Join ITM1 T41 ON T41.ItemCode=T0.ItemCode and T41.PriceList='1' " & _
"LEFT Join [dbo].[OITW]  T2  ON T0.ItemCode = T2.ItemCode and T2.WhsCode='HO' " & _
"LEFT Join [OWHS] T8  ON  T0.ItemCode = T2.ItemCode " & _
"LEFT Join #SOH T43 ON T43.ItemCode=T0.ItemCode and T43.Warehouse=T8.WhsCode  " & _
"LEFT Join #ST T44 ON T44.ItemCode=T0.ItemCode and T44.Warehouse=T8.WhsCode " & _
"LEFT Join #PO T45 ON T45.ItemCode=T0.ItemCode and T45.WhsCode=T8.WhsCode  " & _
"WHERE   T0.InvntItem='Y' and " & _
"isnull(T0.CardCode,0) like '" & SuppCode & "' and isnull(T0.U_Group,0) in(" & Group & ") AND  isnull(T0.U_Brand,0) " & _
"LIKE '" & Brand & "' AND isnull(T0.U_Status,0)  LIKE '" & Status & "' and T8.WhsCode in  (" & WhscSQLCode & ") " & _
"ORDER BY T0.itemcode, T8.[WhsCode]; " & _
"Drop Table #SOH; " & _
"Drop Table #ST; " & _
"Drop Table #PO;"

                            '                            '***************************************
                            '                            Dim str As String = "SELECT distinct (T0.[CardCode]), (select CardName from OCRD where CardCode=T0.CardCode) as CardName,  T0.[U_Group], T0.[U_Family], T0.[U_Model], T0.[U_Screensize]," & _
                            '"T0.[U_HDCapacity], T0.[U_Generation], T0.[U_Network], T0.[U_Kind], T0.[U_ForModel], T0.[U_Watt], T0.[U_type], " & _
                            '"T0.[U_Specification], T0.[U_Fits], T0.[U_Color], T0.[U_Brand],T0.[U_AltCode1], T0.[U_AltCode2]," & _
                            '"T0.[U_AltCode3], T0.[SuppCatNum], Replace(T0.[ItemName],',',''), T0.[U_Status], T0.[CreateDate], T0.[U_Processor]," & _
                            '"T0.[U_RAMSize], T2.[MinStock], T0.[MinOrdrQty], (select [Currency] from OCRD where CardCode=T0.CardCode) as 'Currency',(selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='2') AS COST,(selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') AS 'Retail price with GST' ," & _
                            '"(selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') AS 'Retail price with OUT GST' ," & _
                            '"((selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') -(selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='2')) AS 'Margin $'," & _
                            '" CASE when (selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1') =0" & _
                            '" then (Select '0')" & _
                            '" else ((1-((selECT T4.PRICE FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='2')/(selECT (T4.PRICE/1.07) FROM ITM1 T4 WHERE T4.ItemCode=T0.ItemCode AND T4.PriceList='1')))) " & _
                            '" end AS 'Margin %', T0.U_MarginRank, T0.ItemCode , T8.WhsCode," & _
                            '"(SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode) as 'SOH'," & _
                            '"(SELECT SUM(OUTQTY-INQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype in ('13','14')) As 'ST'," & _
                            '"(SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode) as 'PO'," & _
                            '"(CASE  when cast ((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') AS Int)=0 Then (Select '0' as 'SO')" & _
                            '"else " & _
                            '"(((SELECT SUM(OUTQTY-INQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and '" & STtdt & "' AND Warehouse=T8.WhsCode and transtype in ('13','14'))/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')) * (SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode))-(SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode)-(SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode)" & _
                            '"End) As 'SO',(select '')," & _
                            '"T0.U_Rebate,(SELECT top 1 T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group]) as DOI1,(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') FROM [dbo].[OITM]  T0 , [dbo].[OITW]  T2 , [OWHS] T8" & _
                            ' " WHERE T0.ItemCode = T2.ItemCode AND T2.WhsCode='HO' and T0.InvntItem='Y' and " & _
                            '" isnull(T0.CardCode,0) like '" & SuppCode & "' and isnull(T0.U_Group,0) in(" & Group & ") AND  isnull(T0.U_Brand,0)  LIKE '" & Brand & "' AND isnull(T0.U_Status,0)  LIKE '" & Status & "' and T8.WhsCode in (" & WhscSQLCode & ")  ORDER BY T0.itemcode, T8.[WhsCode]"

                            '***************************************************************
                            ' and  (t0.ItemCode='I0001' or t0.ItemCode='I0002') 
                            '"(((SELECT T10.[OnHand] FROM OITW T10 WHERE T10.[ItemCode] =T0.[ItemCode] AND  T10.[WhsCode] =T8.WhsCode)+" & _
                            '" (CASE when cast ((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') AS Int)=0" & _
                            '" Then (Select '0' as 'SO') else CASE" & _
                            '" when cast(((((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13')/((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')))*(SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode))) AS int) > T0.[MinOrdrQty] " & _
                            '" then" & _
                            '" ((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13')/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M'))*(SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode)" & _
                            '" else" & _
                            '" T0.[MinOrdrQty] End " & _
                            '" End)" & _
                            '" )" & _
                            '" /((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13') /(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M'))" & _
                            '" )" & _
                            '  " as 'DOI'" & _


                            '                            " case when cast ((select  cast(datediff(dd, '" & STfdt & "', '" & STtdt & "') as varchar) as 'M') AS Int)=0 or ((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and '" & STtdt & "' AND Warehouse=T8.WhsCode) /(select  cast(datediff(dd, '" & STfdt & "', '" & STtdt & "') as varchar) as 'M'))=0 then " & _
                            '" (select '0' as 'DOI') else " & _
                            '" ((SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode)  + " & _
                            '" (CASE  when cast ((select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M') AS Int)=0 Then (Select '0' as 'SO') " & _
                            '" else " & _
                            '" (((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and '" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13')/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')) * (SELECT T9.[U_DOI] FROM [dbo].[@DOI]  T9 WHERE T9.[U_Group] =T0.[U_Group] and  T9.[U_Store] =T8.WhsCode))-(SELECT SUM(INQTY - OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode  AND DOCDATE <= '" & SOHdt & "' AND Warehouse=T8.WhsCode)-(SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode) " & _
                            '" End) + " & _
                            '" (SELECT SUM(OnOrder) FROM OITW WHERE ItemCode=T0.ItemCode and WhsCode=T8.WhsCode))/ (((SELECT SUM(OUTQTY) FROM OINM WHERE ItemCode=T0.ItemCode AND DOCDATE between '" & STfdt & "' and'" & STtdt & "' AND Warehouse=T8.WhsCode and transtype='13'))/(select  cast(datediff(dd, '" & STfdt & "','" & STtdt & "') as varchar) + 1 as 'M')) " & _
                            '" End as 'DOI', " & _


                            oRecordSet.DoQuery(str)
                            Dim i As Integer = 1
                            Dim k1 As Integer = oRecordSet.RecordCount / oRecordSet_WHSC.RecordCount
                            Dim co As Integer = oRecordSet.RecordCount
                            For i = 1 To k1
                                
                                Dim Icode As String = oRecordSet.Fields.Item(0).Value.ToString
                                SBO_Application.StatusBar.SetText("" & k1 & "  -of  " & i & "-Row Processing.Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                CardCode = oRecordSet.Fields.Item(0).Value.ToString
                                CardName = oRecordSet.Fields.Item(1).Value.ToString
                                U_Group = oRecordSet.Fields.Item(2).Value.ToString
                                U_Family = oRecordSet.Fields.Item(3).Value.ToString
                                U_Model = oRecordSet.Fields.Item(4).Value.ToString
                                U_Screensize = oRecordSet.Fields.Item(5).Value.ToString
                                U_HDCapacity = oRecordSet.Fields.Item(6).Value.ToString
                                U_Generation = oRecordSet.Fields.Item(7).Value.ToString
                                U_Network = oRecordSet.Fields.Item(8).Value.ToString
                                U_Kind = oRecordSet.Fields.Item(9).Value.ToString
                                U_ForModel = oRecordSet.Fields.Item(10).Value.ToString
                                U_Watt = oRecordSet.Fields.Item(11).Value.ToString
                                U_type = oRecordSet.Fields.Item(12).Value.ToString
                                U_Specification = oRecordSet.Fields.Item(13).Value.ToString
                                U_Fits = oRecordSet.Fields.Item(14).Value.ToString
                                U_Color = oRecordSet.Fields.Item(15).Value.ToString
                                U_Brand = oRecordSet.Fields.Item(16).Value.ToString
                                ' U_UPC = oRecordSet.Fields.Item(17).Value.ToString
                                U_AltCode1 = oRecordSet.Fields.Item(17).Value.ToString
                                U_AltCode2 = oRecordSet.Fields.Item(18).Value.ToString
                                U_AltCode3 = oRecordSet.Fields.Item(19).Value.ToString
                                SuppCatNum = oRecordSet.Fields.Item(20).Value.ToString
                                ItemName = oRecordSet.Fields.Item(21).Value.ToString
                                U_Status = oRecordSet.Fields.Item(22).Value.ToString
                                CreateDate = Format(oRecordSet.Fields.Item(23).Value, "dd/MM/yyyy")
                                U_Processor = oRecordSet.Fields.Item(24).Value.ToString
                                U_RAMSize = oRecordSet.Fields.Item(25).Value.ToString
                                MinStock = oRecordSet.Fields.Item(26).Value.ToString
                                MinOrdrQty = oRecordSet.Fields.Item(27).Value
                                Currency = oRecordSet.Fields.Item(28).Value.ToString
                                COST = oRecordSet.Fields.Item(29).Value.ToString
                                Retail_price_with_GST = oRecordSet.Fields.Item(30).Value.ToString
                                Retail_price_with_OUT_GST = oRecordSet.Fields.Item(31).Value.ToString
                                Margin_Amt = oRecordSet.Fields.Item(32).Value.ToString
                                Margin_Per = oRecordSet.Fields.Item(33).Value.ToString
                                U_MarginRank = oRecordSet.Fields.Item(34).Value.ToString
                                ItemCode = oRecordSet.Fields.Item(35).Value.ToString
                                U_Rebate = oRecordSet.Fields.Item(42).Value.ToString
                                ' Dim DOIUDF As Decimal = 0

                                'MsgBox(Math.Round(2.3566, 1))

                                DOIUDF = oRecordSet.Fields.Item(43).Value

                                Dim Day1 As Integer = 0
                                Day1 = oRecordSet.Fields.Item(44).Value.ToString
                                If oRecordSet_WHSC.RecordCount > 0 Then
                                    oRecordSet_WHSC.MoveFirst()
                                    For J = 1 To oRecordSet_WHSC.RecordCount
                                        soh(J - 1) = oRecordSet.Fields.Item(37).Value.ToString
                                        st(J - 1) = oRecordSet.Fields.Item(38).Value.ToString
                                        PO(J - 1) = oRecordSet.Fields.Item(39).Value.ToString
                                        SO(J - 1) = oRecordSet.Fields.Item(40).Value.ToString

                                        'Math.Round(New FileInfo(SO(J - 1)).Length / 1024, 1)
                                        If SO(J - 1) < 0 Then
                                            SO1(J - 1) = 0
                                            SO1(J - 1) = Math.Round(SO1(J - 1), 1)
                                        Else
                                            SO1(J - 1) = SO(J - 1)
                                            SO1(J - 1) = Math.Round(SO1(J - 1), 1)
                                        End If
                                        OD(J - 1) = oRecordSet.Fields.Item(40).Value.ToString
                                        OD1(J - 1) = OD(J - 1)
                                        'If OD(J - 1) > OD1(J - 1) Then
                                        '    OD1(J - 1) = OD1(J - 1) + 1
                                        'End If
                                        If OD1(J - 1) < 0 Then
                                            OD1(J - 1) = 0
                                        End If
                                        Try
                                            DOI(J - 1) = (soh(J - 1) + OD1(J - 1) + PO(J - 1)) / (st(J - 1) / Day1)
                                        Catch ex As Exception
                                            DOI(J - 1) = 0
                                        End Try
                                        DOI1(J - 1) = DOI(J - 1)
                                        'If DOI(J - 1) > DOI1(J - 1) Then
                                        '    DOI1(J - 1) = DOI1(J - 1) + 1
                                        'End If
                                        If DOI1(J - 1) < 0 Then
                                            DOI1(J - 1) = 0
                                        End If
                                        WhscCode(J - 1) = oRecordSet.Fields.Item(36).Value.ToString
                                        oRecordSet.MoveNext()
                                    Next
                                End If

                                Dim dr As DataRow
                                dr = table1.NewRow()
                                dr.Item(0) = CardCode
                                dr.Item(1) = CardName
                                dr.Item(2) = U_Group
                                dr.Item(3) = U_Family
                                dr.Item(4) = U_Model
                                dr.Item(5) = U_Screensize
                                dr.Item(6) = U_Processor
                                dr.Item(7) = U_RAMSize
                                dr.Item(8) = U_HDCapacity
                                dr.Item(9) = U_Generation
                                dr.Item(10) = U_Network
                                dr.Item(11) = U_Kind
                                dr.Item(12) = U_ForModel
                                dr.Item(13) = U_Watt
                                dr.Item(14) = U_type
                                dr.Item(15) = U_Specification
                                dr.Item(16) = U_Fits
                                dr.Item(17) = U_Color
                                dr.Item(18) = U_Brand
                                ' dr.Item(17) = U_UPC
                                dr.Item(19) = ItemCode
                                dr.Item(20) = U_AltCode1
                                dr.Item(21) = U_AltCode2
                                dr.Item(22) = U_AltCode3
                                dr.Item(23) = SuppCatNum
                                dr.Item(24) = ItemName
                                dr.Item(25) = U_Status
                                dr.Item(26) = CreateDate

                                dr.Item(27) = MinStock
                                dr.Item(28) = MinOrdrQty
                                dr.Item(29) = Currency
                                dr.Item(30) = COST
                                dr.Item(31) = Retail_price_with_GST
                                dr.Item(32) = Retail_price_with_OUT_GST
                                dr.Item(33) = Margin_Amt
                                dr.Item(34) = Margin_Per
                                dr.Item(35) = U_MarginRank
                                dr.Item(36) = U_Rebate
                                'dr.Item(35) = U_AltCode1
                                'dr.Item(36) = ItemCode
                                Dim v As Integer = 0
                                Dim m As Integer = 0
                                m = 1
                                Dim TOTSOH As Integer = 0
                                Dim TOTST As Integer = 0
                                Dim TOTPO As Integer = 0
                                Dim TOTSO As Decimal = 0
                                Dim TOTDOI As Decimal = 0
                                Dim TOTDOI1 As Integer = 0
                                For v = 1 To oRecordSet_WHSC.RecordCount
                                    dr.Item(36 + m) = WhscCode(v - 1)
                                    dr.Item(36 + m + 1) = soh(v - 1)
                                    TOTSOH = TOTSOH + soh(v - 1)
                                    dr.Item(36 + m + 2) = st(v - 1)
                                    TOTST = TOTST + st(v - 1)
                                    dr.Item(36 + m + 3) = PO(v - 1)
                                    TOTPO = TOTPO + PO(v - 1)
                                    dr.Item(36 + m + 4) = SO1(v - 1)
                                    dr.Item(36 + m + 5) = OD1(v - 1)
                                    dr.Item(36 + m + 6) = DOI1(v - 1)

                                    m = m + 7
                                Next
                                Dim whcount As Integer = oRecordSet_WHSC.RecordCount
                                Dim rno As Integer = 37 + (whcount * 7)
                                dr.Item(rno) = TOTSOH
                                dr.Item(rno + 1) = TOTST
                                dr.Item(rno + 2) = TOTPO
                                'TOTSO
                                Try
                                    TOTSO = 0.0
                                    TOTSO = ((TOTST / Day1) * DOIUDF) - TOTSOH - TOTPO
                                    
                                Catch ex As Exception
                                End Try
                                If TOTSO < 0 Then
                                    TOTSO = 0
                                End If
                                Dim TOTSO1 As Decimal = 0
                                TOTSO1 = TOTSO
                                TOTSO1 = Math.Round(TOTSO1, 1)

                                dr.Item(rno + 3) = TOTSO1

                                Try
                                    TOTDOI = 0
                                    TOTDOI = ((TOTSOH + TOTSO1 + TOTPO) / (TOTST / Day1))
                                Catch ex As Exception
                                    TOTDOI = 0
                                End Try
                                TOTDOI1 = TOTDOI
                                dr.Item(rno + 4) = TOTDOI1
                                table1.Rows.Add(dr)
                            Next
                            Try


                                SetDataTable_To_CSV(table1, "C:\ElushOrder" & Format(Now.Date, "yyyyMMdd") & ".txt", ",", WhscCount)

                                ' MsgBox("You can find the file C:\EXCEL\Test.xlsx")
                                SBO_Application.StatusBar.SetText("Report Generated", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                'CreateExcelFromCsvFile("C:\", "ElushOrder" & Format(Now.Date, "yyyyMMdd") & "")

                            Catch ex As Exception
                                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                ' SBO_Application.MessageBox(ex.Message)
                            End Try
                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                    End If
                    '------------End Print into CSV
                Catch ex As Exception
                    'SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End Try
            End If

        Catch ex As Exception
        End Try
    End Sub


    Sub SetDataTable_To_CSV(ByVal dtable As DataTable, ByVal path_filename As String, ByVal sep_char As String, ByVal whsccount As Integer)

        Dim writer As System.IO.StreamWriter

        Try

            writer = New System.IO.StreamWriter(path_filename)


            Dim _sep As String = ""

            Dim builder As New System.Text.StringBuilder

            'For Each col As DataColumn In dtable.Columns

            '    builder.Append(_sep).Append(col.ColumnName)

            '    _sep = sep_char
            '****************
            builder.Append(_sep).Append("Supplier Code")
            _sep = sep_char
            builder.Append(_sep).Append("Supplier Name")
            _sep = sep_char
            builder.Append(_sep).Append("GROUP")
            _sep = sep_char
            builder.Append(_sep).Append("Family")
            _sep = sep_char
            builder.Append(_sep).Append("Model")
            _sep = sep_char
            builder.Append(_sep).Append("Screen Size")
            _sep = sep_char
            builder.Append(_sep).Append("Processor")
            _sep = sep_char
            builder.Append(_sep).Append("RAM Size")
            _sep = sep_char
            builder.Append(_sep).Append("HD Capacity")
            _sep = sep_char
            builder.Append(_sep).Append("Generation")
            _sep = sep_char
            builder.Append(_sep).Append("Network")
            _sep = sep_char
            builder.Append(_sep).Append("Kind")
            _sep = sep_char
            builder.Append(_sep).Append("ForModel")
            _sep = sep_char
            builder.Append(_sep).Append("Watt")
            _sep = sep_char
            builder.Append(_sep).Append("Type")
            _sep = sep_char
            builder.Append(_sep).Append("Specification")
            _sep = sep_char
            builder.Append(_sep).Append("Fits")
            _sep = sep_char
            builder.Append(_sep).Append("Color")
            _sep = sep_char
            builder.Append(_sep).Append("Brand")
            _sep = sep_char
            'xlWorkSheet.Cells(7, 20) = "UPC"
            builder.Append(_sep).Append("Item Code")
            _sep = sep_char
            builder.Append(_sep).Append("AltCode1")
            _sep = sep_char
          
            builder.Append(_sep).Append("AltCode2")
            _sep = sep_char
            builder.Append(_sep).Append("AltCode3")
            _sep = sep_char
            builder.Append(_sep).Append("SuppCatNum")
            _sep = sep_char
            builder.Append(_sep).Append("Item Name")
            _sep = sep_char
            builder.Append(_sep).Append("Status")
            _sep = sep_char
            builder.Append(_sep).Append("Create Date")
            _sep = sep_char
           
            builder.Append(_sep).Append("Min Stock")
            _sep = sep_char
            builder.Append(_sep).Append("Min Order Qty")
            _sep = sep_char
            builder.Append(_sep).Append("Currency")
            _sep = sep_char
            builder.Append(_sep).Append("Cost")
            _sep = sep_char
            builder.Append(_sep).Append("Retail Price With GST")
            _sep = sep_char
            builder.Append(_sep).Append("Retail Price Without GST")
            _sep = sep_char
            builder.Append(_sep).Append("Margin Amount")
            _sep = sep_char
            builder.Append(_sep).Append("Margin %")
            _sep = sep_char
            builder.Append(_sep).Append("Margin Rank")
            _sep = sep_char
            builder.Append(_sep).Append("Rebate %")
            _sep = sep_char
            
            Dim kk As Integer = dtable.Columns.Count
            _sep = sep_char
            Dim kl As Integer = 38
            Dim mm As Integer = 0
            Dim i As Integer
            Dim kk1 As Integer = kk - 38
            For i = 1 To whsccount
                If mm = 0 Then
                    mm = 0 + kl + i
                Else
                    mm = 7 + mm
                End If
                builder.Append(_sep).Append("WHS Code")
                _sep = sep_char
                builder.Append(_sep).Append("SOH")
                _sep = sep_char
                builder.Append(_sep).Append("ST")
                _sep = sep_char
                builder.Append(_sep).Append("PO")
                _sep = sep_char
                builder.Append(_sep).Append("SO")
                _sep = sep_char
                builder.Append(_sep).Append("OD")
                _sep = sep_char
                builder.Append(_sep).Append("DOI")
                _sep = sep_char
            Next
            '********************
            builder.Append(_sep).Append("Total of SOH")
            _sep = sep_char
            builder.Append(_sep).Append("Total Of ST")
            _sep = sep_char
            builder.Append(_sep).Append("Total of PO")
            _sep = sep_char
            builder.Append(_sep).Append("Total of SO")
            _sep = sep_char
            builder.Append(_sep).Append("Total of DOI")
            _sep = sep_char
            ' Next

            writer.WriteLine(builder.ToString())



            For Each row As DataRow In dtable.Rows

                _sep = ""

                builder = New System.Text.StringBuilder



                For Each col As DataColumn In dtable.Columns

                    builder.Append(_sep).Append(row(col.ColumnName))

                    _sep = sep_char

                Next

                writer.WriteLine(builder.ToString())

            Next

        Catch ex As Exception

            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally

            If Not writer Is Nothing Then writer.Close()

        End Try
        Dim FileName As String = "C:\ElushOrder" & Format(Now.Date, "yyyyMMdd") & ".txt"

        If SBO_Application.MessageBox("Would you Like to Open The CSV File?", 1, "Yes", "No") = 1 Then
            Try

                Dim p As New System.Diagnostics.Process
                Dim s As New System.Diagnostics.ProcessStartInfo(FileName)
                s.UseShellExecute = True
                s.WindowStyle = ProcessWindowStyle.Normal
                p.StartInfo = s
                p.Start()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'SBO_Application.StatusBar.SetText("Path Name Is Empty!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End If

    End Sub

    Private Sub ExportToExcel(ByVal dtable As DataTable, ByVal whsccount As Integer)
        Try

            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim i As Integer
            Dim j As Integer
            xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")
            Dim rno As Integer
            rno = dtable.Rows.Count + 7
            xlWorkSheet.Range("C6", "AM" & (rno).ToString).Cells.NumberFormat = "@"
            Dim st As String = dtable.Rows(0)(1).ToString
            For i = 0 To dtable.Rows.Count - 1
                SBO_Application.StatusBar.SetText("" & dtable.Rows.Count & "  -of  " & i + 1 & "-Row Writing.Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                For j = 0 To dtable.Columns.Count - 1
                    xlWorkSheet.Cells(i + 8, j + 3) = dtable.Rows(i)(j).ToString()
                    xlWorkSheet.Cells(i + 8, 2) = i + 1
                Next
            Next
            xlWorkSheet.Cells(7, 2) = "No"
            xlWorkSheet.Cells(7, 3) = "Supplier Code"
            xlWorkSheet.Cells(7, 4) = "Supplier Name"
            xlWorkSheet.Cells(7, 5) = "GROUP"
            xlWorkSheet.Cells(7, 6) = "Family"
            xlWorkSheet.Cells(7, 7) = "Model"
            xlWorkSheet.Cells(7, 8) = "Screen Size"
            xlWorkSheet.Cells(7, 9) = "Processor"
            xlWorkSheet.Cells(7, 10) = "RAM Size"
            xlWorkSheet.Cells(7, 11) = "HD Capacity"
            xlWorkSheet.Cells(7, 12) = "Generation"
            xlWorkSheet.Cells(7, 13) = "Network"
            xlWorkSheet.Cells(7, 14) = "Kind"
            xlWorkSheet.Cells(7, 15) = "ForModel"
            xlWorkSheet.Cells(7, 16) = "Watt"
            xlWorkSheet.Cells(7, 17) = "Type"
            xlWorkSheet.Cells(7, 18) = "Specification"
            xlWorkSheet.Cells(7, 19) = "Fits"
            xlWorkSheet.Cells(7, 20) = "Color"
            xlWorkSheet.Cells(7, 21) = "Brand"
            'xlWorkSheet.Cells(7, 20) = "UPC"
            xlWorkSheet.Cells(7, 39) = "Item Code"
            xlWorkSheet.Cells(7, 38) = "AltCode1"
            xlWorkSheet.Cells(7, 22) = "AltCode2"
            xlWorkSheet.Cells(7, 23) = "AltCode3"
            xlWorkSheet.Cells(7, 24) = "SuppCatNum"
            xlWorkSheet.Cells(7, 25) = "Item Name"
            xlWorkSheet.Cells(7, 26) = "Status"
            xlWorkSheet.Cells(7, 27) = "Create Date"
           
            xlWorkSheet.Cells(7, 28) = "Min Stock"
            xlWorkSheet.Cells(7, 29) = "Min Order Qty"
            xlWorkSheet.Cells(7, 30) = "Currency"
            xlWorkSheet.Cells(7, 31) = "Cost"
            xlWorkSheet.Cells(7, 32) = "Retail Price With GST"
            xlWorkSheet.Cells(7, 33) = "Retail Price With Out GST"
            xlWorkSheet.Cells(7, 34) = "Margin Amount"
            xlWorkSheet.Cells(7, 35) = "Margin %"
            xlWorkSheet.Cells(7, 36) = "Margin Rank"
            xlWorkSheet.Cells(7, 37) = "Rebate %"
           
            Dim kk As Integer = dtable.Columns.Count
            Dim kl As Integer = 39
            Dim mm As Integer = 0
            Dim kk1 As Integer = kk - 39
            For i = 1 To whsccount
                If mm = 0 Then
                    mm = 0 + kl + i
                Else
                    mm = 7 + mm
                End If
                xlWorkSheet.Cells(7, mm) = "WHS Code"
                xlWorkSheet.Cells(7, mm + 1) = "SOH"
                xlWorkSheet.Cells(7, mm + 2) = "ST"
                xlWorkSheet.Cells(7, mm + 3) = "PO"
                xlWorkSheet.Cells(7, mm + 4) = "SO"
                xlWorkSheet.Cells(7, mm + 5) = "OD"
                xlWorkSheet.Cells(7, mm + 6) = "DOI"
            Next
            'dr.Item(rno) = TOTSOH
            'dr.Item(rno + 1) = TOTST
            'dr.Item(rno + 2) = TOTPO
            'dr.Item(rno + 3) = TOTSO
            mm = 40 + (whsccount * 7)
            xlWorkSheet.Cells(7, mm) = "Total of SOH"
            xlWorkSheet.Cells(7, mm + 1) = "Total of ST"
            xlWorkSheet.Cells(7, mm + 2) = "Total of PO"
            xlWorkSheet.Cells(7, mm + 3) = "Total of SO"

            Dim str As String = GetColumnName(dtable.Columns.Count + 2)
            xlWorkSheet = xlApp.Workbooks(1).ActiveSheet

            xlWorkSheet.Range("a1", "z1000").Font.Name = "Times New Roman"
            xlWorkSheet.Range("a1", "z1000").Font.Size = 10
            xlWorkSheet.Range("a7", "" & str & "7").Font.Size = 11
            xlWorkSheet.Range("a7", "" & str & "7").Font.Bold = True
            'xlWorkSheet.Range("b2", "b2").Font.Bold = True
            'xlWorkSheet.Range("b4", "b4").Font.Bold = True
            'xlWorkSheet.Range("f2", "f2").Font.Bold = True
            'xlWorkSheet.Range("f4", "f4").Font.Bold = True
            'xlWorkSheet.Range("j2", "j2").Font.Bold = True
            'xlWorkSheet.Range("j4", "j4").Font.Bold = True
            'xlWorkSheet.Range("a7", "r7").Font.Size = 11
            'xlWorkSheet.Range("b2", "b2").Font.Size = 11
            'xlWorkSheet.Range("b4", "b4").Font.Size = 11
            'xlWorkSheet.Range("f2", "f2").Font.Size = 11
            'xlWorkSheet.Range("f4", "f4").Font.Size = 11
            'xlWorkSheet.Range("j2", "j2").Font.Size = 11
            'xlWorkSheet.Range("j4", "j4").Font.Size = 11
            
            xlWorkSheet.Range("b7", "b" & (rno).ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            xlWorkSheet.Range("d7", "d" & (rno).ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            '  xlWorkSheet.Range("b7", "" & str & "7").Interior.Color = RGB(158, 151, 213)        ' xlWorkSheet.Range("b7", "q7").Interior.ColorIndex = 3
            xlWorkSheet.Range("b7", "" & str & "7").Interior.Color = RGB(169, 169, 169)        ' xlWorkSheet.Range("b7", "q7").Interior.ColorIndex = 3
            xlWorkSheet.Range("e7", "" & str & "" & (rno).ToString).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            xlWorkSheet.Range("b7", "" & str & "" & (rno).ToString).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Range("b7", "" & str & "" & (rno).ToString).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Range("b7", "" & str & "" & (rno).ToString).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Range("b7", "" & str & "" & (rno).ToString).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Range("b7", "" & str & "7").Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Range("b7", "" & str & "" & (rno).ToString).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous

            xlWorkSheet.Range("b8", "" & str & "8").Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Columns.AutoFit()
            xlWorkSheet.SaveAs("C:\ElushOrder" & Format(Now.Date, "yyyyMMdd") & ".xlsx")
            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
            SBO_Application.MessageBox("Report Generated. Path - C:\ElushOrder" & Format(Now.Date, "yyyyMMdd") & ".xlsx")
            If SBO_Application.MessageBox("Would you Like to Open The Excel File?", 1, "Yes", "No") = 1 Then
                Try
                    Dim FileName As String = "C:\ElushOrder" & Format(Now.Date, "yyyyMMdd") & ".xlsx"
                    Dim p As New System.Diagnostics.Process
                    Dim s As New System.Diagnostics.ProcessStartInfo(FileName)
                    s.UseShellExecute = True
                    s.WindowStyle = ProcessWindowStyle.Normal
                    p.StartInfo = s
                    p.Start()

                Catch ex As Exception
                    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'SBO_Application.StatusBar.SetText("Path Name Is Empty!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Function GetColumnName(ByVal colNum As Integer) As String
        Dim d As Integer
        Dim m As Integer
        Dim name As String
        d = colNum
        name = ""
        Do While (d > 0)
            m = (d - 1) Mod 26
            name = Chr(65 + m) + name
            d = Int((d - m) / 26)
        Loop
        GetColumnName = name
    End Function
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
  
End Class
