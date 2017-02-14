Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports System.Threading.Thread
Imports System.Threading

Module Module1
    Public oEdit As SAPbouiCOM.EditText
    Public oMatrix As SAPbouiCOM.Matrix
    Public oColumns As SAPbouiCOM.Columns
    Public oColumn As SAPbouiCOM.Column
    Public oCombo As SAPbouiCOM.ComboBox
    Public oGrid As SAPbouiCOM.Grid
    Public oCheck As SAPbouiCOM.CheckBox
    Public oRecordSet As SAPbobsCOM.Recordset
    Public oRecordSet_SOH As SAPbobsCOM.Recordset
    Public oRecordSet_ST As SAPbobsCOM.Recordset
    Public oRecordSet_PO As SAPbobsCOM.Recordset
    Public oRecordSet_SO As SAPbobsCOM.Recordset
    Public oRecordSet_DOI As SAPbobsCOM.Recordset
    Public oRecordSet_WHSC As SAPbobsCOM.Recordset
    Public oRecordSet_ONHD As SAPbobsCOM.Recordset
    Public oForm As SAPbouiCOM.Form
    Public Count As Integer = 0
    Public Bol As Boolean = False
#Region "CFL"
    Public Sub CFL_BP_Supplier(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application) 'Sales Tax
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFLBPV"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCFLCreationParams.UniqueID = "CFLBPV1"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub
#End Region
#Region "Combo Load"
    Public Sub ComboLoad_Group(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox, ByRef Ocompany As SAPbobsCOM.Company)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@GROUP]  T0")

            'oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@GROUP]  T0 where T0.Code IN ('3PP', 'APPLE')")
            Dim it As Integer = 1
            For it = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(it, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombo.ValidValues.Add("", "")
            Do While Not oRecordSet.EoF
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Loop
            Try
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception
            End Try
            oRecordSet = Nothing
            GC.Collect()
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub ComboLoad_Brand(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox, ByRef Ocompany As SAPbobsCOM.Company)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@BRAND]  T0")
            Dim it As Integer = 1
            For it = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(it, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombo.ValidValues.Add("", "")
            Do While Not oRecordSet.EoF
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Loop
            Try
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception
            End Try
            oRecordSet = Nothing
            GC.Collect()
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub ComboLoad_Status(ByRef Oform As SAPbouiCOM.Form, ByRef oCombo As SAPbouiCOM.ComboBox, ByRef Ocompany As SAPbobsCOM.Company)
        Try
            oRecordSet = Ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT T0.Code, T0.Name FROM [dbo].[@STATUS]  T0")
            Dim it As Integer = 1
            For it = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(it, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombo.ValidValues.Add("", "")
            Do While Not oRecordSet.EoF
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Loop

            Try
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception
            End Try
            oRecordSet = Nothing
            GC.Collect()
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Dim oXmlDoc As New Xml.XmlDocument
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        oXmlDoc.Load(sPath & "\Elush\" & FileName)
        Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub
    Public Sub Loadfile(ByVal FileName As String)
        Try
            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(FileName)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()
        Catch ex As Exception
        End Try
    End Sub
    Public Sub SaveAsXML(ByVal Form As SAPbouiCOM.Form, ByVal FileName As String)
        Dim oXmlDoc As New Xml.XmlDocument
        Dim sXmlString As String
        Dim sPath As String
        sXmlString = Form.GetAsXML
        oXmlDoc.LoadXml(sXmlString)
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        oXmlDoc.Save(sPath & "\HE\" & FileName)
    End Sub

End Module
