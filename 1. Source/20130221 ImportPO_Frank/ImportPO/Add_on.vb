Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Data
Imports System.Threading

Public Class Add_on
    Private WithEvents SBO_Application As SAPbouiCOM.Application

#Region "Initial"
    Public Sub New()
        MyBase.New()
        Class_Init()
        AddMenuItems()    
    End Sub
    Public Sub SetApplication()
        Dim sbogui As SAPbouiCOM.SboGuiApi
        Dim oconnection As String
        sbogui = New SAPbouiCOM.SboGuiApi
        If Environment.GetCommandLineArgs().Length = 1 Then
            oconnection = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        Else
            oconnection = Environment.GetCommandLineArgs.GetValue(1)
        End If

        Try
            sbogui.Connect(oconnection)
        Catch ex As Exception
            MsgBox("No SAP Application Running")
            End
        End Try
        SBO_Application = sbogui.GetApplication
    End Sub
    Private Function SetConnectionContext() As Integer
        Dim sCookie As String
        Dim sConnectionContext As String
        Try
            PublicVariable.oCompany = New SAPbobsCOM.Company
            sCookie = PublicVariable.oCompany.GetContextCookie
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
            If PublicVariable.oCompany.Connected = True Then
                PublicVariable.oCompany.Disconnect()
            End If
            SetConnectionContext = PublicVariable.oCompany.SetSboLoginContext(sConnectionContext)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Private Function ConnectToCompany() As Integer
        ConnectToCompany = PublicVariable.oCompany.Connect

    End Function
    Private Sub Class_Init()
        SetApplication()
        If Not SetConnectionContext() = 0 Then
            SBO_Application.MessageBox("Failed setting a connection to DI API")
            End ' Terminating the Add-On Application
        End If
        If Not ConnectToCompany() = 0 Then
            SBO_Application.MessageBox("Failed connecting to the company's Database")
            End ' Terminating the Add-On Application
        End If
        Dim fn As New Functions
        fn.LoadParameters()

        Dim ocr As New oCRLayout
        ocr.GetCRLayout("PurchaseOrder.rpt", PublicVariable.POlayoutCode)

        SBO_Application.SetStatusBarMessage("Add-on is loaded", , False)
    End Sub
    Private Sub AddMenuItems()
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = SBO_Application.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oMenuItem = SBO_Application.Menus.Item("2304") 'purchase module'

        Dim name As String = "SM_SF"
        If oMenus.Exists(name) Then
            oMenus.RemoveEx(name)
        End If
        Try ' If the manu already exists this code will fail
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "SM_SF"
            oCreationPackage.String = "Import PO"
            oCreationPackage.Position = 0
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Error Handling
            SBO_Application.MessageBox(er.Message)
        End Try

        name = "SM_SF1"
        If oMenus.Exists(name) Then
            oMenus.RemoveEx(name)
        End If
        Try ' If the manu already exists this code will fail
            oMenus = oMenuItem.SubMenus
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "SM_SF1"
            oCreationPackage.String = "Email PO to Supplier"
            oCreationPackage.Position = 0
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Error Handling
            SBO_Application.MessageBox(er.Message)
        End Try
    End Sub
#End Region
#Region "SAP Event"
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        If pVal.BeforeAction = False Then
            Select Case pVal.MenuUID
                Case "SM_SF"
                    Dim othr As ThreadStart, myThread As Thread
                    othr = New ThreadStart(AddressOf CallForm_frmImport)
                    myThread = New Thread(othr)
                    myThread.SetApartmentState(ApartmentState.STA)
                    myThread.Start()
                Case "SM_SF1"
                    Dim othr As ThreadStart, myThread As Thread
                    othr = New ThreadStart(AddressOf CallForm_frmEmailToVendor)
                    myThread = New Thread(othr)
                    myThread.SetApartmentState(ApartmentState.STA)
                    myThread.Start()
            End Select
        End If
    End Sub
    Private Sub SBO_Application__AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                System.Windows.Forms.Application.Exit()
        End Select
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
       
    End Sub
#End Region
#Region "Functions"
    Private Sub CallForm_frmImport()
        Dim frm As New frmImportPO
        frm.Show()
        frm.Activate()
        System.Windows.Forms.Application.Run()
    End Sub
    Private Sub CallForm_frmEmailToVendor()
        Dim frm As New frmEmailToVendor
        frm.Show()
        frm.Activate()
        System.Windows.Forms.Application.Run()
    End Sub
#End Region
End Class
