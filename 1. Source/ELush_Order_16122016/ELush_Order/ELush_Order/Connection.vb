Option Strict Off
Option Explicit On
Imports System.Diagnostics.Process
Imports System.Threading
Imports System.Net.Mail
Public Class Connection
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public ocompany As New SAPbobsCOM.Company
    Private oOrderForm As SAPbouiCOM.Form
    Public SboGuiApi As New SAPbouiCOM.SboGuiApi
    Public sConnectionString As String
    Dim oF_OrderRecReport As F_OrderRecReport
#Region "COnnce"

    '=======================
    Private Sub SetApplication()
        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        SboGuiApi = New SAPbouiCOM.SboGuiApi
        '// by following the steps specified above, the following
        '// Statment should be suficient for either development or run mode
        'sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        sConnectionString = Environment.GetCommandLineArgs.GetValue(1) '
        'sConnectionString = "5645523035496D706C656D656E746174696F6E3A59313931303035313531383699469FA92C3C9A964A219C5862952A90D911E9" 'Environment.GetCommandLineArgs.GetValue(1)'
        Try
            SboGuiApi.Connect(sConnectionString)
            '// connect to a running SBO Application
            '// get an initialized application object
            SBO_Application = SboGuiApi.GetApplication()
        Catch ex As Exception
            MsgBox("Make Sure That SAP Business One Application is running!!! ", MsgBoxStyle.Information)
            End
        End Try
        'SBO_Application.StatusBar.SetText("DI is Connecting now", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    ''=======================
    'Private Sub SetApplication()

    '    '*******************************************************************
    '    '// Use an SboGuiApi object to establish connection
    '    '// with the SAP Business One application and return an
    '    '// initialized appliction object
    '    '*******************************************************************

    '    Dim SboGuiApi As SAPbouiCOM.SboGuiApi
    '    Dim sConnectionString As String

    '    SboGuiApi = New SAPbouiCOM.SboGuiApi

    '    '// by following the steps specified above, the following
    '    '// statment should be suficient for either development or run mode

    '    sConnectionString = Environment.GetCommandLineArgs.GetValue(1)

    '    '// connect to a running SBO Application

    '    SboGuiApi.Connect(sConnectionString)

    '    '// get an initialized application object

    '    SBO_Application = SboGuiApi.GetApplication()

    'End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String
        Dim sConnectionContext As String
        Dim lRetCode As Integer = 0

        '// First initialize the Company object

        oCompany = New SAPbobsCOM.Company

        '// Acquire the connection context cookie from the DI API.
        sCookie = oCompany.GetContextCookie

        '// Retrieve the connection context string from the UI API using the
        '// acquired cookie.
        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

        '// before setting the SBO Login Context make sure the company is not
        '// connected

        If oCompany.Connected = True Then
            oCompany.Disconnect()
        End If

        '// Set the connection context information to the DI API.
        SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

    End Function

    Private Function ConnectToCompany() As Integer

        '// Establish the connection to the company database.
        ConnectToCompany = oCompany.Connect

    End Function
    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
    Private Sub Class_Initialize_Renamed()

        '//*************************************************************
        '// set SBO_Application with an initialized application object
        '//*************************************************************

        SetApplication()

        '//*************************************************************
        '// Set The Connection Context
        '//*************************************************************

        If Not SetConnectionContext() = 0 Then
            SBO_Application.MessageBox("Failed setting a connection to DI API")
            End ' Terminating the Add-On Application
        End If


        '//*************************************************************
        '// Connect To The Company Data Base
        '//*************************************************************

        If Not ConnectToCompany() = 0 Then
            SBO_Application.MessageBox("Failed connecting to the company's Data Base")
            End ' Terminating the Add-On Application
        End If

        '//*************************************************************
        '// send an "hello world" message
        '//*************************************************************

        ' SBO_Application.MessageBox("DI Connected To: " & ocompany.CompanyName & vbNewLine & "Hello World!..Gopi")

    End Sub

#End Region
#Region "Connection"
    Public Sub New()
        MyBase.new()
        Try
            Class_Initialize_Renamed()
            ' LoadFromXML("MyMenus.xml", SBO_Application)
            AddMenuItems()
            oF_OrderRecReport = New F_OrderRecReport(ocompany, SBO_Application)
            'SBO_Application.MessageBox("Welcome To Elush...")
            SBO_Application.StatusBar.SetText("Welcome To Elush...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub conn1()
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String
        Try
            sconn = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sconn)
            SBO_Application = SboGuiApi.GetApplication
            SboGuiApi = Nothing
            scook = ocompany.GetContextCookie
            str = SBO_Application.Company.GetConnectionContext(scook)
            ret = ocompany.SetSboLoginContext(str)
            ocompany.Connect()
            ocompany.GetLastError(ret, str)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

    Private Sub AddMenuItems()
        '//******************************************************************
        '// Let's add a separator, a pop-up menu item and a string menu item
        '//******************************************************************

        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem

        Dim i As Integer '// to be used as counter
        Dim lAddAfter As Integer
        Dim sXML As String

        '// Get the menus collection from the application
        oMenus = SBO_Application.Menus
        '--------------------------------------------
        'Save an XML file containing the menus...
        '--------------------------------------------
        'sXML = SBO_Application.Menus.GetAsXML
        'Dim xmlD As System.Xml.XmlDocument
        'xmlD = New System.Xml.XmlDocument
        'xmlD.LoadXml(sXML)
        'xmlD.Save("c:\\mnu.xml")
        '--------------------------------------------


        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oMenuItem = SBO_Application.Menus.Item("2304") 'moudles'

        Dim sPath As String

        sPath = Application.StartupPath
        sPath = sPath.Remove(sPath.Length - 3, 3)

        '// find the place in wich you want to add your menu item
        '// in this example I chose to add my menu item under
        '// SAP Business One.
        ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        ' ''oCreationPackage.UniqueID = "MyMenu0111"
        ' ''oCreationPackage.String = "Sample Menu"
        ' ''oCreationPackage.Enabled = True
        '' '' oCreationPackage.Image = sPath & "UI.bmp"
        ' ''oCreationPackage.Position = 0

        ' ''oMenus = oMenuItem.SubMenus

        Try ' If the manu already exists this code will fail
            ''oMenus.AddEx(oCreationPackage)

            '// Get the menu collection of the newly added pop-up item
            oMenuItem = SBO_Application.Menus.Item("2304")
            oMenus = oMenuItem.SubMenus

            '// Create s sub menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "MySubMenu01"
            oCreationPackage.String = "Order Replenishment Report"
            oCreationPackage.Position = 1
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
            '  SBO_Application.MessageBox("Menu Already Exists")
        End Try

    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
       
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = True Then
                If pVal.MenuUID = "MySubMenu01" Then
                    LoadFromXML("OrderRec.srf", SBO_Application)
                    oForm = SBO_Application.Forms.Item("Order_Replenishment_Report")
                    oF_OrderRecReport.Order_bind(oForm)
                End If

            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
End Class
