'Imports System.Net.Mail
Imports System.Web.Mail
Imports System.Net
Imports System.Globalization
Imports System.IO

Public Class oEmailPO2VendorWeb

    Public Function SendPOEmail(dt As DataTable) As String
        Try
            Dim ret As String = ""
            Dim Subject As String = ""
            Dim ToEmail As String = ""
            Dim CCEmail As String = ""
            Dim fn As New Functions
            Dim opr As New oPrint
            Dim SQLwd As String = ""
            Dim ors As SAPbobsCOM.Recordset
            Dim ParaValue As String = ""
            Dim ParaName As String = "DocKey@"
            Dim PDFFile As String = Application.StartupPath + "\PO"

            ors = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors.DoQuery("Select * from [@SAPWD]")
            If ors.RecordCount > 0 Then
                SQLwd = ors.Fields.Item("Name").Value.ToString()
            Else
                Return "Please setup SQL password!"
            End If

            Dim Connection As String = PublicVariable.oCompany.Server + ";" + PublicVariable.oCompany.CompanyDB + ";" + PublicVariable.oCompany.DbUserName + ";" + SQLwd
            For Each dr As DataRow In dt.Rows
                Dim dtHeader As DataTable = dt.Clone
                dtHeader.ImportRow(dr)
                ParaValue = dr("DocEntry").ToString
                Subject = "Elush Purchase Order - PO#" + dr("DocNum").ToString + "-" + dr("CardCode").ToString + "-" + dr("WhsCode").ToString
                ToEmail = dr("E_Mail").ToString
                CCEmail = dr("ToEmailList").ToString
                PDFFile = PDFFile + dr("DocEntry").ToString + ".pdf"
                ret = opr.PrintCrystalReport(Application.StartupPath + "\PurchaseOrder.rpt", "", Connection, ParaName, ParaValue, True, PDFFile)
                If ret <> "" Then
                    Return ret
                End If
                ret = SendMailAttachFile(ToEmail, Subject, "", CCEmail, PDFFile)
                If ret <> "" Then
                    Return ret
                End If
                File.Delete(PDFFile)
            Next
            Return ""
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
            Return ex.ToString
        End Try
    End Function
    Public Function SendMailAttachFile(ToEmailList As String, Subject As String, TemplatePath As String, CCEmail As String, FilePath As String) As String
        Dim l_SenderEmail As String = PublicVariable.smtpSenderEmail
        If ToEmailList.Trim().Equals(String.Empty) Then
            ToEmailList = l_SenderEmail
        End If

        Dim mail As New MailMessage()
        mail.To = ToEmailList
        If CCEmail <> "" Then
            mail.Cc = CCEmail
        End If
        mail.From = l_SenderEmail
        mail.Subject = Subject
        mail.Body = "Please find the attached Purchase Order." + vbCrLf + "Thank you"
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'basic authentication
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", l_SenderEmail) 'set your username here
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", PublicVariable.smtpPwd) 'set your password here

        If File.Exists(FilePath) Then
            mail.Attachments.Add(New MailAttachment(FilePath))
        End If
        Try
            SmtpMail.SmtpServer = PublicVariable.smtpServer 'your real server goes here
            SmtpMail.Send(mail)

        Catch ex As System.Web.HttpException
            Return ex.Message
        End Try
        
        Return ""
    End Function

End Class
