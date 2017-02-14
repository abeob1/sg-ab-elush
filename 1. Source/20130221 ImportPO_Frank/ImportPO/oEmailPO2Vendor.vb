Imports System.Net.Mail
Imports System.Net
Imports System.Globalization
Imports System.IO

Public Class oEmailPO2Vendor

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
        'Using msg As New MailMessage

        'End Using

        Dim msg As New System.Net.Mail.MailMessage()
        msg.From = New MailAddress(l_SenderEmail, "ELUSH")
        msg.[To].Add(ToEmailList)
        If CCEmail <> "" Then
            msg.CC.Add(CCEmail)
        End If
        msg.Subject = Subject
        msg.Body = "Please find the attached Purchase Order." + vbCrLf + "Thank you"
        msg.IsBodyHtml = True
        If File.Exists(FilePath) Then

            msg.Attachments.Add(New Attachment(FilePath))
        End If
        Try
            Dim client As New SmtpClient(PublicVariable.smtpServer, PublicVariable.smtpPort)
            client.EnableSsl = True
            client.Timeout = 0
            client.UseDefaultCredentials = False
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.Credentials = New NetworkCredential(l_SenderEmail, PublicVariable.smtpPwd)
            client.Send(msg)
            msg.Attachments.Dispose()
        Catch ex As SmtpException
            Return ex.Message
        End Try
        Return ""
    End Function
    
End Class
