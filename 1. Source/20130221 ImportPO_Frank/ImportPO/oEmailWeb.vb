Imports System.Web.Mail
Imports System.Net
Imports System.Globalization
Imports System.IO

Public Class oEmailWeb

    Public Function SendMailByDS(ds As DataSet, ToEmailList As String, Subject As String, TemplatePath As String) As String


        Dim l_SenderEmail As String = PublicVariable.smtpSenderEmail
        If ToEmailList.Trim().Equals(String.Empty) Then
            ToEmailList = l_SenderEmail
        End If

        Dim mail As New MailMessage()
        mail.To = ToEmailList
        mail.From = l_SenderEmail
        mail.Subject = Subject
        mail.Body = GetTemplateforDS(ds, TemplatePath)
        mail.BodyFormat = MailFormat.Html
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'basic authentication
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", l_SenderEmail) 'set your username here
        mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", PublicVariable.smtpPwd) 'set your password here

        Try
            SmtpMail.SmtpServer = PublicVariable.smtpServer 'your real server goes here
            SmtpMail.Send(mail)

        Catch ex As System.Web.HttpException
            Return ex.Message
        End Try

        Return ""
    End Function
    Private Function GetTemplateforDS(Ds As DataSet, TemplatePath As String) As String
        Dim l_Rs As String = ""
        Try
            Dim l_PathTemplate As String = String.Empty

            If TemplatePath.Trim().Equals(String.Empty) Then
                l_Rs = String.Format("Template is empty")
            Else

                l_Rs = File.ReadAllText(TemplatePath)

                '-----------header-----------------
                Dim dr As DataRow = Ds.Tables(0).Rows(0)
                Dim ivC As CultureInfo = New System.Globalization.CultureInfo("es-US")
                l_Rs = l_Rs.Replace("<@Date>", [String].Format("{0:dd/MM/yyyy}", Convert.ToDateTime(dr("DocDate"), ivC)))

                '-----------line-----------------
                Dim str As String = ""
                Dim Total1 As Double = 0
                Dim Total2 As Double = 0
                For i As Integer = 0 To Ds.Tables(0).Rows.Count - 1
                    str = str & "<tr>"
                    str = str & "<td align=""left"" style=""border: thin solid #008080;""><@Store" & i.ToString() & "></td>"
                    str = str & "<td align=""right"" style=""border: thin solid #008080;""><@Amt1" & i.ToString() & "></td>"
                    str = str & "<td align=""right"" style=""border: thin solid #008080;""><@Amt2" & i.ToString() & "></td>"
                    str = str & "</tr>"
                Next
                l_Rs = l_Rs.Replace("<@ITEMLINEHERE>", str)
                Dim j As Integer = 0
                For Each dr1 As DataRow In Ds.Tables(0).Rows
                    l_Rs = l_Rs.Replace("<@Store" & j.ToString() & ">", dr1("WhsCode"))
                    l_Rs = l_Rs.Replace("<@Amt1" & j.ToString() & ">", String.Format("{0:n2}", dr1("LineTotal")))
                    l_Rs = l_Rs.Replace("<@Amt2" & j.ToString() & ">", String.Format("{0:n2}", dr1("GTotal")))
                    j += 1

                    Total1 = Total1 + dr1("LineTotal")
                    Total2 = Total2 + dr1("GTotal")
                Next
                '-----------footer-----------------
                l_Rs = l_Rs.Replace("<@Amount1>", String.Format("{0:n2}", Total1))
                l_Rs = l_Rs.Replace("<@Amount2>", String.Format("{0:n2}", Total2))
            End If
        Catch ex As Exception
            Functions.WriteLog(ex.ToString())
            Return ""
        End Try
        Return l_Rs
    End Function
End Class
