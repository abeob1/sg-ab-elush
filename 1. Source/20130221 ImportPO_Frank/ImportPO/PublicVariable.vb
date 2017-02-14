Public Class PublicVariable
    'Connection
    Public Shared oCompany As SAPbobsCOM.Company
    Public Shared Token As String = ""
    '----------SEND EMAIL INFORMATION--------------------
    Public Shared ToEmail As String = "thuytruong@electra-ai.com"
    Public Shared ToEmailName As String = "FRANK"
    Public Shared smtpServer As String = "smtp.gmail.com"
    Public Shared smtpPort As String = "587"
    Public Shared smtpSenderEmail As String = "truongthaithuy@gmail.com"
    Public Shared smtpPwd As String = "KHONGbiet"
    Public Shared EmailSub As String = "Order Replenishment"

    Public Shared POlayoutCode As String = "POR2002"
End Class
