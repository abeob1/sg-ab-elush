Imports System.IO

Public Class oCRLayout
    Public Sub GetCRLayout(FileName As String, LayoutCode As String)
        Try
            Dim blobNewFilePath As String = Application.StartupPath + "\" + FileName

            'If File.Exists(blobNewFilePath) Then
            '    Return
            'End If

            Dim oCompanyService As SAPbobsCOM.CompanyService = PublicVariable.oCompany.GetCompanyService()

            ' Specify a table and blob field 
            Dim oBlobParams As SAPbobsCOM.BlobParams
            oBlobParams = DirectCast(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), SAPbobsCOM.BlobParams)
            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            ' Specify key name and key value of a record to update 
            Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"
            oKeySegment.Value = LayoutCode '"POR20002"

            Dim oBlob As SAPbobsCOM.Blob
            oBlob = DirectCast(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob), SAPbobsCOM.Blob)

            ' Get contents of a blob field 
            oBlob = oCompanyService.GetBlob(oBlobParams)

            ' Convert Base64 string to binary 
            Dim buf As Byte()
            buf = Convert.FromBase64String(oBlob.Content)

            ' Write blob file to file system 
            Dim oFile As New FileStream(blobNewFilePath, FileMode.Create, FileAccess.Write)
            oFile.Write(buf, 0, buf.Length)
            oFile.Close()
            oFile.Dispose()
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
        End Try
    End Sub
End Class
