Public Class Functions
    Public Shared Sub WriteLog(ByVal Str As String)
        Dim oWrite As IO.StreamWriter
        Dim FilePath As String
        FilePath = Application.StartupPath + "\logfile.txt"

        If IO.File.Exists(FilePath) Then
            oWrite = IO.File.AppendText(FilePath)
        Else
            oWrite = IO.File.CreateText(FilePath)
        End If
        oWrite.Write(Now.ToString() + ":" + Str + vbCrLf)
        oWrite.Close()
    End Sub
    Public Shared Function CreateToken() As String
        Try
            Dim str As String
            str = DateTime.Now.ToString("yyyyMMddhhmmssfff")
            Return str
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Function DoQueryReturnDT(query As String) As DataTable
        Dim dt As DataTable
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRecordSet.DoQuery(query)
            If oRecordSet.RecordCount > 0 Then
                dt = ConvertRS2DT(oRecordSet)
                Return dt
            Else
                Return Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function
    Private Function ConvertRS2DT(ByVal RS As SAPbobsCOM.Recordset) As DataTable
        Dim dtTable As New DataTable
        Dim NewCol As DataColumn
        Dim NewRow As DataRow
        Dim ColCount As Integer
        Try
            For ColCount = 0 To RS.Fields.Count - 1
                Dim dataType As String = "System."
                Select Case RS.Fields.Item(ColCount).Type
                    Case SAPbobsCOM.BoFieldTypes.db_Alpha
                        dataType = dataType & "String"
                    Case SAPbobsCOM.BoFieldTypes.db_Date
                        dataType = dataType & "DateTime"
                    Case SAPbobsCOM.BoFieldTypes.db_Float
                        dataType = dataType & "Double"
                    Case SAPbobsCOM.BoFieldTypes.db_Memo
                        dataType = dataType & "String"
                    Case SAPbobsCOM.BoFieldTypes.db_Numeric
                        dataType = dataType & "Decimal"
                    Case Else
                        dataType = dataType & "String"
                End Select

                NewCol = New DataColumn(RS.Fields.Item(ColCount).Name, System.Type.GetType(dataType))
                dtTable.Columns.Add(NewCol)
            Next

            Do Until RS.EoF

                NewRow = dtTable.NewRow
                'populate each column in the row we're creating
                For ColCount = 0 To RS.Fields.Count - 1

                    NewRow.Item(RS.Fields.Item(ColCount).Name) = RS.Fields.Item(ColCount).Value

                Next

                'Add the row to the datatable
                dtTable.Rows.Add(NewRow)

                RS.MoveNext()
            Loop
            Return dtTable
        Catch ex As Exception
            MsgBox(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", MsgBoxStyle.Exclamation)
            Return Nothing
        End Try
    End Function

    Public Sub LoadParameters()
        Dim dt As DataTable = DoQueryReturnDT("select * from SAP_Integration..Mapping")
        For Each dr As DataRow In dt.Rows
            Select Case dr("Code").ToString()
                Case "ToEmail"
                    PublicVariable.ToEmail = dr("Value").ToString
                Case "ToEmailName"
                    PublicVariable.ToEmailName = dr("Value").ToString
                Case "smtpServer"
                    PublicVariable.smtpServer = dr("Value").ToString
                Case "smtpPort"
                    PublicVariable.smtpPort = dr("Value").ToString
                Case "smtpSenderEmail"
                    PublicVariable.smtpSenderEmail = dr("Value").ToString
                Case "smtpPwd"
                    PublicVariable.smtpPwd = dr("Value").ToString
                    'Case "EmailSub"
                    '    PublicVariable.EmailSub = "" 'dr("Value").ToString
                Case "POlayoutCode"
                    PublicVariable.POlayoutCode = dr("Value").ToString
            End Select
        Next
    End Sub
End Class
