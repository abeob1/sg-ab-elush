Imports System.Drawing.Printing
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmReport
    Dim rptDocument As ReportDocument
    Public Sub showForm(ByRef myReport As CrystalDecisions.CrystalReports.Engine.ReportDocument)
        CrystalReportViewer1.ReportSource = myReport
        rptDocument = myReport
        CrystalReportViewer1.Visible = True
        CrystalReportViewer1.Show()
        Me.ShowDialog()
    End Sub
End Class