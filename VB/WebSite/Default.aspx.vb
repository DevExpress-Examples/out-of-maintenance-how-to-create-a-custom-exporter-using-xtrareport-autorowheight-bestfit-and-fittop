Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports DevExpress.XtraPrinting
Imports System.IO
Imports DevExpress.XtraReports.UI

Partial Public Class _Default
	Inherits System.Web.UI.Page
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	End Sub

	Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs)
		Using stream As New MemoryStream()
			Dim rep As XtraReport = PivotReportGenerator.GenerateReport(ASPxPivotGrid1, GetReportGeneratorType(RadioButtonList1.Value.ToString()), Convert.ToInt32(ASPxSpinEdit1.Number), ASPxCheckBox1.Checked)
			rep.ExportToPdf(stream)
			ExportToResponse(stream, "Book", "pdf", "application/pdf", True)
		End Using
	End Sub

	Private Function GetReportGeneratorType(ByVal value As String) As ReportGeneratorType
		Return CType(System.Enum.Parse(GetType(ReportGeneratorType), value), ReportGeneratorType)
	End Function

	Protected Sub ExportToResponse(ByVal stream As MemoryStream, ByVal fileName As String, ByVal fileFormat As String, ByVal contentType As String, ByVal saveAsFile As Boolean)
		If Page Is Nothing OrElse Page.Response Is Nothing Then
			Return
		End If
		If String.IsNullOrEmpty(fileName) Then
			fileName = "ASPxPivotGrid"
		End If
		Dim disposition As String
		If saveAsFile Then
			disposition = "attachment"
		Else
			disposition = "inline"
		End If
		Page.Response.Clear()
		Page.Response.Buffer = False
		Page.Response.AppendHeader("Content-Type", contentType)
		Page.Response.AppendHeader("Content-Transfer-Encoding", "binary")
		Page.Response.AppendHeader("Content-Disposition", String.Format("{0}; filename={1}.{2}", disposition, HttpUtility.UrlEncode(fileName).Replace("+", "%20"), fileFormat))
		Page.Response.BinaryWrite(stream.ToArray())
		Page.Response.End()
	End Sub
End Class
