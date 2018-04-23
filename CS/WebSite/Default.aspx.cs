using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using DevExpress.XtraPrinting;
using System.IO;
using DevExpress.XtraReports.UI;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        using (MemoryStream stream = new MemoryStream())
        {
            XtraReport rep = PivotReportGenerator.GenerateReport(ASPxPivotGrid1, GetReportGeneratorType( RadioButtonList1.Value.ToString()), Convert.ToInt32( ASPxSpinEdit1.Number), ASPxCheckBox1.Checked);
            rep.ExportToPdf(stream);
            ExportToResponse(stream, "Book", "pdf", "application/pdf", true);
        }
    }

    private ReportGeneratorType GetReportGeneratorType(string value)
    {
        return (ReportGeneratorType)Enum.Parse(typeof(ReportGeneratorType), value);
    }

    protected void ExportToResponse(MemoryStream stream, string fileName, string fileFormat, string contentType, bool saveAsFile)
    {
        if (Page == null || Page.Response == null) return;
        if (String.IsNullOrEmpty(fileName)) fileName = "ASPxPivotGrid";
        string disposition = saveAsFile ? "attachment" : "inline";
        Page.Response.Clear();
        Page.Response.Buffer = false;
        Page.Response.AppendHeader("Content-Type", contentType);
        Page.Response.AppendHeader("Content-Transfer-Encoding", "binary");
        Page.Response.AppendHeader("Content-Disposition", string.Format("{0}; filename={1}.{2}", disposition,
            HttpUtility.UrlEncode(fileName).Replace("+", "%20"), fileFormat));
        Page.Response.BinaryWrite(stream.ToArray());
        Page.Response.End();
    }
}
