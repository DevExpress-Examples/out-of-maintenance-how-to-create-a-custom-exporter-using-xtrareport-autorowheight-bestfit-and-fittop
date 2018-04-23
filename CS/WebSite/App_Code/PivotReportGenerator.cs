using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraReports.UI;
using DevExpress.XtraPivotGrid;
using DevExpress.XtraEditors;
using DevExpress.Web.ASPxPivotGrid;

public static class PivotReportGenerator
{
    public static XtraReport GenerateReport(ASPxPivotGrid pivot,ReportGeneratorType kind,int columnWidth,bool repeatRowHeader)
    {
            XtraReport rep = new XtraReport();
            rep.DataSource = FillDataset(pivot);
            rep.DataMember = ((DataSet)rep.DataSource).Tables[0].TableName;
            InitBands(rep);
            InitStyles(rep);
            InitDetailsBasedonXRTable(rep, kind, columnWidth, repeatRowHeader);
            return rep;
    }
    public static DataSet FillDataset(ASPxPivotGrid pivot)
    {
        DataSet dataSet1 = new DataSet();
        dataSet1.DataSetName = "PivotGridColumns";
        DataTable dataTable1 = new DataTable();
        dataSet1.Tables.Add(dataTable1);
        FillDatasetColumns(pivot, dataTable1);
        FillDatasetExtracted(pivot, dataTable1);

        return dataSet1;
    }
    #region PreparingDataSet
    private static void FillDatasetExtracted(ASPxPivotGrid pivot, DataTable dataTable1)
    {       

        List<object> rowvalues = new List<object>();
        string tempRowText = "";
        List<PivotGridFieldBase> fieldsInRowArea = GetFieldsInArea(pivot, PivotArea.RowArea);
        for (int i = 0; i < pivot.RowCount; i++)
        {
            DevExpress.XtraPivotGrid.Data.PivotGridCellItem pcea = GetCellItem(pivot, 0, i); 
            if ( pcea.RowValueType == PivotGridValueType.Value) 
            {
                foreach (PivotGridFieldBase item in fieldsInRowArea)
                    tempRowText += pcea.GetFieldValue(GetFieldItemByField(pivot, item)).ToString() + " | "; //add formatting if it's necessary
                tempRowText = tempRowText.Remove(tempRowText.Length - 3, 3);
            }
            else
                tempRowText = pcea.RowValueType.ToString();
            rowvalues.Clear();
            rowvalues.Add(tempRowText);
            tempRowText = "";
            for (int j = 0; j < pivot.ColumnCount; j++)
            {
                pcea = GetCellItem(pivot, j, i);
                if (pcea.Value != null)
                    rowvalues.Add(pcea.Value);
                else
                    rowvalues.Add(DBNull.Value);
            }
            dataTable1.Rows.Add(rowvalues.ToArray());
        }
    }

    private static DevExpress.XtraPivotGrid.Data.PivotFieldItemBase GetFieldItemByField(ASPxPivotGrid pivot, PivotGridFieldBase field)
    {
        return pivot.Data.GetFieldItem(field);
    }



    private static DevExpress.XtraPivotGrid.Data.PivotGridCellItem GetCellItem(ASPxPivotGrid pivot, int columnIndex, int rowIndex)
    {
        DevExpress.XtraPivotGrid.Data.PivotFieldValueItem columnItem = pivot.Data.VisualItems.GetLastLevelItem(true, columnIndex, false);
        DevExpress.XtraPivotGrid.Data.PivotFieldValueItem rowItem = pivot.Data.VisualItems.GetLastLevelItem(false, rowIndex, false);
        return  pivot.Data.VisualItems.CreateCellItem(columnItem, rowItem, columnIndex, rowIndex);
    }
    private static void FillDatasetColumns(ASPxPivotGrid pivot, DataTable dataTable1)
    {
        dataTable1.Columns.Add("RowFields", typeof(string));
        string tempColumnText = "";
        List<PivotGridFieldBase> fieldsInColumnArea = GetFieldsInArea(pivot, PivotArea.ColumnArea);

        for (int i = 0; i < pivot.ColumnCount; i++)
        {
            DevExpress.XtraPivotGrid.Data.PivotGridCellItem pcea = GetCellItem(pivot, i, 0);
            if (pcea.ColumnValueType == PivotGridValueType.Value) 
            {
                foreach (PivotGridFieldBase field in fieldsInColumnArea)
                    tempColumnText += pcea.GetFieldValue(GetFieldItemByField(pivot, field)).ToString() + " | ";//add formatting if it's necessary
                tempColumnText = tempColumnText.Remove(tempColumnText.Length - 3, 3);
                dataTable1.Columns.Add(tempColumnText, typeof(object));
                tempColumnText = "";
            }
            else
                dataTable1.Columns.Add(pcea.ColumnValueType.ToString(), typeof(object));
        }
    }
    private static List<PivotGridFieldBase> GetFieldsInArea(ASPxPivotGrid pivot, PivotArea area)
    {
        List<PivotGridFieldBase> fields = new List<PivotGridFieldBase>();
        for (int i = 0; i < pivot.Fields.Count; i++)
            if (pivot.Fields[i].Area == area)
                fields.Add(pivot.Fields[i]);
        return fields;
    }
    #endregion
    public static void InitBands(XtraReport rep)
    {
        // Create bands
        DetailBand detail = new DetailBand();
        PageHeaderBand pageHeader = new PageHeaderBand();
        ReportFooterBand reportFooter = new ReportFooterBand();
        detail.Height = 20;
        reportFooter.Height = 380;
        pageHeader.Height = 20;

        // Place the bands onto a report
        rep.Bands.AddRange(new DevExpress.XtraReports.UI.Band[] { detail, pageHeader, reportFooter });
    }
    public static void InitStyles(XtraReport rep)
    {
        // Create different odd and even styles
        XRControlStyle oddStyle = new XRControlStyle();
        XRControlStyle evenStyle = new XRControlStyle();

        // Specify the odd style appearance
        oddStyle.BackColor = Color.LightBlue;
        oddStyle.StyleUsing.UseBackColor = true;
        oddStyle.StyleUsing.UseBorders = false;
        oddStyle.Name = "OddStyle";

        // Specify the even style appearance
        evenStyle.BackColor = Color.LightPink;
        evenStyle.StyleUsing.UseBackColor = true;
        evenStyle.StyleUsing.UseBorders = false;
        evenStyle.Name = "EvenStyle";

        // Add styles to report's style sheet
        rep.StyleSheet.AddRange(new DevExpress.XtraReports.UI.XRControlStyle[] { oddStyle, evenStyle });
    }
    public static void InitDetailsBasedonXRTable(XtraReport rep, ReportGeneratorType kind, int columnWidth, bool repeatRowHeader)
    {
        if (!repeatRowHeader || kind == ReportGeneratorType.SinglePage)
            InitDetailsBasedonXRTableWithoutRepeatingRowHeader(rep, kind, columnWidth);
        else
            InitDetailsBasedonXRTableRepeatingRowHeader(rep, kind, columnWidth);
    }

    private static void InitDetailsBasedonXRTableRepeatingRowHeader(XtraReport rep, ReportGeneratorType kind, int columnWidth)
    {
        Font font = new Font("Tahoma", 9.75f);
        DataTable dataTable = ((DataSet)rep.DataSource).Tables[0];
        int processedPage = 0;
        int usablePageWidth = rep.PageWidth - (rep.Margins.Left + rep.Margins.Right);

        List<int> columnsWidth = null;
        if (kind == ReportGeneratorType.FixedColumnWidth)
            columnsWidth = DefineColumnsWidth(columnWidth, dataTable.Columns.Count);
        else
            columnsWidth = GetColumnsBestFitWidth(dataTable, font);

        XRTable tableHeader = null;
        XRTable tableDetail = null;
        InitNewTableInstancesAt(rep, font, out tableHeader, out tableDetail, new PointF(0, 0));
        tableHeader.BeginInit();
        tableDetail.BeginInit();
        int i = 1;
        AddCellsToTables(tableHeader, tableDetail, dataTable.Columns[0], columnsWidth[0], true);
        int remainingSpace = usablePageWidth - columnsWidth[0];
        do
        {
            if (columnsWidth[i] > remainingSpace)
            {
                processedPage++;
                tableHeader.WidthF = usablePageWidth - remainingSpace;
                tableDetail.WidthF = usablePageWidth - remainingSpace;
                tableHeader.EndInit();
                tableDetail.EndInit();
                InitNewTableInstancesAt(rep, font, out tableHeader, out tableDetail, new PointF(usablePageWidth * processedPage, 0));
                tableHeader.BeginInit();
                tableDetail.BeginInit();
                AddCellsToTables(tableHeader, tableDetail, dataTable.Columns[0], columnsWidth[0], true);
                remainingSpace = usablePageWidth - columnsWidth[0];
            }
            else
            {
                AddCellsToTables(tableHeader, tableDetail, dataTable.Columns[i], columnsWidth[i], false);
                remainingSpace -= columnsWidth[i];
                i++;
            }
        }
        while (i < columnsWidth.Count);
        tableHeader.WidthF = usablePageWidth - remainingSpace;
        tableDetail.WidthF = usablePageWidth - remainingSpace;
        tableHeader.EndInit();
        tableDetail.EndInit();
    }
    public static void AddCellsToTables(XRTable header, XRTable detail, DataColumn dc, int columnWidth, bool isFirstColumnInTable)
    {
        XRTableCell headerCell = new XRTableCell();
        headerCell.Text = dc.Caption;
        XRTableCell detailCell = new XRTableCell();
        detailCell.DataBindings.Add("Text", null, dc.Caption);
        headerCell.Width = columnWidth;
        detailCell.Width = columnWidth;
        if (isFirstColumnInTable)
        {
            headerCell.Borders = DevExpress.XtraPrinting.BorderSide.Left | DevExpress.XtraPrinting.BorderSide.Top | DevExpress.XtraPrinting.BorderSide.Bottom;
            detailCell.Borders = DevExpress.XtraPrinting.BorderSide.Left | DevExpress.XtraPrinting.BorderSide.Top | DevExpress.XtraPrinting.BorderSide.Bottom;
        }
        else
        {
            headerCell.Borders = DevExpress.XtraPrinting.BorderSide.All;
            detailCell.Borders = DevExpress.XtraPrinting.BorderSide.All;
        }
        // Place the cells into the corresponding tables
        header.Rows[0].Cells.Add(headerCell);
        detail.Rows[0].Cells.Add(detailCell);
    }
    public static void InitNewTableInstancesAt(XtraReport report, Font font, out XRTable header, out XRTable detail, PointF location)
    {
        header = InitXRTable(font, false);
        detail = InitXRTable(font, true);
        header.LocationF = location;
        detail.LocationF = location;
        XRTableRow headerRow = new XRTableRow();
        header.Rows.Add(headerRow);
        XRTableRow detailRow = new XRTableRow();
        detail.Rows.Add(detailRow);
        report.Bands[BandKind.PageHeader].Controls.Add(header);
        report.Bands[BandKind.Detail].Controls.Add(detail);
    }
    private static XRTable InitXRTable(Font font, bool withStyles)
    {
        XRTable table = new XRTable();
        table.Font = font;
        table.Height = 20;
        if (withStyles)
        {
            table.EvenStyleName = "EvenStyle";
            table.OddStyleName = "OddStyle";
        }
        return table;
    }

    private static List<int> DefineColumnsWidth(int columnWidth, int count)
    {
        List<int> columnsWidth = new List<int>();
        for (int i = 0; i < count; i++)
            columnsWidth.Add(columnWidth);
        return columnsWidth;
    }
     static void InitDetailsBasedonXRTableWithoutRepeatingRowHeader(XtraReport rep, ReportGeneratorType kind, int columnWidth)
    {
        Font font = new Font("Tahoma", 9.75f);
        DataSet ds = ((DataSet)rep.DataSource);
        int colCount = ds.Tables[0].Columns.Count;
        int colWidth = 0;



        XRTable tableHeader = null;
        XRTable tableDetail = null;
        InitNewTableInstancesAt(rep, font, out tableHeader, out tableDetail, new PointF(0, 0));


        List<int> columnsWidth = null;
        switch (kind)
        {
            case ReportGeneratorType.FixedColumnWidth:
                colWidth = columnWidth;
                tableHeader.Width = columnWidth * colCount;
                tableDetail.Width = columnWidth * colCount;
                break;
            case ReportGeneratorType.BestFitColumns:
                columnsWidth = GetColumnsBestFitWidth(ds.Tables[0], font);
                colWidth = 0;
                tableHeader.Width = GetTotalWidth(columnsWidth);
                tableDetail.Width = tableHeader.Width;
                break;
            default:
                colWidth = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right)) / colCount;
                tableHeader.Width = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right));
                tableDetail.Width = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right));
                break;
        }

        tableHeader.BeginInit();
        tableDetail.BeginInit();
        // Create table cells, fill the header cells with text, bind the cells to data
        for (int i = 0; i < colCount; i++)
        {
            AddCellsToTables(tableHeader, tableDetail, ds.Tables[0].Columns[i], kind == ReportGeneratorType.BestFitColumns ? columnsWidth[i] : colWidth, i == 0 ? true : false);
        }
        tableDetail.EndInit();
        tableHeader.EndInit();
        // Place the table onto a report's Detail band

    }

    private static int GetTotalWidth(List<int> columnsWidth)
    {
        int i = 0;
        foreach (int colWidth in columnsWidth)
            i += colWidth;
        return i;
    }

    private static List<int> GetColumnsBestFitWidth(DataTable dataTable, Font font)
    {
        List<int> optimalColumnWidth = new List<int>();
        float maxWidth = 0;
        float tempWidth = 0;
        for (int i = 1; i < dataTable.Rows.Count; i++)
        {
            tempWidth = TextRenderer.MeasureText(dataTable.Rows[i][0].ToString(), font).Width;
            maxWidth = maxWidth > tempWidth ? maxWidth : tempWidth;
        }
        optimalColumnWidth.Add(Convert.ToInt32(XRConvert.Convert(maxWidth, GraphicsUnit.Pixel, GraphicsUnit.Inch) * 100 + 1));
        for (int i = 1; i < dataTable.Columns.Count; i++)
        {
            tempWidth = TextRenderer.MeasureText(dataTable.Columns[i].ColumnName.ToString(), font).Width;
            maxWidth = 50 > tempWidth ? 50 : tempWidth;
            optimalColumnWidth.Add(Convert.ToInt32(XRConvert.Convert(maxWidth, GraphicsUnit.Pixel, GraphicsUnit.Inch) * 100 + 1));
        }
        return optimalColumnWidth;
    }
}
public enum ReportGeneratorType
{
    SinglePage,
    FixedColumnWidth,
    BestFitColumns,
}