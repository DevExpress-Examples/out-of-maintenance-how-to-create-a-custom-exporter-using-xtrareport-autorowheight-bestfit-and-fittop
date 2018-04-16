Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraPivotGrid
Imports DevExpress.XtraEditors
Imports DevExpress.Web.ASPxPivotGrid
Imports System.Text

Public NotInheritable Class PivotReportGenerator

    Private Sub New()
    End Sub

    Public Shared Function GenerateReport(ByVal pivot As ASPxPivotGrid, ByVal kind As ReportGeneratorType, ByVal columnWidth As Integer, ByVal repeatRowHeader As Boolean) As XtraReport
            Dim rep As New XtraReport()
            rep.DataSource = FillDataset(pivot)
            rep.DataMember = CType(rep.DataSource, DataSet).Tables(0).TableName
            InitBands(rep)
            InitStyles(rep)
            InitDetailsBasedonXRTable(rep, kind, columnWidth, repeatRowHeader)
            Return rep
    End Function
    Public Shared Function FillDataset(ByVal pivot As ASPxPivotGrid) As DataSet
        Dim dataSet1 As New DataSet()
        dataSet1.DataSetName = "PivotGridColumns"
        Dim dataTable1 As New DataTable()
        dataSet1.Tables.Add(dataTable1)
        FillDatasetColumns(pivot, dataTable1)
        FillDatasetExtracted(pivot, dataTable1)

        Return dataSet1
    End Function
    #Region "PreparingDataSet"
    Private Shared Sub FillDatasetExtracted(ByVal pivot As ASPxPivotGrid, ByVal dataTable1 As DataTable)

        Dim rowvalues As New List(Of Object)()
        Dim tempRowText As String = ""
        Dim fieldsInRowArea As List(Of PivotGridFieldBase) = GetFieldsInArea(pivot, PivotArea.RowArea)
        For i As Integer = 0 To pivot.RowCount - 1
            Dim pcea As DevExpress.XtraPivotGrid.Data.PivotGridCellItem = GetCellItem(pivot, 0, i)
            If pcea.RowValueType = PivotGridValueType.Value Then
                For Each item As PivotGridFieldBase In fieldsInRowArea
                    tempRowText &= pcea.GetFieldValue(GetFieldItemByField(pivot, item)).ToString() & " | " 'add formatting if it's necessary
                Next item
                tempRowText = tempRowText.Remove(tempRowText.Length - 3, 3)
            Else
                tempRowText = pcea.RowValueType.ToString()
            End If
            rowvalues.Clear()
            rowvalues.Add(tempRowText)
            tempRowText = ""
            For j As Integer = 0 To pivot.ColumnCount - 1
                pcea = GetCellItem(pivot, j, i)
                If pcea.Value IsNot Nothing Then
                    rowvalues.Add(pcea.Value)
                Else
                    rowvalues.Add(DBNull.Value)
                End If
            Next j
            dataTable1.Rows.Add(rowvalues.ToArray())
        Next i
    End Sub

    Private Shared Function GetFieldItemByField(ByVal pivot As ASPxPivotGrid, ByVal field As PivotGridFieldBase) As DevExpress.XtraPivotGrid.Data.PivotFieldItemBase
        Return pivot.Data.GetFieldItem(field)
    End Function



    Private Shared Function GetCellItem(ByVal pivot As ASPxPivotGrid, ByVal columnIndex As Integer, ByVal rowIndex As Integer) As DevExpress.XtraPivotGrid.Data.PivotGridCellItem
        Dim columnItem As DevExpress.XtraPivotGrid.Data.PivotFieldValueItem = pivot.Data.VisualItems.GetLastLevelItem(True, columnIndex, False)
        Dim rowItem As DevExpress.XtraPivotGrid.Data.PivotFieldValueItem = pivot.Data.VisualItems.GetLastLevelItem(False, rowIndex, False)
        Return pivot.Data.VisualItems.CreateCellItem(columnItem, rowItem, columnIndex, rowIndex)
    End Function
    Private Shared Sub FillDatasetColumns(ByVal pivot As ASPxPivotGrid, ByVal dataTable1 As DataTable)
        dataTable1.Columns.Add("RowFields", GetType(String))
        Dim sb As New StringBuilder()
        Dim multipleDataField As Boolean = pivot.GetFieldsByArea(PivotArea.DataArea).Count > 1
        For i As Integer = 0 To pivot.ColumnCount - 1
            Dim pcea As DevExpress.Web.ASPxPivotGrid.PivotCellBaseEventArgs = pivot.GetCellInfo(i, 0)
            For Each field As DevExpress.Web.ASPxPivotGrid.PivotGridField In pcea.GetColumnFields()
                sb.AppendFormat("{0} | ", field.GetDisplayText(pcea.GetFieldValue(field))) 'add formatting if it's necessary
            Next field
            If multipleDataField Then
                sb.AppendFormat("{0} | ", pcea.DataField)
            End If
            If pcea.ColumnValueType = PivotGridValueType.Value Then
                sb.Remove(sb.Length - 3, 3)
            Else
                sb.Append(pcea.ColumnValueType.ToString())
            End If
            dataTable1.Columns.Add(sb.ToString(), GetType(Object))
            sb.Clear()
        Next i
    End Sub
    Private Shared Function GetFieldsInArea(ByVal pivot As ASPxPivotGrid, ByVal area As PivotArea) As List(Of PivotGridFieldBase)
        Dim fields As New List(Of PivotGridFieldBase)()
        For i As Integer = 0 To pivot.Fields.Count - 1
            If pivot.Fields(i).Area = area Then
                fields.Add(pivot.Fields(i))
            End If
        Next i
        Return fields
    End Function
    #End Region
    Public Shared Sub InitBands(ByVal rep As XtraReport)
        ' Create bands
        Dim detail As New DetailBand()
        Dim pageHeader As New PageHeaderBand()
        Dim reportFooter As New ReportFooterBand()
        detail.Height = 20
        reportFooter.Height = 380
        pageHeader.Height = 20

        ' Place the bands onto a report
        rep.Bands.AddRange(New DevExpress.XtraReports.UI.Band() { detail, pageHeader, reportFooter })
    End Sub
    Public Shared Sub InitStyles(ByVal rep As XtraReport)
        ' Create different odd and even styles
        Dim oddStyle As New XRControlStyle()
        Dim evenStyle As New XRControlStyle()

        ' Specify the odd style appearance
        oddStyle.BackColor = Color.LightBlue
        oddStyle.StyleUsing.UseBackColor = True
        oddStyle.StyleUsing.UseBorders = False
        oddStyle.Name = "OddStyle"

        ' Specify the even style appearance
        evenStyle.BackColor = Color.LightPink
        evenStyle.StyleUsing.UseBackColor = True
        evenStyle.StyleUsing.UseBorders = False
        evenStyle.Name = "EvenStyle"

        ' Add styles to report's style sheet
        rep.StyleSheet.AddRange(New DevExpress.XtraReports.UI.XRControlStyle() { oddStyle, evenStyle })
    End Sub
    Public Shared Sub InitDetailsBasedonXRTable(ByVal rep As XtraReport, ByVal kind As ReportGeneratorType, ByVal columnWidth As Integer, ByVal repeatRowHeader As Boolean)
        If (Not repeatRowHeader) OrElse kind = ReportGeneratorType.SinglePage Then
            InitDetailsBasedonXRTableWithoutRepeatingRowHeader(rep, kind, columnWidth)
        Else
            InitDetailsBasedonXRTableRepeatingRowHeader(rep, kind, columnWidth)
        End If
    End Sub

    Private Shared Sub InitDetailsBasedonXRTableRepeatingRowHeader(ByVal rep As XtraReport, ByVal kind As ReportGeneratorType, ByVal columnWidth As Integer)
        Dim font As New Font("Tahoma", 9.75F)
        Dim dataTable As DataTable = CType(rep.DataSource, DataSet).Tables(0)
        Dim processedPage As Integer = 0
        Dim usablePageWidth As Integer = rep.PageWidth - (rep.Margins.Left + rep.Margins.Right)

        Dim columnsWidth As List(Of Integer) = Nothing
        If kind = ReportGeneratorType.FixedColumnWidth Then
            columnsWidth = DefineColumnsWidth(columnWidth, dataTable.Columns.Count)
        Else
            columnsWidth = GetColumnsBestFitWidth(dataTable, font)
        End If

        Dim tableHeader As XRTable = Nothing
        Dim tableDetail As XRTable = Nothing
        InitNewTableInstancesAt(rep, font, tableHeader, tableDetail, New PointF(0, 0))
        tableHeader.BeginInit()
        tableDetail.BeginInit()
        Dim i As Integer = 1
        AddCellsToTables(tableHeader, tableDetail, dataTable.Columns(0), columnsWidth(0), True)
        Dim remainingSpace As Integer = usablePageWidth - columnsWidth(0)
        Do
            If columnsWidth(i) > remainingSpace Then
                processedPage += 1
                tableHeader.WidthF = usablePageWidth - remainingSpace
                tableDetail.WidthF = usablePageWidth - remainingSpace
                tableHeader.EndInit()
                tableDetail.EndInit()
                InitNewTableInstancesAt(rep, font, tableHeader, tableDetail, New PointF(usablePageWidth * processedPage, 0))
                tableHeader.BeginInit()
                tableDetail.BeginInit()
                AddCellsToTables(tableHeader, tableDetail, dataTable.Columns(0), columnsWidth(0), True)
                remainingSpace = usablePageWidth - columnsWidth(0)
            Else
                AddCellsToTables(tableHeader, tableDetail, dataTable.Columns(i), columnsWidth(i), False)
                remainingSpace -= columnsWidth(i)
                i += 1
            End If
        Loop While i < columnsWidth.Count
        tableHeader.WidthF = usablePageWidth - remainingSpace
        tableDetail.WidthF = usablePageWidth - remainingSpace
        tableHeader.EndInit()
        tableDetail.EndInit()
    End Sub
    Public Shared Sub AddCellsToTables(ByVal header As XRTable, ByVal detail As XRTable, ByVal dc As DataColumn, ByVal columnWidth As Integer, ByVal isFirstColumnInTable As Boolean)
        Dim headerCell As New XRTableCell()
        headerCell.Text = dc.Caption
        Dim detailCell As New XRTableCell()
        detailCell.DataBindings.Add("Text", Nothing, dc.Caption)
        headerCell.Width = columnWidth
        detailCell.Width = columnWidth
        If isFirstColumnInTable Then
            headerCell.Borders = DevExpress.XtraPrinting.BorderSide.Left Or DevExpress.XtraPrinting.BorderSide.Top Or DevExpress.XtraPrinting.BorderSide.Bottom
            detailCell.Borders = DevExpress.XtraPrinting.BorderSide.Left Or DevExpress.XtraPrinting.BorderSide.Top Or DevExpress.XtraPrinting.BorderSide.Bottom
        Else
            headerCell.Borders = DevExpress.XtraPrinting.BorderSide.All
            detailCell.Borders = DevExpress.XtraPrinting.BorderSide.All
        End If
        ' Place the cells into the corresponding tables
        header.Rows(0).Cells.Add(headerCell)
        detail.Rows(0).Cells.Add(detailCell)
    End Sub
    Public Shared Sub InitNewTableInstancesAt(ByVal report As XtraReport, ByVal font As Font, <System.Runtime.InteropServices.Out()> ByRef header As XRTable, <System.Runtime.InteropServices.Out()> ByRef detail As XRTable, ByVal location As PointF)
        header = InitXRTable(font, False)
        detail = InitXRTable(font, True)
        header.LocationF = location
        detail.LocationF = location
        Dim headerRow As New XRTableRow()
        header.Rows.Add(headerRow)
        Dim detailRow As New XRTableRow()
        detail.Rows.Add(detailRow)
        report.Bands(BandKind.PageHeader).Controls.Add(header)
        report.Bands(BandKind.Detail).Controls.Add(detail)
    End Sub
    Private Shared Function InitXRTable(ByVal font As Font, ByVal withStyles As Boolean) As XRTable
        Dim table As New XRTable()
        table.Font = font
        table.Height = 20
        If withStyles Then
            table.EvenStyleName = "EvenStyle"
            table.OddStyleName = "OddStyle"
        End If
        Return table
    End Function

    Private Shared Function DefineColumnsWidth(ByVal columnWidth As Integer, ByVal count As Integer) As List(Of Integer)
        Dim columnsWidth As New List(Of Integer)()
        For i As Integer = 0 To count - 1
            columnsWidth.Add(columnWidth)
        Next i
        Return columnsWidth
    End Function
     Private Shared Sub InitDetailsBasedonXRTableWithoutRepeatingRowHeader(ByVal rep As XtraReport, ByVal kind As ReportGeneratorType, ByVal columnWidth As Integer)
        Dim font As New Font("Tahoma", 9.75F)
        Dim ds As DataSet = (CType(rep.DataSource, DataSet))
        Dim colCount As Integer = ds.Tables(0).Columns.Count
        Dim colWidth As Integer = 0



        Dim tableHeader As XRTable = Nothing
        Dim tableDetail As XRTable = Nothing
        InitNewTableInstancesAt(rep, font, tableHeader, tableDetail, New PointF(0, 0))


        Dim columnsWidth As List(Of Integer) = Nothing
        Select Case kind
            Case ReportGeneratorType.FixedColumnWidth
                colWidth = columnWidth
                tableHeader.Width = columnWidth * colCount
                tableDetail.Width = columnWidth * colCount
            Case ReportGeneratorType.BestFitColumns
                columnsWidth = GetColumnsBestFitWidth(ds.Tables(0), font)
                colWidth = 0
                tableHeader.Width = GetTotalWidth(columnsWidth)
                tableDetail.Width = tableHeader.Width
            Case Else
                colWidth = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right)) / colCount
                tableHeader.Width = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right))
                tableDetail.Width = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right))
        End Select

        tableHeader.BeginInit()
        tableDetail.BeginInit()
        ' Create table cells, fill the header cells with text, bind the cells to data
        For i As Integer = 0 To colCount - 1
            AddCellsToTables(tableHeader, tableDetail, ds.Tables(0).Columns(i),If(kind = ReportGeneratorType.BestFitColumns, columnsWidth(i), colWidth),If(i = 0, True, False))
        Next i
        tableDetail.EndInit()
        tableHeader.EndInit()
        ' Place the table onto a report's Detail band

     End Sub

    Private Shared Function GetTotalWidth(ByVal columnsWidth As List(Of Integer)) As Integer
        Dim i As Integer = 0
        For Each colWidth As Integer In columnsWidth
            i += colWidth
        Next colWidth
        Return i
    End Function

    Private Shared Function GetColumnsBestFitWidth(ByVal dataTable As DataTable, ByVal font As Font) As List(Of Integer)
        Dim optimalColumnWidth As New List(Of Integer)()
        Dim maxWidth As Single = 0
        Dim tempWidth As Single = 0
        For i As Integer = 1 To dataTable.Rows.Count - 1
            tempWidth = TextRenderer.MeasureText(dataTable.Rows(i)(0).ToString(), font).Width
            maxWidth = If(maxWidth > tempWidth, maxWidth, tempWidth)
        Next i
        optimalColumnWidth.Add(Convert.ToInt32(XRConvert.Convert(maxWidth, GraphicsUnit.Pixel, GraphicsUnit.Inch) * 100 + 1))
        For i As Integer = 1 To dataTable.Columns.Count - 1
            tempWidth = TextRenderer.MeasureText(dataTable.Columns(i).ColumnName.ToString(), font).Width
            maxWidth = If(50 > tempWidth, 50, tempWidth)
            optimalColumnWidth.Add(Convert.ToInt32(XRConvert.Convert(maxWidth, GraphicsUnit.Pixel, GraphicsUnit.Inch) * 100 + 1))
        Next i
        Return optimalColumnWidth
    End Function
End Class
Public Enum ReportGeneratorType
    SinglePage
    FixedColumnWidth
    BestFitColumns
End Enum