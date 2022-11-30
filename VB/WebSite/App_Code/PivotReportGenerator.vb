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
Imports DevExpress.Drawing
Imports DevExpress.XtraPrinting

Public Module PivotReportGenerator
	Public Function GenerateReport(ByVal pivot As ASPxPivotGrid, ByVal kind As ReportGeneratorType, ByVal columnWidth As Integer, ByVal repeatRowHeader As Boolean) As XtraReport
		Dim rep As New XtraReport()
		rep.DataSource = FillDataset(pivot)
		rep.DataMember = CType(rep.DataSource, DataSet).Tables(0).TableName
		InitBands(rep)
		InitStyles(rep)
		InitDetailsBasedonXRTable(rep, kind, columnWidth, repeatRowHeader)
		Return rep
	End Function
	Public Function FillDataset(ByVal pivot As ASPxPivotGrid) As DataSet
		Dim dataSet1 As New DataSet()
		dataSet1.DataSetName = "PivotGridColumns"
		Dim dataTable1 As New DataTable()
		dataSet1.Tables.Add(dataTable1)
		FillDatasetColumns(pivot, dataTable1)
		FillDatasetExtracted(pivot, dataTable1)

		Return dataSet1
	End Function
	#Region "PreparingDataSet"
	Private Sub FillDatasetExtracted(ByVal pivot As ASPxPivotGrid, ByVal dataTable1 As DataTable)

		Dim rowvalues As New List(Of Object)()
		Dim tempRowText As String = ""
		Dim fieldsInRowArea As List(Of PivotGridFieldBase) = GetFieldsInArea(pivot, PivotArea.RowArea)
		For i As Integer = 0 To pivot.RowCount - 1
			Dim pcea As DevExpress.XtraPivotGrid.Data.PivotGridCellItem = GetCellItem(pivot, 0, i)
			If pcea.RowValueType = PivotGridValueType.Value Then
				For Each item As PivotGridFieldBase In fieldsInRowArea
					tempRowText &= pcea.GetFieldValue(item).ToString() & " | " 'add formatting if it's necessary
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

	Private Function GetCellItem(ByVal pivot As ASPxPivotGrid, ByVal columnIndex As Integer, ByVal rowIndex As Integer) As DevExpress.XtraPivotGrid.Data.PivotGridCellItem
		Dim columnItem As DevExpress.XtraPivotGrid.Data.PivotFieldValueItem = pivot.Data.VisualItems.GetLastLevelItem(True, columnIndex, False)
		Dim rowItem As DevExpress.XtraPivotGrid.Data.PivotFieldValueItem = pivot.Data.VisualItems.GetLastLevelItem(False, rowIndex, False)
		Return pivot.Data.VisualItems.CreateCellItem(columnItem, rowItem, columnIndex, rowIndex)
	End Function
	Private Sub FillDatasetColumns(ByVal pivot As ASPxPivotGrid, ByVal dataTable1 As DataTable)
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
	Private Function GetFieldsInArea(ByVal pivot As ASPxPivotGrid, ByVal area As PivotArea) As List(Of PivotGridFieldBase)
		Dim fields As New List(Of PivotGridFieldBase)()
		For i As Integer = 0 To pivot.Fields.Count - 1
			If pivot.Fields(i).Area = area Then
				fields.Add(pivot.Fields(i))
			End If
		Next i
		Return fields
	End Function
	#End Region
	Public Sub InitBands(ByVal rep As XtraReport)
		' Create bands
		Dim detail As New DetailBand()
		Dim pageHeader As New PageHeaderBand()
		Dim reportFooter As New ReportFooterBand()
		detail.Height = 20
		reportFooter.Height = 380
		pageHeader.Height = 20

		' Place the bands onto a report
		rep.Bands.AddRange(New Band() { detail, pageHeader, reportFooter })
	End Sub
	Public Sub InitStyles(ByVal rep As XtraReport)
		' Create different odd and even styles
		Dim oddStyle As New XRControlStyle()
		Dim evenStyle As New XRControlStyle()

		' Specify the odd style appearance
		oddStyle.BackColor = System.Drawing.Color.LightBlue
		oddStyle.StyleUsing.UseBackColor = True
		oddStyle.StyleUsing.UseBorders = False
		oddStyle.Name = "OddStyle"

		' Specify the even style appearance
		evenStyle.BackColor = System.Drawing.Color.LightPink
		evenStyle.StyleUsing.UseBackColor = True
		evenStyle.StyleUsing.UseBorders = False
		evenStyle.Name = "EvenStyle"

		' Add styles to report's style sheet
		rep.StyleSheet.AddRange(New XRControlStyle() { oddStyle, evenStyle })
	End Sub
	Public Sub InitDetailsBasedonXRTable(ByVal rep As XtraReport, ByVal kind As ReportGeneratorType, ByVal columnWidth As Single, ByVal repeatRowHeader As Boolean)
		If Not repeatRowHeader OrElse kind = ReportGeneratorType.SinglePage Then
			InitDetailsBasedonXRTableWithoutRepeatingRowHeader(rep, kind, columnWidth)
		Else
			InitDetailsBasedonXRTableRepeatingRowHeader(rep, kind, columnWidth)
		End If
	End Sub

	Private Sub InitDetailsBasedonXRTableRepeatingRowHeader(ByVal rep As XtraReport, ByVal kind As ReportGeneratorType, ByVal columnWidth As Single)
		Dim font As New DXFont("Tahoma", 9.75F)
		Dim dataTable As DataTable = CType(rep.DataSource, DataSet).Tables(0)
		Dim processedPage As Integer = 0
		Dim usablePageWidth As Single = rep.PageWidth - (rep.Margins.Left + rep.Margins.Right)

		Dim columnsWidth As List(Of Single) = Nothing
		If kind = ReportGeneratorType.FixedColumnWidth Then
			columnsWidth = DefineColumnsWidth(columnWidth, dataTable.Columns.Count)
		Else
			columnsWidth = GetColumnsBestFitWidth(dataTable, font, rep.ReportUnit)
		End If

		Dim tableHeader As XRTable = Nothing
		Dim tableDetail As XRTable = Nothing
		InitNewTableInstancesAt(rep, font, tableHeader, tableDetail, New PointF(0, 0))
		tableHeader.BeginInit()
		tableDetail.BeginInit()
		Dim i As Integer = 1
		AddCellsToTables(tableHeader, tableDetail, dataTable.Columns(0), columnsWidth(0), True)
		Dim remainingSpace As Single = usablePageWidth - columnsWidth(0)
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
	Public Sub AddCellsToTables(ByVal header As XRTable, ByVal detail As XRTable, ByVal dc As DataColumn, ByVal columnWidth As Single, ByVal isFirstColumnInTable As Boolean)
		Dim headerCell As New XRTableCell()
		headerCell.Text = dc.Caption
		Dim detailCell As New XRTableCell()
		detailCell.DataBindings.Add("Text", Nothing, dc.Caption)
		headerCell.WidthF = columnWidth
		detailCell.WidthF = columnWidth
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
	Public Sub InitNewTableInstancesAt(ByVal report As XtraReport, ByVal font As DXFont, <System.Runtime.InteropServices.Out()> ByRef header As XRTable, <System.Runtime.InteropServices.Out()> ByRef detail As XRTable, ByVal location As PointF)
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
	Private Function InitXRTable(ByVal font As DXFont, ByVal withStyles As Boolean) As XRTable
		Dim table As New XRTable()
		table.Font = font
		table.Height = 20
		If withStyles Then
			table.EvenStyleName = "EvenStyle"
			table.OddStyleName = "OddStyle"
		End If
		Return table
	End Function

	Private Function DefineColumnsWidth(ByVal columnWidth As Single, ByVal count As Integer) As List(Of Single)
		Dim columnsWidth As New List(Of Single)()
		For i As Integer = 0 To count - 1
			columnsWidth.Add(columnWidth)
		Next i
		Return columnsWidth
	End Function
	Private Sub InitDetailsBasedonXRTableWithoutRepeatingRowHeader(ByVal rep As XtraReport, ByVal kind As ReportGeneratorType, ByVal columnWidth As Single)
		Dim font As New DXFont("Tahoma", 9.75F)
		Dim ds As DataSet = (CType(rep.DataSource, DataSet))
		Dim colCount As Integer = ds.Tables(0).Columns.Count
		Dim colWidth As Single = 0



		Dim tableHeader As XRTable = Nothing
		Dim tableDetail As XRTable = Nothing
		InitNewTableInstancesAt(rep, font, tableHeader, tableDetail, New PointF(0, 0))


		Dim columnsWidth As List(Of Single) = Nothing
		Select Case kind
			Case ReportGeneratorType.FixedColumnWidth
				colWidth = columnWidth
				tableHeader.WidthF = columnWidth * colCount
				tableDetail.WidthF = columnWidth * colCount
			Case ReportGeneratorType.BestFitColumns
				columnsWidth = GetColumnsBestFitWidth(ds.Tables(0), font, rep.ReportUnit)
				colWidth = 0
				tableHeader.WidthF = GetTotalWidth(columnsWidth)
				tableDetail.WidthF = tableHeader.Width
			Case Else
'INSTANT VB WARNING: Instant VB cannot determine whether both operands of this division are integer types - if they are then you should use the VB integer division operator:
				colWidth = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right)) / colCount
				tableHeader.WidthF = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right))
				tableDetail.WidthF = (rep.PageWidth - (rep.Margins.Left + rep.Margins.Right))
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

	Private Function GetTotalWidth(ByVal columnsWidth As List(Of Single)) As Single
		Dim i As Single = 0
		For Each colWidth As Single In columnsWidth
			i += colWidth
		Next colWidth
		Return i
	End Function

	Private Function GetColumnsBestFitWidth(ByVal dataTable As DataTable, ByVal font As DXFont, ByVal unit As ReportUnit) As List(Of Single)
		Dim optimalColumnWidth As New List(Of Single)()
		Dim maxWidth As Single = 0
		Dim tempWidth As Single = 0
		For i As Integer = 1 To dataTable.Rows.Count - 1
			tempWidth = MeasureWidth(dataTable.Rows(i)(0).ToString(), font, unit)
			maxWidth = If(maxWidth > tempWidth, maxWidth, tempWidth)
		Next i
		optimalColumnWidth.Add(maxWidth)
		For i As Integer = 1 To dataTable.Columns.Count - 1
			tempWidth = MeasureWidth(dataTable.Columns(i).ColumnName.ToString(), font, unit)
			maxWidth = If(50 > tempWidth, 50, tempWidth)
			optimalColumnWidth.Add(maxWidth)
		Next i
		Return optimalColumnWidth
	End Function

	Private Function MeasureWidth(ByVal candidate As String, ByVal font As DXFont, ByVal unit As ReportUnit) As Single
		Return BestSizeEstimator.GetBoundsToFitText(candidate, New BrickStyle() With {.Font = font}, unit).Width
	End Function
End Module
Public Enum ReportGeneratorType
	SinglePage
	FixedColumnWidth
	BestFitColumns
End Enum