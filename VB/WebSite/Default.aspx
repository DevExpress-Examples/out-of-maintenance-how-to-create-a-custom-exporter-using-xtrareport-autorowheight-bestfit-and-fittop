<%@ Page Language="vb" AutoEventWireup="true" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<%@ Register Assembly="DevExpress.Web.ASPxPivotGrid.v10.1.Export, Version=10.1.7.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
	Namespace="DevExpress.XtraPivotGrid.Web" TagPrefix="dxpgw" %>
<%@ Register Assembly="DevExpress.Web.ASPxPivotGrid.v10.1, Version=10.1.7.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
	Namespace="DevExpress.Web.ASPxPivotGrid" TagPrefix="dxwpg" %>
<%@ Register Assembly="DevExpress.Web.ASPxEditors.v10.1, Version=10.1.7.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
	Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>Untitled Page</title>
</head>
<body>
	<form id="form1" runat="server">
	<div>
		<table>
			<tr>
				<td>
					<asp:Button ID="Button2" runat="server" OnClick="Button1_Click" Text="Export" />
				</td>
				<td>
					<dx:ASPxRadioButtonList ID="RadioButtonList1" runat="server" SelectedIndex="2">
						<Items>
							<dx:ListEditItem Selected="True" Text="autosize to page width" Value="SinglePage" />
							<dx:ListEditItem Text="do not autosize and use column width" Value="FixedColumnWidth" />
							<dx:ListEditItem Text="do not autosize and best fit columns" Value="BestFitColumns" />
						</Items>
					</dx:ASPxRadioButtonList>
				</td>
				<td>                   
					<dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="Column Width">
					</dx:ASPxLabel>
					<dx:ASPxSpinEdit ID="ASPxSpinEdit1" runat="server" Height="21px" Number="40" />
					<dx:ASPxCheckBox ID="ASPxCheckBox1" runat="server" Text="Repeat row header on every page">
					</dx:ASPxCheckBox>
				</td>
			</tr>
		</table>
		<dxwpg:ASPxPivotGrid ID="ASPxPivotGrid1" runat="server" DataSourceID="AccessDataSource1">
			<Fields>
				<dxwpg:PivotGridField ID="fieldCompanyName" Area="RowArea" AreaIndex="0" FieldName="CompanyName">
				</dxwpg:PivotGridField>
				<dxwpg:PivotGridField ID="fieldProductAmount" Area="DataArea" AreaIndex="0" FieldName="ProductAmount">
				</dxwpg:PivotGridField>
				<dxwpg:PivotGridField ID="fieldOrderDate" Area="FilterArea" AreaIndex="0" FieldName="OrderDate"
					GroupInterval="DateYear" UnboundFieldName="fieldOrderDate">
				</dxwpg:PivotGridField>
				<dxwpg:PivotGridField ID="fieldProductName" Area="ColumnArea" AreaIndex="0" FieldName="ProductName"
					TopValueCount="10">
				</dxwpg:PivotGridField>
			</Fields>
		</dxwpg:ASPxPivotGrid>
		<dxpgw:ASPxPivotGridExporter ID="ASPxPivotGridExporter1" runat="server" ASPxPivotGridID="ASPxPivotGrid1">
		</dxpgw:ASPxPivotGridExporter>
		<asp:AccessDataSource ID="AccessDataSource1" runat="server" DataFile="~/App_Data/nwind.mdb"
			SelectCommand="SELECT [CompanyName], [ProductAmount], [OrderDate], [ProductName] FROM [CustomerReports]">
		</asp:AccessDataSource>
	</div>
	</form>
</body>
</html>