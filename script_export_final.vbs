Option Explicit
   Dim rowsNum
   rowsNum = 0
'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------
' Get the current active model
Dim Model
Set Model = ActiveModel
If (Model Is Nothing) Or (Not Model.IsKindOf(PdPDM.cls_Model)) Then
  MsgBox "The current model is not an PDM model."
Else
 'Get the tables collection
 'Establish EXCEL APP
 dim beginrow
 DIM EXCEL, SHEET
 set EXCEL = CREATEOBJECT("Excel.Application")
 EXCEL.workbooks.add(-4167)'Add sheet
 EXCEL.workbooks(1).sheets(1).name ="Export"
 set sheet = EXCEL.workbooks(1).sheets("Export")
 
 ShowProperties Model, SHEET
 EXCEL.visible = true
 'Set column width and wrap
 sheet.Columns(1).ColumnWidth = 40 
 sheet.Columns(2).ColumnWidth = 80 
 sheet.Columns(3).ColumnWidth = 40
 sheet.Columns(4).ColumnWidth = 40 
 sheet.Columns(5).ColumnWidth = 40 
 sheet.Columns(6).ColumnWidth = 40
 sheet.Columns(7).ColumnWidth = 40
 sheet.Columns(8).ColumnWidth = 80
 sheet.Columns(9).ColumnWidth = 40
 sheet.Columns(10).ColumnWidth = 40
 sheet.Columns(11).ColumnWidth = 40
 sheet.Columns(12).ColumnWidth = 80
 sheet.Columns(13).ColumnWidth = 40
 sheet.Columns(14).ColumnWidth = 40
 sheet.Columns(15).ColumnWidth = 40
 sheet.Columns(1).WrapText =true
 sheet.Columns(2).WrapText =true
 sheet.Columns(3).WrapText =true
 sheet.Columns(4).WrapText =true
 sheet.Columns(5).WrapText =true
 sheet.Columns(6).WrapText =true
 sheet.Columns(7).WrapText =true
 sheet.Columns(8).WrapText =true
 sheet.Columns(9).WrapText =true
 sheet.Columns(10).WrapText =true
 sheet.Columns(11).WrapText =true
 sheet.Columns(12).WrapText =true
 sheet.Columns(13).WrapText =true
 sheet.Columns(14).WrapText =true
 sheet.Columns(15).WrapText =true
 
 End If
'-----------------------------------------------------------------------------
' Show properties of tables
'-----------------------------------------------------------------------------
Sub ShowProperties(mdl, sheet)
   ' Show tables of the current model/package
   rowsNum=0
   beginrow = rowsNum+1
   ' For each table
   output "begin"
   Dim tab
   For Each tab In mdl.tables
      ShowTable tab,sheet
   Next
   if mdl.tables.count > 0 then
        sheet.Range("A" & beginrow + 1 & ":A" & rowsNum).Rows.Group
   end if
   output "end"
End Sub
'-----------------------------------------------------------------------------
' Show table properties
'-----------------------------------------------------------------------------

Sub ShowTable(tab, sheet)

     Dim rangFlag

Dim col
Dim colsNum
colsNum = 0
for each col in tab.columns
      rowsNum = rowsNum + 1
      colsNum = colsNum + 1
	  'dados das tabelas
		sheet.cells(rowsNum, 1) = tab.name
		sheet.cells(rowsNum, 2) = tab.comment
		sheet.cells(rowsNum, 3) = tab.GetExtendedAttribute("Axon Viewing",1)
		sheet.cells(rowsNum, 4) = tab.GetExtendedAttribute("DataSet Type",1)
		sheet.cells(rowsNum, 5) = tab.GetExtendedAttribute("Sigla Sistema",1)
		sheet.cells(rowsNum, 6) = tab.GetExtendedAttribute("Lifecycle Axon",1)
		sheet.cells(rowsNum, 7) = tab.GetExtendedAttribute("Axon Status",1)
		sheet.cells(rowsNum, 8) = tab.GetExtendedAttribute("Glossário AXON",1)
		'dados das colunas
		sheet.cells(rowsNum, 9) = col.name
		sheet.cells(rowsNum, 10) = col.code
		sheet.cells(rowsNum, 11) = col.comment
		sheet.cells(rowsNum, 12) = col.GetExtendedAttribute("Glossário AXON",1)
		sheet.cells(rowsNum, 13) = col.table
		sheet.cells(rowsNum, 14) = col.datatype
		sheet.cells(rowsNum, 15) = col.GetExtendedAttribute("Sigla Sistema",1)
	  
      next
      sheet.Range(sheet.cells(rowsNum-colsNum+1,1),sheet.cells(rowsNum,8)).Borders.LineStyle = "2"      
 
      Output "FullDescription: "       + tab.Name
End Sub
