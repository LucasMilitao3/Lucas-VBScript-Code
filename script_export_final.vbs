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
 ' Get the tables collection
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
 sheet.Columns(2).ColumnWidth = 40 
 sheet.Columns(3).ColumnWidth = 30
 sheet.Columns(4).ColumnWidth = 50 
 sheet.Columns(5).ColumnWidth = 40 
 sheet.Columns(6).ColumnWidth = 40
 sheet.Columns(7).ColumnWidth = 40 
 sheet.Columns(1).WrapText =true
 sheet.Columns(2).WrapText =true
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
      sheet.cells(rowsNum, 1) = tab.name
      sheet.cells(rowsNum, 2) = tab.code
	  sheet.cells(rowsNum, 3) = tab.comment
      sheet.cells(rowsNum, 4) = col.name
      sheet.cells(rowsNum, 5) = col.code
      sheet.cells(rowsNum, 6) = col.datatype
	  sheet.cells(rowsNum, 7) = col.comment
      next
      sheet.Range(sheet.cells(rowsNum-colsNum+1,1),sheet.cells(rowsNum,7)).Borders.LineStyle = "2"      
 
      Output "FullDescription: "       + tab.Name
End Sub