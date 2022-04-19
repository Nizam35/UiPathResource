# UiPathResource - Useful Code Snippet

 ## Gembox Code to Color the Header and Set Default Width to each Column
``` VB.Net
Try
 SpreadsheetInfo.SetLicense(in_GemboxLicenseKey)
 
Dim tempWorkBook As ExcelFile = ExcelFile.Load(in_ExcelFilePath)

Dim tempSheetName As ExcelWorksheet =tempWorkBook.Worksheets(in_ExcelSheetName)
Dim Range As CellRange
																
'Color the Header
Range = tempSheetName.Cells.GetSubrange("A1", "I1")
Range.Style.FillPattern.SetPattern(FillPatternStyle.Solid,SpreadsheetColor.FromArgb(255, 150, 10), SpreadsheetColor.FromName(ColorName.Black))				
Range.Style.Font.Color= SpreadsheetColor.FromName(ColorName.Black)
Range.Style.Font.Weight = 600

'Autofit Columns
Dim columnCount As Integer = tempSheetName.CalculateMaxUsedColumns()
Dim colIndex As Integer 

For colIndex = 0 To columnCount - 1
            tempSheetName.Columns(colIndex).SetWidth(Math.Truncate((20* 7 + 5) / 7 * 256) / 256,LengthUnit.ZeroCharacterWidth)
 Next
 
tempWorkBook.save(in_ExcelFilePath) 

Catch Ex As Exception
               Console.WriteLine(ex.Message)
 End Try           
```


## Working With DataTable 

### Linq To Perform LeftJoin with Two Datatables

1. Prepare the Result Datatable 
	
``` Linq

out_TransactionDt = DT1.Clone

// Make sure that  out_TransactionDt has all the required columns. use Add Data Column activity if need to add additional Columns

( From dt1Row In Dt1.AsEnumerable()
  Group Join dt2Row In DT2.AsEnumerable()
   On dt1row.Item("ColumnName").ToString Equals dt2Row.item("ColumnName") Into Group
   Let matchedfirstRow = Group.FirstOrDefault()
    Select ra = { 
		      dt1Row("columnName"),
			   dt1Row("columnName"),
			    dt1Row("columnName"),
				dt1Row("columnName"),
			   If(isNothing(matchedfirstRow), Nothing, matchedfirstRow("columnName")),
			    If(isNothing(matchedfirstRow), Nothing, matchedfirstRow("columnName")),
				If(isNothing(matchedfirstRow), Nothing, matchedfirstRow("columnName"))
				}
	Select out_TransactionDt.Rows.Add(ra)).CopyToDataTable
```






