### Code Snippet Specific To Datatable

## Using DataView to Sort Datatable

``` Sort
Datatable.DefaultView.Sort = "ColName ASC,colName DESC"
DataTable = Datatable.DefaultView.ToTable
```

## Gembox to Color the Header and Autofit Columns
``` Vb.net 

Try
 SpreadsheetInfo.SetLicense(in_GemboxLicenseKey)
 
 Dim tempWorkBook As ExcelFile = ExcelFile.Load(in_ExcelFilePath)

Dim tempSheetName As ExcelWorksheet =tempWorkBook.Worksheets(in_ExcelSheetName)
Dim Range As CellRange
                
'Color the Header
Range = tempSheetName.Cells.GetSubrange("A1", "H1")
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

## Join multiple rows of single column with "," as separator

``` Vb.net
String.Join("," , DataTable.DefaultView.ToTable(True,"Request Type").Copy.AsEnumerable().Select(function(r) r.Item(0).ToString).ToList)
```
  
 
