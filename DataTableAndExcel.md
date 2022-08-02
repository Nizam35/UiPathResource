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

## Insert Datatable to Exisitng or new Sheet in new or exisitng Workbook 

``` Vb.net
 Try
 
   SpreadsheetInfo.SetLicense(in_LicenseKey)
     Dim TempWorkBook As ExcelFile
	 
	If File.Exists(in_FilePath).Equals(False) Then
		TempWorkBook = New ExcelFile()
	Else
		TempWorkBook = ExcelFile.Load(in_FilePath)
	End If
   

Dim WorkSheet As ExcelWorksheet =   If(TempWorkBook.Worksheets.Contains(in_SheetName) , TempWorkBook.Worksheets(in_SheetName) ,TempWorkBook.Worksheets.Add(in_SheetName))

WorkSheet.InsertDataTable(in_DataTable,
            New InsertDataTableOptions() With
            {
                .ColumnHeaders = True,
                .StartRow = 0
            })
	TempWorkBook.Save(in_FilePath)
	
Catch ex As Exception
	Console.WriteLine(ex.ToString + " at soure " + ex.Source)
End Try

```


## Get Sheet Name using Sheet Index

```cs
try{
       	   var ExcelApp = new Microsoft.Office.Interop.Excel.Application();
           Microsoft.Office.Interop.Excel.Workbook workbook = ExcelApp.Workbooks.Open(@"C:\Users\Ahmed.Nizamuddin\Music\Membership_CorpDataValidationRPA\Required Files\Caesar Validation.xlsx");
            //Microsoft.Office.Interop.Excel.Worksheet sheet;
		//ExcelApp.Visible = true;	
		try{
				 var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[in_SheetIndex];
				out_SheetName = worksheet.Name;
				out_IsSuccess = true;
			}catch(Exception ex){
			out_IsSuccess = false;
		}
			ExcelApp.Visible = false;
            workbook.Save();
            workbook.Close(false);  
            ExcelApp.Quit();
			System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
      // excel.Quit();
	
}catch(Exception ex){
	Console.WriteLine("Excpetion "+ex.Message + "at source " + ex.Source);
}
```



## Join multiple rows of single column with "," as separator

``` Vb.net
String.Join("," , DataTable.DefaultView.ToTable(True,"Request Type").Copy.AsEnumerable().Select(function(r) r.Item(0).ToString).ToList)
```
  
 
