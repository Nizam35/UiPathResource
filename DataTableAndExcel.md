### Code Snippet Specific To Datatable

## Excel Sheet to Datatable using Gem Obx

``` vb.net
try {
  SpreadsheetInfo.SetLicense(in_GemboxLicenseKey);
   var workbook = ExcelFile.Load(in_FilePath);
	var worksheet = workbook.Worksheets[in_SheetName];
	
	int TotalColumns = worksheet.CalculateMaxUsedColumns();

    out_DataTable  = worksheet.CreateDataTable(new CreateDataTableOptions()
        {
            ColumnHeaders = in_HasHeader,
            StartRow = in_StartRow,
            NumberOfColumns = TotalColumns,
            NumberOfRows = worksheet.Rows.Count,
           // Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
        });

} catch(Exception ex){
	Console.WriteLine("Exception "+ex.Message + "at source " + ex.Source);
	//io_DataTable = null;
}
```

##  Generate to Datatable using Gembox
```csharp
try {
  SpreadsheetInfo.SetLicense(in_GemboxLicenseKey);
   var workbook = ExcelFile.Load(in_FilePath);
	var worksheet = ! String.IsNullOrEmpty(in_SheetName) ?workbook.Worksheets[in_SheetName]:workbook.Worksheets[in_SheetIndex];
	
	//Get Column index
	//Console.WriteLine("current Column Index is {0}", worksheet.Columns.Count);
	int TotalColumns = worksheet.CalculateMaxUsedColumns();
	//Console.WriteLine("Total Allocated cell is {0}", TotalColumns);
	//Total No. of Rows
	//	Console.WriteLine("Current last row index {0}", worksheet.Rows.Count);

    io_DataTable  = worksheet.CreateDataTable(new CreateDataTableOptions()
        {
			ColumnHeaders = in_HasHeader,
            StartRow = 0,
            NumberOfColumns = TotalColumns,
            NumberOfRows = worksheet.Rows.Count,
            Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
        });

	io_DataTable = io_DataTable.AsEnumerable().Where( row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field.ToString())) ).CopyToDataTable();
	
} catch(Exception ex){
	Console.WriteLine("Excpetion "+ex.Message + "at source " + ex.Source);
	//io_DataTable = null;
}
```

## Using DataView to Sort Datatable

```vb.net
Datatable.DefaultView.Sort = "ColName ASC,colName DESC"
DataTable = Datatable.DefaultView.ToTable
```

### Vb.net to Group the Datatable and get the  Sum,Count and percentage
```vb.net

Dim testDt As DataTable = New DataTable()
testDt.Columns.Add("Name", GetType(System.String))
testDt.Columns.Add("Age", GetType(system.Int32 ))
testDt.Columns.Add("Sum", GetType(system.Int32 ) )
testDt.Columns.Add("Count", GetType(system.Int32 ) )
testDt.Columns.Add("Percentage", GetType(system.Decimal ) )


testDt= (From dte In in_TestDt.AsEnumerable
	Group dte By col1=dte("Name").ToString.Trim Into Group
	Select testDt.Rows.Add(
				{col1, Group.Sum(Function (x) CInt(x("Age").toString.Trim)),Group.Count() ,(Group.Count()/ in_TestDt.RowCount())*100}
			       )
	).CopyToDataTable
				
						
out_FinalDt =testDt.Copy
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
  
## Get Excel Column Alphabet using Index 

```vb.net
 Try
 Dim dividend As Integer = in_ColumnIndex
 Dim columnName As String = String.Empty
 Dim modulo As Integer

  While dividend > 0
       modulo = (dividend - 1) Mod 26
       columnName = Convert.ToChar(65 + modulo).ToString() & columnName
       dividend = CInt((dividend - modulo) / 26)
   End While
   
   out_ColumnName = columnName
   
Catch ex As Exception
	Console.WriteLine("Exception "+ex.ToString)
End Try
```

## Convert Excel Sheet to PDF
```vb.net
Try 
		Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MTUzODIzQDMxMzcyZTMzMmUzMG1KckNZUjlKU1hpRFpCRTByK2I4TDZMRXlLamVYSzRTRDBzSFlCRzBwMVE9")
		Dim excelEngine As ExcelEngine = New ExcelEngine()
		Dim application As IApplication = excelEngine.Excel
 		application.DefaultVersion = ExcelVersion.Excel2016
  		
	  'Open Work Book
	   Dim workbook As IWorkbook = application.Workbooks.Open("C:\Users\Ahmed.Nizamuddin\Documents\Retention.xlsx", ExcelOpenType.Automatic)
		
	   Dim  sheet As IWorksheet =  workbook.Worksheets("Retention Model Scorecard")
	   	sheet.PageSetup.Orientation = ExcelPageOrientation.Landscape
	   'For Each sheet As IWorksheet In workbook.Worksheets
	   		
		'Next		
	 	
		'Initialize ExcelToPdfConverterSettings
			  Dim settings As ExcelToPdfConverterSettings = New ExcelToPdfConverterSettings()
			  'Set the gridlines display style as Invisible
 			 settings.DisplayGridLines = GridLinesDisplayStyle.Invisible
			 'Disable ExportDocumentProperties
  			settings.ExportDocumentProperties = False
			'Enable ExportQualityImage
  				settings.ExportQualityImage = True
				'Disable ShowHeader
  			settings.HeaderFooterOption.ShowHeader = False
			 'Disable ShowFooter
 			 settings.HeaderFooterOption.ShowFooter = False
				'Disable IsConvertBlankPage
 				 settings.IsConvertBlankPage = False	
				'Disable IsConvertBlankSheet
  				settings.IsConvertBlankSheet = False
				 'Set layout option as FitAllColumnsOnOnePage
 				 settings.LayoutOptions = LayoutOptions.FitAllColumnsOnOnePage
				 'Set layout option as FitAllRowsOnOnePage
  				'settings.LayoutOptions = LayoutOptions.FitAllRowsOnOnePage
				
				
			 'Open the Excel document to convert
  			Dim converter As ExcelToPdfConverter = New ExcelToPdfConverter(workbook)
 			 'Initialize the PDF document
 			Dim pdfDocument As New PdfDocument()

  			'Convert Excel document into PDF document
  			pdfDocument = converter.Convert(settings)
  			'Save the PDF file
  			pdfDocument.Save("test.pdf")
	

	Catch ex As Exception
			Throw New BusinessRuleException(ex.Message)
	End Try	


```
## Left Join

```vb.net
Try
Dim dt As New DataTable()
dt = out_FinalDatatable.Clone()

dt = ( From cognosRow In in_CognosDataTable.AsEnumerable()
  Group Join cchiRow In in_CCHIDataTable.AsEnumerable()
   On cognosRow.Item("CCHI Acc-Lic no").ToString Equals cchiRow.item("License No From CCHI").ToString Into Group
   Let matchedfirstRow = Group.FirstOrDefault()
    Select ra = { 
		      cognosRow("Provider Code"),
			   cognosRow("Provider Name"),
			    cognosRow("Provider Level"),
			    cognosRow("CCHI Acc-Lic no"),
				If(String.IsNullOrEmpty(cognosRow.Item("CCHI Eff To").ToString), Nothing, DateTime.Parse(cognosRow("CCHI Eff To").ToString).ToString("dd/MM/yyyy")),
				cognosRow("CCHI Status").ToString,
			   If(isNothing(matchedfirstRow), Nothing, matchedfirstRow("License No From CCHI")),
			    If(isNothing(matchedfirstRow), Nothing, DateTime.Parse(matchedfirstRow("Latest CCHI Eff To").ToString).ToString("dd/MM/yyyy")),
				If(isNothing(matchedfirstRow),"Not found",If(matchedfirstRow("Latest CCHI Eff To").ToString.Equals(cognosRow("CCHI Eff To").ToString),"True","False"))
				}
	Select dt.Rows.Add(ra)).CopyToDataTable
		
		'out_FinalDatatable = dt.Copy.Select("[Provider Level] <> 'Out of Network'").CopyToDataTable
	
		Console.WriteLine("Total Rows are " + dt.RowCount().ToString)
		out_FinalDatatable = dt.Copy
		
	Catch ex As Exception
		Console.WriteLine("exception is "+ ex.Message)
	End Try
```





