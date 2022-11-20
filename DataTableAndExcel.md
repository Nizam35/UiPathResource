### Code Snippet Specific To Datatable


## Dynamic Linq to Get Row from one table against second table based on conditional columns

### Sample Conditonal columns (Transaction Item)

| Name Col         | Age Col | Empl ID Col | Description | 
|--------------|:-----:|-----------:|-----------:|
| True |  false |        false |  to get the rows from Current Table when only Name matches and rest of the columns are not mathing |
| false      |  false |   true | to get the rows from Current Table when both Age & Empl ID matches and Name of the column are not mathing |

### Smaple Linq Code

```Linq
(From row1 In   in_CurrentDt
From row2 In  in_PreviousDt

Let RegisterNumber =If(row1(ConditionDictList(0).Key).ToString.tolower.trim.equals(row2(ConditionDictList(0).Key).ToString.tolower.trim),"True","False")
Let TradeName=If(row1( ConditionDictList(1).Key).ToString.tolower.trim.equals(row2(ConditionDictList(1).Key).ToString.tolower.trim),"True","False")
Let  Strength=If(row1( ConditionDictList(2).Key).ToString.tolower.trim.equals(row2(ConditionDictList(2).Key).ToString.tolower.trim),"True","False")

Let boolRegisterNumber= RegisterNumber.ToString.equals(ConditionDictList(0).Value)
Let  boolTradeName=  TradeName.ToString.equals(ConditionDictList(1).Value)
Let boolStrength=  Strength.ToString.equals(ConditionDictList(2).Value)

Where  boolRegisterNumber And  boolTradeName And  boolStrength And boolStrengthUnit And boolPharmaceuticalForm And boolSize And boolSizeUnit And boolPackageSize  And boolPublicprice
Select row1).tolist


## Using DataView to Sort Datatable

``` Sort
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






