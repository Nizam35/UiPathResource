### Code Snippet Specific To Datatable

## Code to Add Missing Row based on Mapping Table
```vb.net
Try
    Dim gradeGroups As Integer()() = {
        New Integer() {2, 3, 4},
        New Integer() {5, 6, 7},
        New Integer() {10, 11}
    }

    Dim groupedData = InputDt.AsEnumerable().
        GroupBy(Function(row) row("Role")).
        Select(Function(Group) New With {
            .Role = Group.Key,
            .Rows = Group.ToList()
        }).ToList()

    FinalDataTable = InputDt.Clone()

    Console.WriteLine($"Input {InputDt.Rows.Count}")
    Console.WriteLine($"Total Groups Created {groupedData.Count}")

    For Each Group In groupedData
        Console.WriteLine($"Looping {Group.Role} Group")
        Dim groupTableDt As DataTable = Group.Rows.CopyToDataTable()

        Dim gradesInGroup = groupTableDt.AsEnumerable().
            Where(Function(row) Not row.IsNull("Grade")).
            Select(Function(row) row.Field(Of Integer)("Grade")).Distinct().ToList()

        Console.WriteLine($"Available grades in Input are {String.Join(",", gradesInGroup)}")

        For Each gradeGroup In gradeGroups
            Console.WriteLine($"Validate Group with {String.Join(",", gradeGroup)}")
            If gradeGroup.Any(Function(grade) gradesInGroup.Contains(grade)) Then
                Console.WriteLine("Fetch Temp Row")
                Dim TempRow As DataRow = groupTableDt.AsEnumerable().
                    Where(Function(row) Not row.IsNull("Grade") AndAlso gradeGroup.Contains(row.Field(Of Integer)("Grade"))).
                    FirstOrDefault()

                If TempRow IsNot Nothing Then
                    Console.WriteLine($"Temp Row Grade is {TempRow.Field(Of Integer)("Grade")}")

                    For Each grade In gradeGroup
                        If Not gradesInGroup.Contains(grade) Then
                            Console.WriteLine($"Adding The Grade {grade}")
                            Dim MapRow = MapDt.Select("[New Level] =" & grade).FirstOrDefault()

                            If MapRow IsNot Nothing Then
                                Dim newRow As DataRow = groupTableDt.NewRow()
                                newRow.ItemArray = TempRow.ItemArray

                                newRow("Profile Name (Position/Role Title)") = $"{MapRow("New Title")} - {newRow("Role")}"
                                newRow("Grade") = grade

                                groupTableDt.Rows.Add(newRow)
                                Console.WriteLine($"Group Rows After adding new Row {groupTableDt.Rows.Count}")
                            End If
                        End If
                    Next
                End If
            End If
        Next

        Console.WriteLine($"Group Rows After {groupTableDt.Rows.Count}")
        FinalDataTable.Merge(groupTableDt)
    Next

Catch ex As Exception
    Console.WriteLine($"{ex.Message} at source {ex.Source}")
End Try
```

## Convert HTML to Dataset using vb.net

```vb.net
Try
	'Initailize a html document
	Dim doc As New HtmlDocument()
	'Initailize the DataSet
	Dim Ds As New DataSet
	'load the html
	doc.LoadHtml(html)
	'load all the tables
	Dim tables As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//table")

For Each table As HtmlNode In tables
	'Initailize the Dt
	Dim Dt As New DataTable()
	'Load the tr tags of table
	Dim rows As HtmlNodeCollection = table.SelectNodes(".//tr")
	If rows IsNot Nothing Then
		Dim maxTdCount As Integer = 0
		'Check the maximum  td tags exists to create a columns
			For Each trNode  In rows
			' Count the number of <td> elements within the <tr>
					Dim tdCount As Integer = trNode.SelectNodes(".//td").Count
					' Update the maxTdCount if necessary
					If tdCount > maxTdCount Then
						maxTdCount = tdCount
					End If
			Next
    		'Add Columns
        	For Each Value In  Enumerable.Range(0, maxTdCount)
        			Dt.Columns.Add("Column"+Value.ToString,  Type.GetType("System.String"))
        	
        	Next		
	
	
        For Each row In rows
            'Initailize datarow and Index
        	Dim Index As Int32 = 0
        	Dim dr As DataRow = Dt.NewRow()
        	'Load all the td tags of trNode
        	Dim cells As HtmlNodeCollection = row.SelectNodes("td")
        	If cells IsNot Nothing Then
            	For Each cell In cells
                    'If inner text has garbage values  replace it with respect to the values.
	            Dim cellvalue = cell.InnerText.Replace("�","").Replace("&nbsp;","").Replace("&amp;","&")
                    dr("Column"+ index.ToString) = cellvalue
            	    Index = Index+1
              	Next
        	    Dt.rows.Add(dr)
            End If
        Next
        'Add the dt to Dataset
        If  Dt IsNot Nothing  andalso Dt.rows.Count >0 Then
        	Ds.Tables.Add(Dt)
        	Console.WriteLine("Added to Dataset")
        End If
    End If
Next
'Pass the Dataset to outside
io_Ds = Ds
Catch exp As Exception
	Console.WriteLine(exp.Message)
End Try

```

## Create DataSet for Each Grouped Rows frm Datatable
``` vb.net
' Grouping and filtering rows using LINQ
Dim groupedData = From row In DT_WorkItems.AsEnumerable()
	Let email = row.Field(Of String)("Agent Email")
	Where String.IsNullOrEmpty(email).Equals(False)
	Group row By email Into Group
	Select New With {
		.Email = email,
		.Rows = Group.CopyToDataTable()
		}							}
'Creating a dataset and adding tables for each group
Dim dataSet As New DataSet()	
		For Each Group In groupedData
			Dim datatable = Group.Rows
			datatable.TableName= Group.Email '  Make it Agent Name
			dataSet.Tables.Add(datatable)
		Next
out_AgentDataSet =dataSet.Copy
```


## Generate Datatable with all the files in a folder

```vb.net

Try

Dim testDt As Datatable = New DataTable()

testDt.Columns.Add("FilePath",Type.GetType("System.String")) ' Adding a new Column of Type string

Dim FilesCollection As New List(Of String)

FilesCollection = Directory.GetFiles("Data\ExampleDocuments").ToList()

FilesCollection.ForEach( Function(rowItem) testDt.Rows.Add({rowItem}) )

Catch ex As Exception
	Console.WriteLine(ex.ToString)
End Try
```


## Dynamic Linq to Get Row from one table against second table based on conditional columns

### Sample Conditonal columns (Transaction Item)

| Name   | Age  | Gender  | Description | 
|--------------|:-----:|-----------:|-----------:|
| True |  false |        false |  to get the rows from Current Table when only Name matches and rest of the columns are not mathing |
| false      |  false |   true | to get the rows from Current Table when both Age & Empl ID matches and Name of the column are not mathing |

### Sample Linq Code for converting the DataRow into Dictionary

```Linq
ConditionDictList = in_DataRow.Table.Columns.Cast<DataColumn>().ToDictionary((c) => c.ColumnName, c => in_DataRow[c])
```

```Vb.net
(From row1 In   in_CurrentDt ' We will be fetching the Row form this Datatable
From row2 In  in_PreviousDt '  Against this Datatable

Let  Name = If(row1(ConditionDictList(0).Key).ToString.tolower.trim.equals(row2(ConditionDictList(0).Key).ToString.tolower.trim),"True","False")
Let Age  =If(row1( ConditionDictList(1).Key).ToString.tolower.trim.equals(row2(ConditionDictList(1).Key).ToString.tolower.trim),"True","False")
Let  Gender  =If(row1( ConditionDictList(2).Key).ToString.tolower.trim.equals(row2(ConditionDictList(2).Key).ToString.tolower.trim),"True","False")

Let boolNamer= Name.ToString.equals(ConditionDictList(0).Value)
Let  boolAge =  Age.ToString.equals(ConditionDictList(1).Value)
Let boolGender=  Gender.ToString.equals(ConditionDictList(2).Value)

Where  boolRegisterNumber And  boolTradeName And  boolStrength
Select row1).tolist
```

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






