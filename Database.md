### Connect UiPath with diffrent Database instance

## Oracle Database

``` vb.net
Step 1 install Package = "Oracle.ManagedDataAccess.Core": "2.19.31"

Stpe 2 : Import "Oracle.ManagedDataAccess.Client"

Step 3 :  Make a DB connection


' Make Seperate connection
Dim inout_oralceCon as System.Data.Common.DbConnection

If (IsNothing(inout_oralceCon)) Then
	inout_oralceCon = New Oracle.ManagedDataAccess.Client.OracleConnection(in_OraConString)
	inout_oralceCon.Open()
End If

' Execute the Query

Dim objCon As OracleConnection =CType(in_OraConnection, OracleConnection)

Dim dt1 As DataTable
Dim ds As DataSet= New DataSet()

Dim cmdSqlAdapter As OracleDataAdapter= New OracleDataAdapter(sql_Query,objCon)

cmdSqlAdapter.Fill(ds)
 
dt1 = ds.Tables(0).Clone() ' clone Header Format

' updte Datatable
For Each row As DataRow In ds.Tables(0).Rows
    dt1.ImportRow(row)
Next

```
## Below SQL Queries used for Identifying the New Rows in one Table1 By Comparing another Table2

```Sql

SELECT *  FROM  Table1 AS c1
	WHERE C1.RegisterNumber NOT IN(SELECT RegisterNumber FROM dbo.Table2)
	AND c1.TradeName Not In (SELECT  TradeName FROM  dbo.Table2)
	AND C1.Strength  Not IN (SELECT    Strength FROM  dbo.Table2)
	AND c1.StrengthUnit  NOT IN (SELECT  StrengthUnit FROM  dbo.Table2)
	AND C1.PharmaceuticalForm  Not IN (SELECT    PharmaceuticalForm FROM  dbo.Table2)
	AND c1.Size  NOT IN (SELECT  Size FROM  dbo.Table2)
	AND C1.SizeUnit  Not IN (SELECT    SizeUnit FROM  dbo.Table2)
	AND c1.PackageSize Not In (SELECT  PackageSize FROM  dbo.Table2)
	AND C1.Publicprice  Not IN (SELECT    Publicprice FROM  dbo.Table2)
	
```

## Get the Latest Row from Table 2 After Successful Match for each Row in Table 1

``` sql
-- Consider Table 1 has 2 Rows, While matching with Table 2 we received 5 rows as Successful Match for 2 the Rows in Table1. Below Query will get only 2 rows from Table 2 by getting the latest rows out of 5 matched Rows.

Select * From 
(Select *, ROW_NUMBER() OVER ( PARTITION By Res.TradeName Order By  CONVERT(DATETIME,Res.DateOfChange) DESC )  AS row 
From ( Select T2.* From Table2 As T2  Where Exists (Select * From Table1 As T1  Where  T2.SFDACode <>  T1.SFDACode And  
T2.TradeName =  T1.TradeName And  T2.Strength <>  T1.Strength And  T2.Unitofstrength =  T1.Unitofstrength And  
T2.DosageForm =  T1.DosageForm And  T2.Size =  T1.Size And  T2.Unit =  T1.Unit And  T2.Packagesize =  T1.Packagesize And  T2.Price =  T1.Price) ) As Res) As Sub Where Sub.row =1
```

  


