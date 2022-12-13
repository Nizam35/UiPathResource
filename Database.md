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


