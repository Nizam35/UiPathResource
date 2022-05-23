### Connect UiPath with diffrent Database instance

## Oracle Database

``` vb.net
Step 1 install Package = "Oracle.ManagedDataAccess.Core": "2.19.31"

Stpe 2 : Import "Oracle.ManagedDataAccess.Client"

Step 3 :  Make a DB connection
'User Id=Basamad;Password=b2asmad;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST =prodcsr-scan.bupame.com)(PORT = 1551)) (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.102.116)(PORT = 1551)) (LOAD_BALANCE = yes) (CONNECT_DATA = (SERVER = DEDICATED) (SERVICE_NAME = PROD.bupa.net)) (FAILOVER_MODE = (TYPE = SELECT) (METHOD = BASIC) (RETRIES = 180) (DELAY = 5)))

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

