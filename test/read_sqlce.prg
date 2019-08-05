LOCAL m.lcPath

m.lcPath=SYS(16)
m.lcPath=IIF(RAT("\", m.lcPath)>0, LEFT(m.lcPath, RAT("\", m.lcPath)), m.lcPath)


loConnection = CREATEOBJECT('ADODB.Connection')
loConnection.Open("Provider=Microsoft.SQLSERVER.CE.OLEDB.3.5;Data Source="+m.lcPath+"..\out\test35.sdf;ssce:database password=;persist security info=false;")

m.loRS = CREATEOBJECT("ADODB.Recordset")
m.loRS.ActiveConnection = loConnection

m.loCA=CREATEOBJECT("CursorAdapter")
m.loCA.DataSourceType = "ADO"
m.loCA.Datasource = m.loRS
m.loCA.MapBinary = .T.
m.loCA.MapVarchar = .T.
m.loCA.Alias = SYS(2015)
m.loCA.SelectCmd = "SELECT * FROM INFORMATION_SCHEMA.TABLES"

=m.loCA.CursorFill(.F.)
m.loCA.CursorDetach()

SELECT (m.loCA.Alias)
SCAN ALL

     m.loRS2 = CREATEOBJECT("ADODB.Recordset")
     m.loRS2.ActiveConnection = loConnection
     m.loCA2=CREATEOBJECT("CursorAdapter")
     m.loCA2.DataSourceType = "ADO"
     m.loCA2.Datasource = m.loRS2
     m.loCA2.MapBinary = .T.
     m.loCA2.MapVarchar = .T.
     m.loCA2.Alias = SYS(2015)
     m.loCA2.SelectCmd = "SELECT * FROM "+ALLTRIM(TABLE_NAME)

     =m.loCA2.CursorFill(.F.)
     m.loCA2.CursorDetach()

     SELECT (m.loCA2.Alias)
     BROWSE NORMAL
     USE
     
     SELECT (m.loCA.Alias)
ENDSCAN
loConnection.Close()