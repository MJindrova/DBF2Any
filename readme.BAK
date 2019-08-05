# DBF to Any export

You can export some DBF or cursor to Microsoft SQL Server, Oracle, Microsoft Access, MySQL or Microsoft SQL Server Compact Edition. 

## VFP Compatibility
VFP 6 SP3, VFP 7, VFP 8, VFP 9, VFP Advanced, VFP Advanced 64 bit

## Limitations
- Export to Oracle doesn't support empty general field.
- Export general field to Microsoft SQL Server Compact Edition need VFP 9, VFP Advanced or VFP Advanced 64 bit.


## Files
src\dbf2any.h - header file
src\dbf2any_API.h - header file for API functions
src\dbf2any.prg - base class
src\dbf2mdb.prg - class for export to Microsoft Access
src\dbf2mssql.prg - class for export to Microsoft SQL Server
src\dbf2mysql.prg - class for export to MySQL/MariaDB
src\dbf2oracle.prg - class for export to Oracle
src\dbf2sqlce.prg - class for export to Microsoft SQL Server Compact Edition
src\template\empty.mdb  - empty mdb file for checking ODBC driver version

test\test.dbf/fpt - table for testing
test\test.rtf     - rtf file or testing
test\test_mdb.prg - test for export to Microsoft Access
test\test_mssql.prg - test for export to Microsoft SQL Server
test\test_mysql.prg - test for export to MySQL/MariaDB
test\test_oracle.prg - test for export to Oracle
test\test_sqlce.prg - test for export to Microsoft SQL Server Compact Edition


## Examples
## Examples for Microsoft SQL Server
```foxpro
#INCLUDE "..\src\dbf2any.h"

LOCAL m.lcPath, m.loDBF2MSSQL, m.lcAlias, m.lii

m.lcPath=SYS(16)
m.lcPath=IIF(RAT("\", m.lcPath)>0, LEFT(m.lcPath, RAT("\", m.lcPath)), m.lcPath)

m.lcAlias=SYS(2015)
#IF VAL(SUBS(VERSION(),LEN("Visual FoxPro ")+1,2))<9
   CREATE CURSOR (m.lcAlias) (C_CHAR C(10) NULL, C_MEMO M NULL, C_BOOL L NULL, C_DATE D NULL, C_DT T NULL, ;
                              C_NUM N (10,5) NULL, C_FLOAT F(10,5) NULL, C_INT I NULL, C_DOUBLE B NULL, C_CURRENCY Y NULL)

   FOR m.lii=1 TO 10
       INSERT INTO (m.lcAlias) (C_CHAR, C_MEMO, C_BOOL, C_DATE, C_DT, C_NUM, C_FLOAT, C_INT, C_DOUBLE, C_CURRENCY);
           VALUES (REPLICATE(CHR(m.lii+64), 10), REPLICATE(CHR(m.lii+64), 100), ;
                   m.lii%2=0, DATE()+m.lii, DATETIME()+m.lii, m.lii*1.6, m.lii*1.6, m.lii, m.lii*1.6, m.lii*1.6)
   NEXT
#ELSE
   CREATE CURSOR (m.lcAlias) (C_CHAR C(10) NULL, C_MEMO M NULL, C_BOOL L NULL, C_DATE D NULL, C_DT T NULL, ;
                              C_NUM N (10,5) NULL, C_FLOAT F(10,5) NULL, C_INT I NULL, C_DOUBLE B NULL, C_CURRENCY Y NULL,;
                              C_VARCHAR V(10) NULL, C_VARBINARY Q(10) NULL, C_BLOB W NULL)

   FOR m.lii=1 TO 10
       INSERT INTO (m.lcAlias) (C_CHAR, C_MEMO, C_BOOL, C_DATE, C_DT, C_NUM, C_FLOAT, C_INT, C_DOUBLE, C_CURRENCY, C_VARCHAR, C_VARBINARY, C_BLOB);
           VALUES (REPLICATE(CHR(m.lii+64), 10), REPLICATE(CHR(m.lii+64), 100), ;
                   m.lii%2=0, DATE()+m.lii, DATETIME()+m.lii, m.lii*1.6, m.lii*1.6, m.lii, m.lii*1.6, m.lii*1.6,;
                   REPLICATE(CHR(m.lii+64), 5), CHR(m.lii+50), FILETOSTR(m.lcPath+"\test.rtf"))
   NEXT
   GO BOTTOM
   REPLACE C_BLOB WITH .NULL.
   SKIP -1
   REPLACE C_BLOB WITH 0h
#ENDIF

DO (m.lcPath+"..\src\dbf2any") WITH "create", "mssql", m.loDBF2MSSQL
IF ISNULL(m.loDBF2MSSQL)
   RETURN
ENDIF

m.loDBF2MSSQL.cServer="localhost\SQLEXPRESS"
m.loDBF2MSSQL.cDB    ="test"
m.loDBF2MSSQL.cUser  ="test"
m.loDBF2MSSQL.cPWD   ="test"

m.loDBF2MSSQL.lReFill=.T.                             && If table in destination exists, delete data
m.loDBF2MSSQL.AttachDataSession(1)                    && Attach default data session
?m.loDBF2MSSQL.UploadDBF("", m.lcAlias, "")           && Upload cursor
?m.loDBF2MSSQL.UploadDBF(m.lcPath+"test.dbf", "", "") && Upload table
*?m.loDBF2MSSQL.UploadDBC(m.lcPath+"test.dbc", .T.)    && Upload all tables in DBC

DO (m.lcPath+"..\src\dbf2any") WITH "release", "mssql"

USE IN (m.lcAlias)
```


## General features
```foxpro
#INCLUDE "..\src\dbf2any.h"

LOCAL m.lcPath, m.loDBF2MSSQL, m.lcAlias, m.lii

m.lcPath=SYS(16)
m.lcPath=IIF(RAT("\", m.lcPath)>0, LEFT(m.lcPath, RAT("\", m.lcPath)), m.lcPath)
DO (m.lcPath+"..\src\dbf2any") WITH "create", "mssql", m.loDBF2MSSQL
IF ISNULL(m.loDBF2MSSQL)
   RETURN
ENDIF

* select driver (.T. - driver selected, .f. - driver not selected)
?m.loDBF2MSSQL.SelectDriver("ODBC Driver 11 for SQL Server")

m.loDBF2MSSQL.DeleteDriver(6)                               && use index
m.loDBF2MSSQL.DeleteDriver("ODBC Driver 11 for SQL Server") && use name
m.loDBF2MSSQL.DeleteDriver()                                && all drivers

m.loDBF2MSSQL.AddDriver("ODBC Driver 11 for SQL Server")    && add new driver

* list of all installed drivers
FOR m.lii=1 TO ALEN(m.loDBF2MSSQL.aExistsDrivers, 1)
    ?m.loDBF2MSSQL.aExistsDrivers(m.lii, 1)+' '+m.loDBF2MSSQL.aExistsDrivers(m.lii, 2)
NEXT


DO (m.lcPath+"..\src\dbf2any") WITH "release", "mssql"
```
