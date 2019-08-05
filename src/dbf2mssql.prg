#INCLUDE "dbf2any.h"
#INCLUDE "dbf2any_API.h"

DEFINE CLASS _dbf2mssql AS _dbf2any && Upload DBF file to MSSQL

   * System properties
   Version="0.0.0.2"  && Version

   iDriver=7          && Drivers count
   DIMENSION aDriver(7)
   aDriver(1)="SQL Server Native Client 12.0"
   aDriver(2)="SQL Server Native Client 11.0"
   aDriver(3)="SQL Server Native Client 10.0"
   aDriver(4)="SQL Server Native Client"
   aDriver(5)="ODBC Driver 11 for SQL Server"
   aDriver(6)="ODBC Driver 13 for SQL Server"
   aDriver(7)="SQL Server"


   cServer=""
   cDB    =""
   cUser  =""
   cPWD   =""


   cGWDT="VARBINARY(MAX)"
   cMDT="VARCHAR(MAX)"


   PROCEDURE CreateSQLCommand
      * 
      * _dbf2mssql::CreateSQLCommand
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV, m.lii
      LOCAL ARRAY m.laFields(1)

      STORE "" TO m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV
      * Read field list
      FOR m.lii=1 TO AFIELDS(m.laFields, m.loPar.cAlias)
          m.lcSQLC=m.lcSQLC+[ "]+m.laFields(m.lii, 1)+[" ]+;
                   IIF(m.laFields(m.lii, 2)='C', "char("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='V', "varchar("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='Q', "varbinary("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='M', This.cMDT,;
                   IIF(m.laFields(m.lii, 2)='Y', "decimal(20,4)",;
                   IIF(m.laFields(m.lii, 2)='F', "float",;
                   IIF(m.laFields(m.lii, 2)='B', "float",;
                   IIF(m.laFields(m.lii, 2)='I', "integer",;
                   IIF(m.laFields(m.lii, 2)='N', "decimal("+LTRIM(STR(m.laFields(m.lii, 3), 3))+IIF(m.laFields(m.lii, 4)=0,")",","+LTRIM(STR(m.laFields(m.lii, 4), 3))+")") ,;
                   IIF(m.laFields(m.lii, 2)$'D,T', "datetime",;
                   IIF(m.laFields(m.lii, 2)='G', This.cGWDT,;
                   IIF(m.laFields(m.lii, 2)='W', This.cGWDT, "decimal(1)"))))))))))))+" "+IIF(m.laFields(m.lii, 5),"NULL,","NOT NULL,")

          m.lcSQLI=m.lcSQLI+" ["+m.laFields(m.lii, 1)+"], "
          m.lcSQLIV=m.lcSQLIV+IIF(NOT (INLIST(m.laFields(m.lii, 2),'G','W') AND This.cGWDT="VARBINARY(MAX)"), "?"+m.loPar.cAlias+"."+m.laFields(m.lii, 1),;
                                  'CAST(?'+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+' AS VARBINARY(MAX))')+", "
          m.lcSQLIBV=m.lcSQLIBV+"?m.laBF(%ROW%)."+m.laFields(m.lii, 1)+", "
      NEXT
      This.cCreateSQL="CREATE TABLE dbo."+m.loPar.cTable+" ("+LEFT(m.lcSQLC, LEN(m.lcSQLC)-1)+")"
      This.cInsertSQL="INSERT INTO dbo."+m.loPar.cTable+" ("+LEFT(m.lcSQLI, LEN(m.lcSQLI)-2)+") VALUES ("+LEFT(m.lcSQLIV, LEN(m.lcSQLIV)-2)+")"
      This.cInsertBatchSQL="INSERT INTO dbo."+m.loPar.cTable+" ("+LEFT(m.lcSQLI, LEN(m.lcSQLI)-2)+") VALUES ("+LEFT(m.lcSQLIBV, LEN(m.lcSQLIBV)-2)+")"
      
      This.cDropSQL=[DROP TABLE dbo."]+m.loPar.cTable+["]
      This.cDeleteSQL=[DELETE FROM  dbo."]+m.loPar.cTable+["]
   ENDPROC 


   PROCEDURE CheckTableIfExists
      * 
      * _dbf2mssql::CheckTableIfExists
      * 
      LPARAMETERS m.loPar
      
      SELE (m.loPar.cATables)
      LOCATE FOR UPPER(ALLT(TABLE_NAME))==UPPER(m.loPar.cTable)
      RETURN FOUND()
   ENDPROC 


   PROCEDURE Upload2Remote
      * 
      * _dbf2mssql::Upload2Remote
      * 
      LPARAMETERS m.loPar
       
      LOCAL m.liErr, m.lcSQL, m.lii, m.lihdbc
      LOCAL ARRAY m.laBF(2,1)

      m.liErr=__DBF2ANY_ErrOK
      m.lihdbc=This.hdbc

      SELE (m.loPar.cAlias)
      GO TOP

      =SQLSETPROP(m.lihdbc, "Transactions", 2)
      IF m.liErr=__DBF2ANY_ErrOK
         =SQLPREPARE(m.lihdbc, This.cInsertSQL)
         SCAN REST
              IF SQLEXEC(m.lihdbc, This.cInsertSQL)<=0
                 m.liErr=__DBF2ANY_ErrNotInsert
                 =AERROR(m.loPar.laErr)
                 EXIT
              ENDIF
         ENDSCAN
      ENDIF

      RETURN m.liErr
   ENDPROC 


   PROCEDURE BeforeUpload2Remote
      * 
      * _dbf2mssql::BeforeUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      =SQLEXEC(This.hdbc, "SET IDENTITY_INSERT dbo.["+m.loPar.cTable+"] ON")
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE AfterUpload2Remote
      * 
      * _dbf2mssql::AfterUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      =SQLEXEC(This.hdbc, "SET IDENTITY_INSERT dbo.["+m.loPar.cTable+"] OFF")
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE GetConnectionString
      * 
      * _dbf2MSSQL::GetConnectionString
      * 
      LPARAMETERS m.lcDriver
      
      IF UPPER(m.lcDriver)=="{SQL SERVER}"
         This.cGWDT="IMAGE"
         This.cMDT="TEXT"
      ENDIF
      RETURN "Driver={"+m.lcDriver+"};SERVER="+This.cServer+";Database="+This.cDB+";uid="+This.cUser+";pwd="+This.cPWD+";"+This.cExtendedCS

   ENDPROC

ENDDEFINE
