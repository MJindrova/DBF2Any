#INCLUDE "dbf2any.h"
#INCLUDE "dbf2any_API.h"

DEFINE CLASS _dbf2oracle AS _dbf2any && Upload DBF file to ORACLE

   * System properties
   Version="0.0.0.2"  && Version

   iDriver=7          && Drivers count
   DIMENSION aDriver(7)
   aDriver(1)="Oracle in OraClient12Home1"
   aDriver(2)="Oracle in OraClient12Home1_32bit"
   aDriver(3)="Oracle in OraOdac11g_home1"
   aDriver(4)="Oracle in OraClient10g_home1"
   aDriver(5)="Oracle in OraHome92"
   aDriver(6)="Oracle ODBC Driver"
   aDriver(7)="ORA"


   cServer=""
   cUser  =""
   cPWD   =""



   PROCEDURE CreateSQLCommand
      * 
      * _dbf2ORACLE::CreateSQLCommand
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV, m.lii
      LOCAL ARRAY m.laFields(1)

      STORE "" TO m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV
      * Read field list
      FOR m.lii=1 TO AFIELDS(m.laFields, m.loPar.cAlias)
          m.lcSQLC=m.lcSQLC+[ "]+m.laFields(m.lii, 1)+[" ]+;
                   IIF(m.laFields(m.lii, 2)='C', "CHAR("+LTRIM(STR(m.laFields(m.lii, 3), 11))+" CHAR)",;
                   IIF(m.laFields(m.lii, 2)='V', "VARCHAR2("+LTRIM(STR(m.laFields(m.lii, 3), 11))+" CHAR)",;
                   IIF(m.laFields(m.lii, 2)='Q', "RAW("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='M', "CLOB",;
                   IIF(m.laFields(m.lii, 2)='Y', "DECIMAL(20,4)",;
                   IIF(m.laFields(m.lii, 2)='F', "FLOAT",;
                   IIF(m.laFields(m.lii, 2)='B', "FLOAT",;
                   IIF(m.laFields(m.lii, 2)='I', "INTEGER",;
                   IIF(m.laFields(m.lii, 2)='N', "DECIMAL("+LTRIM(STR(m.laFields(m.lii, 3), 3))+IIF(m.laFields(m.lii, 4)=0,")",","+LTRIM(STR(m.laFields(m.lii, 4), 3))+")") ,;
                   IIF(m.laFields(m.lii, 2)$'D,T', "DATE",;
                   IIF(m.laFields(m.lii, 2)='W', "BLOB",;
                   IIF(m.laFields(m.lii, 2)='G', "BLOB", "decimal(1)"))))))))))))+" "+IIF(m.laFields(m.lii, 5),"NULL,","NOT NULL,")

          m.lcSQLI=m.lcSQLI+' "'+m.laFields(m.lii, 1)+'", '
          m.lcSQLIV=m.lcSQLIV+"?"+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+", "
          m.lcSQLIBV=m.lcSQLIBV+"?m.laBF(%ROW%)."+m.laFields(m.lii, 1)+", "
      NEXT

      This.cCreateSQL="CREATE TABLE "+m.loPar.cTable+" ("+LEFT(m.lcSQLC, LEN(m.lcSQLC)-1)+")"
      This.cInsertSQL="INSERT INTO "+m.loPar.cTable+" ("+LEFT(m.lcSQLI, LEN(m.lcSQLI)-2)+") VALUES ("+LEFT(m.lcSQLIV, LEN(m.lcSQLIV)-2)+")"
      This.cInsertBatchSQL="INSERT INTO "+m.loPar.cTable+" ("+LEFT(m.lcSQLI, LEN(m.lcSQLI)-2)+") VALUES ("+LEFT(m.lcSQLIBV, LEN(m.lcSQLIBV)-2)+")"
      
      This.cDropSQL=[DROP TABLE "]+m.loPar.cTable+["]
      This.cDeleteSQL=[DELETE FROM  "]+m.loPar.cTable+["]
   ENDPROC 


   PROCEDURE CheckTableIfExists
      * 
      * _dbf2ORACLE::CheckTableIfExists
      * 
      LPARAMETERS m.loPar
      
      SELE (m.loPar.cATables)
      LOCATE FOR UPPER(ALLT(TABLE_SCHEM))==UPPER(This.cUser) AND UPPER(ALLT(TABLE_NAME))==UPPER(m.loPar.cTable)
      RETURN FOUND()
   ENDPROC 


   PROCEDURE Upload2Remote
      * 
      * _dbf2ORACLE::Upload2Remote
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.liErr, m.lihdbc

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
      * _dbf2ORACLE::BeforeUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE AfterUpload2Remote
      * 
      * _dbf2ORACLE::AfterUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE GetConnectionString && Create connection string
      * 
      * _dbf2ORACLE::GetConnectionString
      * 
      LPARAMETERS m.lcDriver
      
      RETURN "Driver={"+m.lcDriver+"};DBQ="+This.cServer+";uid="+This.cUser+";pwd="+This.cPWD+";"+This.cExtendedCS
   ENDPROC

ENDDEFINE
