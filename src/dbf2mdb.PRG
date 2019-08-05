#INCLUDE "dbf2any.h"
#INCLUDE "dbf2any_API.h"


DEFINE CLASS _dbf2MDB AS _dbf2any && Upload DBF file to MS Access

   * System properties
   Version="0.0.0.2"  && Version

   iDriver=2          && Drivers count
   DIMENSION aDriver(2)
   aDriver(1)="Microsoft Access Driver (*.mdb, *.accdb)"
   aDriver(2)="Microsoft Access driver (*.mdb)"


   iFormat=1         && Format/version MDB file (1 - MS Access 2.0, 2 - MS Access 95/97, 3 - MS Access 2000, 4 MS Access 2007)
   lMDBReCreate=.F.  && 

   cTemplateMDB=""   && Template MDB for creating new MDB file
   cTargetMDB=""     && Target MDB

   PROCEDURE Init()
      * 
      * _dbf2MDB::Init
      * 
      
      DODEFAULT()
      IF EMPTY(This.cTemplateMDB)
         This.cTemplateMDB=This.cFolder+"template\empty.mdb"
      ENDIF
   ENDPROC


   PROCEDURE CreateSQLCommand
      * 
      * _dbf2MDB::CreateSQLCommand
      * 
      LPARAMETERS m.loPar
      LOCAL m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV, m.lii
      LOCAL ARRAY m.laFields(1)

      STORE "" TO m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV
      * Read field list
      FOR m.lii=1 TO AFIELDS(m.laFields, m.loPar.cAlias)
          IF m.laFields(m.lii, 2)="V" AND m.laFields(m.lii, 3)>255
             m.laFields(m.lii, 2)="M" 
          ENDIF
          IF m.laFields(m.lii, 2)="Q" AND m.laFields(m.lii, 3)>255
             m.laFields(m.lii, 2)="W" 
          ENDIF
          
          m.lcSQLC=m.lcSQLC+[ "]+m.laFields(m.lii, 1)+[" ]+;
                   IIF(m.laFields(m.lii, 2)='C', "TEXT("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='V', "TEXT("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='Q', "VARBINARY("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='M', "LONGTEXT",;
                   IIF(m.laFields(m.lii, 2)='Y', "CURRENCY",;
                   IIF(m.laFields(m.lii, 2)='F', "FLOAT",;
                   IIF(m.laFields(m.lii, 2)='B', "DOUBLE",;
                   IIF(m.laFields(m.lii, 2)='I', "INTEGER",;
                   IIF(m.laFields(m.lii, 2)='N', "SINGLE" ,;
                   IIF(m.laFields(m.lii, 2)$'D,T', "DATETIME",;
                   IIF(m.laFields(m.lii, 2)='W', "LONGBINARY",;
                   IIF(m.laFields(m.lii, 2)='G', "LONGBINARY", "BIT"))))))))))))+" "+IIF(m.laFields(m.lii, 5),"NULL,","NOT NULL,")

          m.lcSQLI=m.lcSQLI+' ['+m.laFields(m.lii, 1)+'], '
          IF This.iSelectedDriver=1 AND INLIST(m.laFields(m.lii, 2), "Y", "F", "B", "I", "N")
             DO CASE
                CASE m.laFields(m.lii, 2)="Y"
                     m.lcSQLIV=m.lcSQLIV+"?CHRTRAN(STR("+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+",24,4),'.',','), "
                CASE m.laFields(m.lii, 2)="B"
                     m.lcSQLIV=m.lcSQLIV+"?CHRTRAN(STR("+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+",24,"+LTRIM(STR(m.laFields(m.lii, 4), 11))+"),'.',','), "
                CASE m.laFields(m.lii, 2)="I"
                     m.lcSQLIV=m.lcSQLIV+"?CAST("+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+" AS I), "
                CASE INLIST(m.laFields(m.lii, 2), "N", 'F')
                     m.lcSQLIV=m.lcSQLIV+"?CHRTRAN(STR("+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+","+LTRIM(STR(m.laFields(m.lii, 3), 11))+","+LTRIM(STR(m.laFields(m.lii, 4), 11))+"),'.',','), "
             ENDCASE
          ELSE
             m.lcSQLIV=m.lcSQLIV+"?"+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+", "
          ENDIF
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
      * _dbf2MDB::CheckTableIfExists
      * 
      LPARAMETERS m.loPar
      
      SELE (m.loPar.cATables)
      LOCATE FOR UPPER(ALLT(TABLE_NAME))==UPPER(m.loPar.cTable)
      RETURN FOUND()
   ENDPROC 


   PROCEDURE Upload2Remote
      * 
      * _dbf2MDB::Upload2Remote
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.liErr, m.lcSQL, m.lii, m.lihdbc
      LOCAL ARRAY m.laBF(5,1)

      m.liErr=__DBF2ANY_ErrOK
      m.lihdbc=This.hdbc

      SELE (m.loPar.cAlias)
      GO TOP

      =SQLSETPROP(m.lihdbc,"Transactions", 2)
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
      * _dbf2MDB::BeforeUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE AfterUpload2Remote
      * 
      * _dbf2MDB::AfterUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      This.lMDBReCreate=.F. && reset the flag
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE OpenRemoteDB()
      * 
      * _dbf2MDB::OpenRemoteDB
      * 
      LOCAL m.liErr,m.lii

      * Get out if template mdb not defined or not exists
      IF !FILE(This.cTargetMDB) AND (EMPTY(This.cTemplateMDB) OR !FILE(This.cTemplateMDB))
         RETURN __DBF2ANY_ErrFewParameters
      ENDIF

      IF This.iSelectedDriver=0
         This.CheckDrivers()
      ENDIF

      * detekce ODBC driveru dle souboru
      IF This.iSelectedDriver=0
         m.lii=RAT(".", This.cTargetMDB)
         This.iSelectedDriver=IIF(UPPER(SUBSTR(This.cTargetMDB, m.lii))==".MDB", 2, 1)
      ENDIF

      IF !FILE(This.cTargetMDB) OR This.lMDBReCreate
         m.liErr=This.CreateRemoteDB(This.cTargetMDB, This.cTemplateMDB)
         IF m.liErr#__DBF2ANY_ErrOK
            RETURN m.liErr
         ENDIF
      ENDIF

      RETURN DODEFAULT()
   ENDPROC


   PROCEDURE GetConnectionString
      * 
      * _dbf2MDB::GetConnectionString
      * 
      LPARAMETERS m.lcDriver
      
      RETURN "Driver={"+This.aDriver(This.iSelectedDriver)+"};DBQ="+This.cTargetMDB+";"+This.cExtendedCS
   ENDPROC


   PROCEDURE CreateRemoteDB && Create new MDB file
      * 
      * _dbf2MDB::CreateRemoteDB
      * 
      LPARAMETERS m.lcMDB, m.lcVMDB

      * lcMDB  - Target MDB file
      * lcVMDB - Template MDB file
      
      LOCAL m.lcMDB, m.lcVMDB, m.lcSQL, m.lihdbc, m.liODBChdbc, m.lii, m.lcMDBVersion, m.lcCRKey

      * How is version of ODBC driver for MS Access?
      * ( because key words are different for various version ODBC drivers)
      m.lihdbc=SQLSTRINGCONNECT("driver="+This.aDriver(This.iSelectedDriver)+";dbq="+m.lcVMDB)
      IF m.lihdbc=0
         RETURN __DBF2ANY_ErrConnectionFailed
      ENDIF
      m.liODBChdbc=SQLGETPROP(m.lihdbc, "ODBChdbc") &&
      m.lcVersion=SPACE(254) &&
      m.lii=0
      =DFB2ANY_SQLGetInfo(m.liODBChdbc, SQL_DRIVER_VER, @m.lcVersion, LEN(m.lcVersion), @m.lii) && Naèti verzi driveru
      =SQLDISCONNECT(m.lihdbc)

      * Create key word
      m.lcVersion=LTRIM(STR(VAL(LEFT(m.lcVersion, AT(".", m.lcVersion)-1)), 3))
      m.lcCRKey = IIF(m.lcVersion="3",IIF(This.iFormat=1, "CREATE_V2DB", "CREATE_DB"),;
                  IIF(This.iFormat=1, "CREATE_DBV2", IIF(This.iFormat=2, "CREATE_DBV3", IIF(This.iFormat=3, "CREATE_DBV4", "CREATE_DBV12"))))

      IF FILE(m.lcMDB) && If MDB file exist
         This.DeleteFile(m.lcMDB) && delete it
      ENDIF

      m.lcDriver=This.aDriver(This.iSelectedDriver)
      * Create new mdb
      IF DFB2ANY_SQLConfigDataSource(0, 1, m.lcDriver, m.lcCRKey+"="+m.lcMDB)#1
         * Creating MDB file - fail
         RETURN __DBF2ANY_ErrDBNotCreated
      ENDIF
      RETURN __DBF2ANY_ErrOK
   ENDPROC

ENDDEFINE