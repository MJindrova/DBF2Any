LPARAMETERS m.lcCommand, m.lcDST, m.loEngine

#INCLUDE "dbf2any.h"
#INCLUDE "dbf2any_API.h"


LOCAL m.lcPath, m.lcONERROR
m.loEngine=.NULL.
m.lcCommand=UPPER(ALLTRIM(m.lcCommand))

m.lcPath=SYS(16)
m.lcPath=SUBST(m.lcPath, AT(" ", m.lcPath, 2)+1)
m.lcPath=IIF(RAT("\", m.lcPath)>0, LEFT(m.lcPath, RAT("\", m.lcPath)), m.lcPath) 
m.lcDST=UPPER("dbf2"+m.lcDST)

IF m.lcCommand=="CREATE"
   * check if file exists
   IF FILE(m.lcPath+m.lcDST+".FXP") OR FILE(m.lcPath+m.lcDST+".PRG")
      * add procedural file
      IF FILE(m.lcPath+m.lcDST+".FXP") AND ATC(m.lcDST+".FXP,", UPPER(SET("PROCEDURE"))+",")=0
         SET PROCEDURE TO (m.lcPath+m.lcDST+".FXP") ADDITIVE
      ENDIF
      IF FILE(m.lcPath+m.lcDST+".PRG") AND ATC(m.lcDST+".FXP,", UPPER(SET("PROCEDURE"))+",")=0
         SET PROCEDURE TO (m.lcPath+m.lcDST+".PRG") ADDITIVE
      ENDIF
      m.lcONERROR=ON("ERROR")
      ON ERROR m.loEngine=.NULL.
      m.loEngine=CREATEOBJECT("_"+m.lcDST) && create objects
      ON ERROR &lcONERROR.
   ENDIF
ENDIF

IF m.lcCommand=="RELEASE"
   * check if file exists
   IF FILE(m.lcPath+m.lcDST+".FXP")
      IF ATC(m.lcDST+".FXP,", UPPER(SET("PROCEDURE"))+",")>0
         RELEASE PROCEDURE (m.lcPath+m.lcDST+".FXP")
      ENDIF
   ENDIF
ENDIF
   
RETURN .T.



DEFINE CLASS _dbf2any_Upload2Remote AS CUSTOM
   Name="_dbf2any_Upload2Remote"
   cATables=""
   cAlias=""
   cTable=""
   iErr=0
   cDBF=""
   DIMENSION laErr(1)

   PROCEDURE Init
      * 
      * _dbf2any_Upload2Remote::Init
      * 
      LPARAMETERS m.lcATables, m.lcAlias, m.lcTable
      
      This.cATables=m.lcATables
      This.cAlias=m.lcAlias
      This.cTable=m.lcTable
      This.cDBF=DBF(m.lcAlias)
   ENDPROC
ENDDEF


DEFINE CLASS _dbf2any AS Session && Upload DBF file to ODBC source
   HIDDEN DataSessionID, _DataSessionID, hdbc, Version, Catch
   DIMENSION aErrors(1), aDriver(1), aExistsDrivers(1)

   * System properties
   Version="0.0.0.2"  && Version
   Catch=.F.          && Mode Error catching
   nErr=0             && Error number
   cErrMSG=""         && Message error
   _DataSessionID = 0 && DataSession ID
   hdbc=0             && VFP handle of ODBC connection
   iDriver=0          && Drivers count
   iSelectedDriver=0  && Selected driver
   
   nVFPVersion=VFP_VERSION

   cFolder=""         && 
   nRemoteMode=__DBF2ANY_Mode_ODBC
   oConnection=.NULL. && OLE DB connection Object

   * User properties
   lReCreate=.T.      && If .T. and table exist in dbf file, a table will be delete
   lReFill=.F.        && If .T., old data will be delete

   cDropSQL=""
   cDeleteSQL=""
   cCreateSQL=""
   cInsertSQL=""
   cInsertBatchSQL=""
   
   cExtendedCS=""

   PROCEDURE Init()
      * 
      * _dbf2any::Init
      * 
      SET EXACT OFF
      SET TALK OFF
      SET DELE ON
      SET SAFETY OFF
      SET DECI TO 2
      SET POINT TO "."

      This._DataSessionID=This.DataSessionID

      LOCAL m.lcPath 
      m.lcPath=SYS(16)
      m.lcPath=SUBST(m.lcPath, AT(" ", m.lcPath, 2)+1)
      This.cFolder=IIF(RAT("\", m.lcPath)>0, LEFT(m.lcPath, RAT("\", m.lcPath)), m.lcPath) 

      DECLARE INTEGER ExpandEnvironmentStrings IN KERNEL32.DLL AS DBF2ANY_ExpandEnvironmentStrings STRING @ ,STRING @, INTEGER

      DECLARE INTEGER RegQueryValueEx IN ADVAPI32.DLL AS DBF2ANY_RegQueryValueEx LONG, STRING @, INTEGER, INTEGER @, STRING @, INTEGER @
      DECLARE INTEGER RegOpenKeyEx    IN ADVAPI32.DLL AS DBF2ANY_RegOpenKeyEx    LONG, STRING @, INTEGER, INTEGER, LONG @                     
      DECLARE INTEGER RegCloseKey     IN ADVAPI32.DLL AS DBF2ANY_RegCloseKey     LONG                                                           
      DECLARE INTEGER RegEnumKeyExA   IN ADVAPI32.DLL AS DBF2ANY_RegEnumKeyExA   LONG, INTEGER, STRING @, INTEGER @, INTEGER, STRING @, INTEGER @, STRING @

      DECLARE INTEGER SQLDrivers IN ODBC32.DLL AS DBF2ANY_SQLDrivers LONG henv, SHORT fdirection, STRING @szDSN, INTEGER cbDSNMax, ;
                                                                     INTEGER @pcbDSN, STRING @szDescription, INTEGER cbDescriptionMax, INTEGER @pcbDescriptionn
      DECLARE INTEGER SQLConfigDataSource IN odbccp32.dll AS DFB2ANY_SQLConfigDataSource LONG, INTEGER, STRING, STRING
      DECLARE INTEGER SQLGetInfo IN odbc32.dll AS DFB2ANY_SQLGetInfo LONG, INTEGER, STRING @, INTEGER, INTEGER @

      This.ReadInstalledDrivers()
   ENDPROC


   PROCEDURE DeleteFile
      * 
      * _dbf2any::DeleteFile
      * 
      LPARAMETERS m.lcFile
      
      LOCAL m.lcSAFETY
      m.lcSAFETY=SET("SAFETY")
      SET SAFETY OFF
      DELETE FILE (m.lcFile)
      SET SAFETY &lcSAFETY.
   ENDPROC


   PROCEDURE AddDriver
      * 
      * _dbf2any::AddDriver
      * 
      LPARAMETERS m.lcDriver
      
      This.iDriver=This.iDriver+1
      DIMENSION This.aDriver(This.iDriver)
      This.aDriver(This.iDriver)=m.lcDriver
   ENDPROC 


   PROCEDURE DeleteDriver
      * 
      * _dbf2any::DeleteDriver
      * 
      LPARAMETERS m.luDriver
      
      DO CASE 
         CASE EMPTY(m.luDriver)
              STORE 0 TO This.iDriver, This.iSelectedDriver
              DIMENSION This.aDriver(1)
              STORE "" TO This.aDriver
              RETURN
      
         CASE TYPE("m.luDriver")="N"
              IF m.luDriver>0 AND m.luDriver<=This.iDriver
                 =ADEL(This.aDriver, m.luDriver)
                 This.iDriver=This.iDriver-1
                 IF This.iSelectedDriver=m.luDriver
                    This.iSelectedDriver=0
                 ENDIF
                 DIMENSION This.aDriver(This.iDriver)
              ENDIF
              RETURN

         CASE TYPE("m.luDriver")="C"
              m.luDriver=UPPER(m.luDriver)
              LOCAL m.lii
              FOR m.lii=1 TO This.iDriver
                  IF UPPER(This.aDriver(m.lii))==m.luDriver
                     =ADEL(This.aDriver, m.lii)
                     This.iDriver=This.iDriver-1
                     DIMENSION This.aDriver(This.iDriver)
                     IF This.iSelectedDriver=m.lii
                        This.iSelectedDriver=0
                     ENDIF
                     EXIT
                  ENDIF
              NEXT

      ENDCASE

   ENDPROC 


   PROCEDURE SelectDriver
      * 
      * _dbf2any::SelectDriver
      * 
      LPARAMETERS m.luDriver

      IF EMPTY(m.luDriver)
         RETURN .F.
      ENDIF
      This.CheckDrivers(m.luDriver)
      RETURN This.iSelectedDriver>0
   ENDPROC 



   PROCEDURE ReadInstalledODBCDrivers
      * 
      * _dbf2any::ReadInstalledODBCDrivers
      * 

      LOCAL m.hEnv, m.cbDsnMax, m.pcbDSN, m.cbDescriptMax, m.pcbDescript, m.fDirection, m.retVal, m.szDsn, m.szDescript, m.lii,
            m.liRoot, m.lcBaseKey, m.lcKey, m.liHKEY, m.lcName, m.lcData, m.liData, m.liType, m.lcOUT, m.lcIN, m.liy
      LOCAL ARRAY m.laVers(1)            
            
      m.hEnv = VAL(SYS(3053))            && The hEnv (converted to a number)
      m.cbDsnMax = MAX_STRING            && How long can the DSN Name be?
      m.pcbDSN = 0                       && How many were actually returned
      m.cbDescriptMax = MAX_STRING       && How long can the description be?
      m.pcbDescript = 0                  && How long the description actually was
      m.fDirection = SQL_FETCH_FIRST     && The first time start at the top
      m.retVal = SQL_SUCCESS             && Start with no errors
      m.lii=1

      DO CASE
         CASE LEN(GETENV("ProgramFiles(x86)"))=0 && 32 bit OS
              m.liRoot=HKEY_LOCAL_MACHINE
              m.lcBaseKey="SOFTWARE\ODBC\ODBCINST.INI"
         
         CASE LEN(GETENV("ProgramFiles(x86)"))>0 AND GETENV("PROCESSOR_ARCHITECTURE")="x86" && 64 bit OS, 32 bit VFP 
              m.liRoot=HKEY_LOCAL_MACHINE
              m.lcBaseKey="SOFTWARE\ODBC\ODBCINST.INI"

         OTHERWISE && 64 bit OS, 64 bit VFP 
              m.liRoot=HKEY_LOCAL_MACHINE
              m.lcBaseKey="SOFTWARE\ODBC\ODBCINST.INI"
      
      ENDCASE

      DO WHILE (INLIST(m.retVal, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO))
         m.szDsn = SPACE(MAX_STRING+1)   && Make sure there is enough space
         m.szDescript = SPACE(MAX_STRING+1)
         m.retval = DBF2ANY_SQLDrivers(m.hEnv, m.fDirection, @m.szDSN, m.cbDSNMax, @m.pcbDSN, @m.szDescript, m.cbDescriptMax, @m.pcbDescript)
         IF INLIST(m.retVal, SQL_SUCCESS, SQL_SUCCESS_WITH_INFO)
            DIMENSION This.aExistsDrivers(m.lii, 2)
            This.aExistsDrivers(m.lii, 1)=LEFT(m.szDSN, m.pcbDSN)
            This.aExistsDrivers(m.lii, 2)=""

            m.lcKey=m.lcBaseKey+"\"+This.aExistsDrivers(m.lii, 1)
            m.liHKEY=0
            IF DBF2ANY_RegOpenKeyEx(m.liRoot, @m.lcKey, 0, STANDARD_RIGHTS_READ+KEY_WOW64_32KEY+KEY_QUERY_VALUE, @m.liHKEY)=REGAPI_SUCCESS
               m.lcName="Driver"
               m.lcData=REPLICATE(CHR(0), 255)
               m.liData=255
               m.liType=REG_EXPAND_SZ
               IF DBF2ANY_RegQueryValueEx(m.liHKEY, @m.lcName, 0, @m.liType, @m.lcData, @m.liData)=REGAPI_SUCCESS
                  m.lcIN=LEFT(m.lcData, m.liData-1)
                  m.lcOUT=SPACE(255)
                  m.liy=DBF2ANY_ExpandEnvironmentStrings(@m.lcIN, @m.lcOUT, LEN(m.lcOUT))
                  m.lcIN=LEFT(m.lcOUT, m.liy-1)

                  =AGETFILEVERSION(m.laVers, m.lcIN)
                  This.aExistsDrivers(m.lii, 2)=m.laVers(3) && Name of driver
               ENDIF
               =DBF2ANY_RegCloseKey(m.liHKEY)
            ENDIF
            
            * We only want to do the SQL_FETCH_FIRST the first time
            m.lii=m.lii+1
         ENDIF
         m.fDirection = SQL_FETCH_NEXT
      ENDDO
   ENDPROC


   PROCEDURE ReadInstalledOLEDBDrivers
      * 
      * _dbf2any::ReadInstalledOLEDBDrivers
      * 
      
      LOCAL m.lii, m.liCNT, m.lcKey, m.liErr, m.liIndex, m.lcValue, m.lcData, m.liData, m.liValue, m.liHKEY,;
            m.lcClass, m.liCount, m.lcName, m.lcDT, m.liReserved, m.liHKEY2, m.lcKey2, m.liType,;
            m.liHKEY3, m.lcKey3, m.liOLEDBD, m.lcOLEDBDrivers, m.lcOLEDBDesc
      LOCAL ARRAY m.laBranch(1, 2)

      m.liCNT=1
      DO CASE
         CASE LEN(GETENV("ProgramFiles(x86)"))=0 && 32 bit OS
              m.laBranch(1, 1)=HKEY_LOCAL_MACHINE
              m.laBranch(1, 2)="SOFTWARE\Classes\CLSID"
         
         CASE LEN(GETENV("ProgramFiles(x86)"))>0 AND GETENV("PROCESSOR_ARCHITECTURE")="x86" && 64 bit OS, 32 bit VFP 
              m.laBranch(1, 1)=HKEY_LOCAL_MACHINE
              m.laBranch(1, 2)="SOFTWARE\Classes\Wow6432Node\CLSID"

         OTHERWISE && 64 bit OS, 64 bit VFP 
              m.laBranch(1, 1)=HKEY_LOCAL_MACHINE
              m.laBranch(1, 2)="SOFTWARE\Classes\CLSID"
      
      ENDCASE

      m.liOLEDBD=0
      FOR m.lii=1 TO m.liCNT
          m.lcKey=m.laBranch(m.lii, 2)
          m.liHKEY=0
          
          IF DBF2ANY_RegOpenKeyEx(m.laBranch(m.lii, 1), @m.lcKey, 0, STANDARD_RIGHTS_READ+KEY_WOW64_32KEY+KEY_ENUMERATE_SUB_KEYS, @m.liHKEY)<>REGAPI_SUCCESS
             LOOP
          ENDIF

          STORE 0 TO m.liIndex
          DO WHILE .T.
             STORE 0 TO m.liReserved
             STORE 255 TO m.liCount, m.liMaxClassLen
             STORE REPLICATE(CHR(0), 255) TO m.lcName, m.lcClass
             m.lcDT=REPLICATE(CHR(0),8)
          
             m.liErr=DBF2ANY_RegEnumKeyExA(m.liHKEY, m.liIndex, @m.lcName, @m.liCount, m.liReserved, @m.lcClass, @m.liMaxClassLen, @m.lcDT)
             IF m.liErr=REGAPI_NO_MORE_ITEMS
                m.liIndex=m.liIndex-1
                EXIT 
             ENDIF

             m.lcKey2=m.lcKey+"\"+LEFT(m.lcName, m.liCount)
             m.liHKEY2=0
             IF DBF2ANY_RegOpenKeyEx(m.laBranch(m.lii, 1), @m.lcKey2, 0, STANDARD_RIGHTS_READ+KEY_WOW64_32KEY+KEY_QUERY_VALUE, @m.liHKEY2)=REGAPI_SUCCESS
                m.lcName="OLEDB_SERVICES"
                m.lcData=REPLICATE(CHR(0), 4)
                m.liData=4
                m.liType=REG_DWORD
                IF DBF2ANY_RegQueryValueEx(m.liHKEY2, @m.lcName, 0, @m.liType, @m.lcData, @m.liData)=REGAPI_SUCCESS
                   IF m.lcData<>CHR(0x00)+CHR(0x00)+CHR(0x00)+CHR(0x00)
                      STORE "" TO m.lcOLEDBDrivers, m.lcOLEDBDesc
                      m.lcKey3=m.lcKey2+"\ProgID"
                      m.liHKEY3=0
                      
                      IF DBF2ANY_RegOpenKeyEx(m.laBranch(m.lii, 1), @m.lcKey3, 0, STANDARD_RIGHTS_READ+KEY_WOW64_32KEY+KEY_QUERY_VALUE, @m.liHKEY3)=REGAPI_SUCCESS
                         m.lcName=""
                         m.lcData=REPLICATE(CHR(0), 255)
                         m.liData=255
                         m.liType=REG_SZ
                         IF DBF2ANY_RegQueryValueEx(m.liHKEY3, @m.lcName, 0, @m.liType, @m.lcData, @m.liData)=REGAPI_SUCCESS
                            m.lcOLEDBDrivers=LEFT(m.lcData, m.liData-1)
                         ENDIF
                         =DBF2ANY_RegCloseKey(m.liHKEY3)
                      ENDIF

                      m.lcKey3=m.lcKey2+"\OLE DB Provider"
                      m.liHKEY3=0
                      IF DBF2ANY_RegOpenKeyEx(m.laBranch(m.lii, 1), @m.lcKey3, 0, STANDARD_RIGHTS_READ+KEY_WOW64_32KEY+KEY_QUERY_VALUE, @m.liHKEY3)=REGAPI_SUCCESS
                         m.lcName=""
                         m.lcData=REPLICATE(CHR(0), 255)
                         m.liData=255
                         m.liType=REG_SZ
                         IF DBF2ANY_RegQueryValueEx(m.liHKEY3, @m.lcName, 0, @m.liType, @m.lcData, @m.liData)=REGAPI_SUCCESS
                            m.lcOLEDBDesc=LEFT(m.lcData, m.liData-1)
                         ENDIF
                         =DBF2ANY_RegCloseKey(m.liHKEY3)
                      ENDIF

                      IF LEN(m.lcOLEDBDrivers)=0
                         m.lcName=""
                         m.lcData=REPLICATE(CHR(0), 255)
                         m.liData=255
                         m.liType=REG_SZ
                         IF DBF2ANY_RegQueryValueEx(m.liHKEY2, @m.lcName, 0, @m.liType, @m.lcData, @m.liData)=REGAPI_SUCCESS
                            m.lcOLEDBDrivers=LEFT(m.lcData, m.liData-1)
                         ENDIF
                      ENDIF

                      IF LEN(m.lcOLEDBDesc)=0
                         m.lcName=""
                         m.lcData=REPLICATE(CHR(0), 255)
                         m.liData=255
                         m.liType=REG_SZ
                         IF DBF2ANY_RegQueryValueEx(m.liHKEY2, @m.lcName, 0, @m.liType, @m.lcData, @m.liData)=REGAPI_SUCCESS
                            m.lcOLEDBDesc=LEFT(m.lcData, m.liData-1)
                         ENDIF
                      ENDIF

                      m.liOLEDBD=m.liOLEDBD+1
                      DIMENSION This.aExistsDrivers(m.liOLEDBD, 2)
                      This.aExistsDrivers(m.liOLEDBD, 1)=m.lcOLEDBDrivers
                      This.aExistsDrivers(m.liOLEDBD, 2)=m.lcOLEDBDesc
                   ENDIF
                ENDIF
                =DBF2ANY_RegCloseKey(m.liHKEY2)
             ENDIF
             
             m.liIndex=m.liIndex+1
          ENDDO
          =DBF2ANY_RegCloseKey(m.liHKEY)
      NEXT
   ENDPROC


   PROCEDURE ReadInstalledDrivers
      * 
      * _dbf2any::ReadInstalledDrivers
      * 

      IF NOT EMPTY(This.aExistsDrivers(1))
         RETURN
      ENDIF

      IF This.nRemoteMode=__DBF2ANY_Mode_ODBC
         This.ReadInstalledODBCDrivers()
      ELSE
         This.ReadInstalledOLEDBDrivers()
      ENDIF
   ENDPROC


   PROCEDURE CheckDrivers
      * 
      * _dbf2any::CheckDrivers
      * 
      LPARAMETERS m.luDriver

      LOCAL m.lii, m.liy, m.liStart, m.liCNT
      DO CASE 
         CASE EMPTY(m.luDriver)
              m.liStart=1
              m.liCNT=This.iDriver

         CASE TYPE("m.luDriver")="N"
              STORE m.luDriver TO m.liStart, m.liCNT
     
         CASE TYPE("m.luDriver")="C"
              m.luDriver=UPPER(m.luDriver)
              STORE 0 TO m.liStart, m.liCNT
              FOR m.lii=1 TO This.iDriver
                  IF UPPER(This.aDriver(m.lii))==m.luDriver
                     STORE m.lii TO m.liStart, m.liCNT
                     EXIT
                  ENDIF
              NEXT

      ENDCASE
      
      This.iSelectedDriver=0
      FOR m.lii=m.liStart TO m.liCNT
          FOR m.liy=1 TO ALEN(This.aExistsDrivers, 1)
              IF UPPER(This.aDriver(m.lii))==UPPER(This.aExistsDrivers(m.liy, 1))
                 This.iSelectedDriver=m.lii
                 EXIT
              ENDIF
          NEXT
          IF This.iSelectedDriver>0
             EXIT
          ENDIF
      NEXT
   ENDPROC


   PROCEDURE CreateSQLCommand
      * 
      * _dbf2any::CreateSQLCommand
      * 
      LPARAMETERS m.loPar
   ENDPROC 


   PROCEDURE CheckTableIfExists
      * 
      * _dbf2any::CheckTableIfExists
      * 
      LPARAMETERS m.loPar
   ENDPROC


   PROCEDURE Upload2Remote
      * 
      * _dbf2any::Upload2Remote
      * 
      LPARAMETERS m.loPar

      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE BeforeUpload2Remote
      * 
      * _dbf2any::BeforeUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE AfterUpload2Remote
      * 
      * _dbf2any::AfterUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE DropTable
      * 
      * _dbf2any::DropTable
      * 
      LPARAMETERS m.loPar

      LOCAL m.liErr
      m.liErr=__DBF2ANY_ErrOK

      IF This.nRemoteMode=__DBF2ANY_Mode_ODBC
         IF SQLEXEC(This.hdbc, This.cDropSQL)<=0
            m.liErr=__DBF2ANY_ErrCannotDTable
            =AERROR(m.loPar.laErr)
         ENDIF
      ELSE
         This.SetCatch(.T.)
         This.oConnection.Execute(This.cDropSQL)
         IF This.nErr<>0
            m.liErr=__DBF2ANY_ErrCannotDTable
            =AERROR(m.loPar.laErr)
         ENDIF
         This.SetCatch(.NULL.)
      ENDIF

      RETURN m.liErr
   ENDPROC 


   PROCEDURE CreateTable
      * 
      * _dbf2any::CreateTable
      * 
      LPARAMETERS m.loPar

      LOCAL m.liErr
      m.liErr=__DBF2ANY_ErrOK

      IF This.nRemoteMode=__DBF2ANY_Mode_ODBC
         IF SQLEXEC(This.hdbc, This.cCreateSQL)<=0
            m.liErr=__DBF2ANY_ErrCannotCTable
            =AERROR(m.loPar.laErr)
         ENDIF
      ELSE
         This.SetCatch(.T.)
         This.oConnection.Execute(This.cCreateSQL)
         IF This.nErr<>0
            m.liErr=__DBF2ANY_ErrCannotCTable
            =AERROR(m.loPar.laErr)
         ENDIF
         This.SetCatch(.NULL.)
      ENDIF

      RETURN m.liErr
   ENDPROC 
   

   PROCEDURE DeleteData
      * 
      * _dbf2any::DeleteData
      * 
      LPARAMETERS m.loPar

      LOCAL m.liErr
      m.liErr=__DBF2ANY_ErrOK

      IF This.nRemoteMode=__DBF2ANY_Mode_ODBC
         IF SQLEXEC(This.hdbc, This.cDeleteSQL)<=0
            m.liErr=__DBF2ANY_ErrNotDelete
            =AERROR(m.loPar.laErr)
         ENDIF
      ELSE
         This.SetCatch(.T.)
         This.oConnection.Execute(This.cDeleteSQL)
         IF This.nErr<>0
            m.liErr=__DBF2ANY_ErrNotDelete
            =AERROR(m.loPar.laErr)
         ENDIF
         This.SetCatch(.NULL.)
      ENDIF

      RETURN m.liErr
   ENDPROC 
   

   PROCEDURE GetListOfTables
      * 
      * _dbf2any::GetListOfTables
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.liErr
      m.liErr=__DBF2ANY_ErrOK

      IF This.nRemoteMode=__DBF2ANY_Mode_ODBC
         IF SQLTABLES(This.hdbc, "'TABLE','VIEW','SYSTEM TABLE'", m.loPar.cATables)<=0
            m.liErr=__DBF2ANY_ErrNoTables
            =AERROR(m.loPar.laErr)
         ENDIF
      ELSE
         LOCAL m.loRS, m.loCmd, m.loFields
         This.SetCatch(.T.)
         CREATE CURSOR (m.loPar.cATables) (TABLE_CAT  C(128)  NULL, TABLE_SCHEM C(128) NULL, TABLE_NAME C(128) NULL, TABLE_TYPE C(128) NULL, REMARKS C(254) NULL)

         m.loCmd=CREATEOBJECT("ADODB.Command")
         m.loCmd.ActiveConnection=This.oConnection
         m.loCmd.CommandText="SELECT * FROM INFORMATION_SCHEMA.TABLES"
         m.loRS=m.loCmd.Execute()

         IF This.nErr=0 AND NOT m.loRS.BOF
            m.loRS.MoveFirst()
            DO WHILE !m.loRS.EOF
               m.loFields=m.loRS.Fields
               INSERT INTO (m.loPar.cATables) ;
                 VALUES;
                  (m.loFields.Item("TABLE_CATALOG").Value, m.loFields.Item("TABLE_SCHEMA").Value, m.loFields.Item("TABLE_NAME").Value,;
                   m.loFields.Item("TABLE_TYPE").Value, m.loFields.Item("DESCRIPTION").Value)

               m.loRS.MoveNext()
            ENDDO
         ENDIF

         IF This.nErr<>0
            m.liErr=__DBF2ANY_ErrNoTables
            =AERROR(m.loPar.laErr)
         ENDIF

         This.SetCatch(.NULL.)
      ENDIF

      RETURN m.liErr
   ENDPROC 
      
   
   HIDDEN PROCEDURE Upload
      * 
      * _dbf2any::Upload
      * 
      LPARAMETERS m.lcAlias
      
      LOCAL m.liErr, m.lcPom, m.lii, m.lcTable, m.lcATables, m.llExists, m.loPar
      LOCAL ARRAY m.laErr(1)
      
      * Get out if template mdb not defined or not exists
      IF EMPTY(m.lcAlias)
         RETURN __DBF2ANY_ErrFewParameters
      ENDIF

      IF !USED(m.lcAlias)
         RETURN __DBF2ANY_ErrAliasNotOpen
      ENDIF


      m.lcPom=IIF(RIGHT(DBF(m.lcAlias), 4)=".TMP", "CRS_"+m.lcAlias, DBF(m.lcAlias))
      m.lii=RAT("\", m.lcPom)
      m.lcTable=IIF(m.lii>0, SUBS(m.lcPom, m.lii+1), m.lcPom)
      m.lii=RAT(".", m.lcTable)
      m.lcTable=IIF(m.lii>0, LEFT(m.lcTable, m.lii-1), m.lcTable)

      m.llDal=.T. && Internal flag, All are OK

      * Table is exist in MSSQL
      m.lcATables=SYS(2015) && 
      m.loPar=CREATEOBJECT("_dbf2any_Upload2Remote", m.lcATables, m.lcAlias, m.lcTable)
      m.liErr=This.GetListOfTables(@m.loPar) 
      m.llDal=m.liErr=__DBF2ANY_ErrOK

      IF m.llDal
         m.llExists=This.CheckTableIfExists(@m.loPar)

         * Test if table exist and can be recreate, but is it View or System Table
         IF m.llExists AND This.lReCreate AND (UPPER(ALLT(TABLE_TYPE))=='VIEW' OR UPPER(ALLT(TABLE_TYPE))=='SYSTEM TABLE')
            m.llDal=.F.
            m.liErr=__DBF2ANY_ErrNotaTable
            =AERROR(m.loPar.laErr)
         ENDIF
      ENDIF

      IF m.llDal
         This.CreateSQLCommand(@m.loPar)
      ENDIF
      
      IF m.llDal AND (!m.llExists OR This.lReCreate)
         * Drop table
         IF m.llExists
            m.liErr=This.DropTable(@m.loPar)
            m.llDal=m.liErr=__DBF2ANY_ErrOK
         ENDIF

         * Create Table
         IF m.llDal
            m.liErr=This.CreateTable(@m.loPar)
            m.llDal=m.liErr=__DBF2ANY_ErrOK
         ENDIF
      ENDIF

      IF USED(m.lcATables) && Close cursor
         USE IN (m.lcATables)
      ENDIF

      * Delete all records if table exist and not be create and be fill
      IF m.llDal AND m.llExists AND !This.lReCreate AND This.lReFill
         m.liErr=This.DeleteData(@m.loPar)
         m.llDal=m.liErr=__DBF2ANY_ErrOK
      ENDIF

      IF m.llDal
         m.liErr=This.BeforeUpload2Remote(@m.loPar)

         IF m.liErr=__DBF2ANY_ErrOK
            m.liErr=This.Upload2Remote(@m.loPar)
         ENDIF

         * If not Auto Transaction, commit changes or rollback it
         IF This.hdbc>0
            =IIF(SQLGETPROP(This.hdbc, "Transactions")=2,IIF(m.liErr=__DBF2ANY_ErrOK, SQLCOMMIT(This.hdbc), SQLROLLBACK(This.hdbc)), .T.)
            =SQLSETPROP(This.hdbc, "Transactions", 1)
         ENDIF

         IF m.liErr=__DBF2ANY_ErrOK
            m.liErr=This.AfterUpload2Remote(@m.loPar)
         ENDIF

      ENDIF
      m.loPar.iErr=m.liErr
      =This.Uploaded(@m.loPar)
      RETURN m.liErr
   ENDPROC


   PROCEDURE GetVersion
      * 
      * _dbf2any::GetVersion
      * 
      RETURN This.Version
   ENDPROC


   PROCEDURE UploadDBC && Upload all tables from DBC 
      * 
      * _dbf2any::UploadDBC
      * 
      LPARAMETERS m.lcDBC, m.llOpen
      * lcDBC  - Folder and Name od DBC
      * llOpen - Open flag
      
      LOCAL m.lcAlias, m.liCount, m.lii, m.liErr

      m.liErr=This.OpenRemoteDB()
      IF m.liErr#__DBF2ANY_ErrOK
         RETURN m.liErr
      ENDIF

      IF IIF(m.llOpen, This.OpenDBC(-1, m.lcDBC, "")#__DBF2ANY_ErrOK, !DBUSED(m.lcDBC))
         RETURN __DBF2ANY_ErrDBCNotOpen
      ENDIF

      m.liCount=ADBOBJECTS(m.laTables, "TABLE")
      m.lcAlias=SYS(2015)

      FOR m.lii=1 TO m.liCount
          IF This._UploadDBF(m.laTables(m.lii), m.lcAlias, "IN 0 NOUPDATE")=__DBF2ANY_ErrOK && Open table
       
          ELSE

          ENDIF
      NEXT
      This.CloseRemoteDB()
      IF m.llOpen
         CLOSE DATABASE
      ENDIF
      RETURN __DBF2ANY_ErrOK
   ENDPROC


   HIDDEN PROCEDURE _UploadDBF
      * 
      * _dbf2any::_UploadDBF
      * 
      LPARAMETERS m.lcDBF, m.lcAlias, m.lcParams
      
      * lcDBF    - DBF
      * lcAlias  - Alias
      * lcParams - Parameters

      LOCAL m.liErr
      m.liErr=IIF(EMPTY(m.lcDBF),;
                  IIF(USED(m.lcAlias), __DBF2ANY_ErrOK, __DBF2ANY_ErrAliasNotOpen),;
                      This.OpenTable(-1, m.lcDBF, @m.lcAlias, m.lcParams))
      IF m.liErr=__DBF2ANY_ErrOK && Open table
         m.liErr=This.Upload(m.lcAlias)
         =IIF(EMPTY(m.lcDBF), .T., This.CloseTable(-1, m.lcAlias))
      ENDIF

      RETURN m.liErr
   ENDPROC


   PROCEDURE UploadDBF
      * 
      * _dbf2any::UploadDBF
      * 
      LPARAMETERS m.lcDBF, m.lcAlias, m.lcParams
      * lcDBF    - DBF
      * lcAlias  - Alias
      * lcParams - Parameters

      LOCAL m.liErr
      m.liErr=This.OpenRemoteDB()
      IF m.liErr#__DBF2ANY_ErrOK
         RETURN m.liErr
      ENDIF

      m.liErr=IIF(EMPTY(m.lcDBF),;
                 IIF(USED(m.lcAlias), __DBF2ANY_ErrOK, __DBF2ANY_ErrAliasNotOpen),;
                     This.OpenTable(-1, m.lcDBF, @m.lcAlias, m.lcParams))
      IF m.liErr=__DBF2ANY_ErrOK && Open table
         m.liErr=This.Upload(m.lcAlias)
         =IIF(EMPTY(m.lcDBF), .T., This.CloseTable(-1, m.lcAlias))
      ENDIF

      This.CloseRemoteDB()

      RETURN m.liErr
   ENDPROC


   PROCEDURE Error
      * 
      * _dbf2any::Error
      * 
      LPARAMETERS m.nError, m.cMethod, m.nLine
      
      * nError  - Error number
      * cMethod - Method name
      * nLine   - Row number

      IF ISNULL(m.This.Catch)
         This.nErr=m.nError
         This.cErrMSG=MESSAGE()
         RETURN
      ENDIF

      IF This.Catch
         This.nErr=m.nError
         This.cErrMSG=MESSAGE()
      ELSE
      ENDIF
   ENDPROC


   PROCEDURE SetCatch && Seting mode of catching error
      * 
      * _dbf2any::SetCatch
      * 
      LPARAMETERS m.llCatch
       
      LOCAL m.luCatch
      m.luCatch=m.This.Catch
      This.Catch=m.llCatch
      IF llCatch && If have error catch
         * Reset Error info
         This.nErr=0
         This.cErrMSG=""
      ENDIF
      RETURN m.luCatch && Return old value
   ENDPROC


   PROCEDURE OpenDBC()
      * 
      * _dbf2any::OpenDBC
      * 
      LPARAMETERS m.liDSID, m.lcDBC, m.lcParam
      
      * liDSID  - ID Data Session
      * lcDBC  -  Directory and Name DBC container
      * lcParam - Parameters

      LOCAL m.liDSIDOld, m.luCatch
      m.liDSIDOld=This.DataSessionID &&  Remember Datasession ID
      m.luCatch=This.SetCatch(.T.) && Set chatch error
      IF m.liDSID>=0
         SET DATASESSION TO m.liDSID && Set new Datasession ID
      ENDIF
      OPEN DATA (m.lcDBC) &lcParam.
      SET DATASESSION TO (m.liDSIDOld) && Set datasession ID
      This.SetCatch(m.luCatch) && Cancel chatch error
      RETURN This.nErr && Return number Error
   ENDPROC


   PROCEDURE OpenTable
      * 
      * _dbf2any::OpenTable
      * 
      LPARAMETERS m.liDSID, m.lcFile, m.lcAlias, m.lcParam
      
      * liDSID  - ID Data Session
      * lcFile  - Folder and name DBF file
      * lcAlias - Alias 
      * lcParam - Parameters

      LOCAL m.liDSIDOld, m.luCatch
      m.liDSIDOld=m.This.DataSessionID
      m.luCatch=m.This.SetCatch(.T.)
      IF m.liDSID>=0
         SET DATASESSION TO m.liDSID
      ENDIF
      IF EMPTY(m.lcAlias)
         m.lcAlias=SYS(2015)
      ENDIF
      USE (m.lcFile) ALIAS (m.lcAlias) &lcParam. IN 0
      SET DATASESSION TO (m.liDSIDOld)
      This.SetCatch(m.luCatch)
      RETURN This.nErr
   ENDPROC


   PROCEDURE CloseTable
      * 
      * _dbf2any::CloseTable
      * 
      LPARAMETERS m.liDSID, m.lcAlias
      
      * liDSID  - ID Data Session
      * lcAlias - Alias 
      LOCAL m.liDSIDOld
      m.liDSIDOld=m.This.DataSessionID
      IF m.liDSID>=0
         SET DATASESSION TO m.liDSID
      ENDIF
      IF USED(m.lcAlias)
         USE IN (m.lcAlias)
      ENDIF
      SET DATASESSION TO (m.liDSIDOld)
      RETURN This.nErr
   ENDPROC


   PROCEDURE Uploaded && Event - after uploading table
      * 
      * _dbf2any::Uploaded
      * 
      LPARAMETERS m.loPar
      
   ENDPROC


   PROCEDURE AttachDataSession
      * 
      * _dbf2any::AttachDataSession
      * 
      LPARAMETERS m.liDSID
      
      SET DATASESSION TO m.liDSID
   ENDPROC


   PROCEDURE RestoreDataSession
      * 
      * _dbf2any::RestoreDataSession
      * 
      LPARAMETERS m.liDSID
      
      SET DATASESSION TO (This._DataSessionID)
   ENDPROC


   PROCEDURE GetConnectionString
      * 
      * _dbf2any::GetConnectionString
      * 
      LPARAMETERS m.lcDriver
      
      RETURN ""
   ENDPROC


   PROCEDURE OpenRemoteDB 
      * 
      * _dbf2any::OpenRemoteDB
      * 

      IF This.iSelectedDriver=0
         This.CheckDrivers()
      ENDIF
      
      IF This.nRemoteMode=__DBF2ANY_Mode_ODBC
         LOCAL m.liErr, m.lihdbc, m.lii

         =SQLSETPROP(0,"Asynchronous",.F.) && Asynchronous mode
         =SQLSETPROP(0,"Transactions",1)   && Automatic transaction
         =SQLSETPROP(0,"DispLogin",3)      && Don't show dialog

         IF This.iSelectedDriver>0 && driver exists, maybe
            m.lihdbc=SQLSTRINGCONNECT(This.GetConnectionString(This.aDriver(This.iSelectedDriver)))
            IF m.lihdbc<0
               This.iSelectedDriver=0 && hmm, failed, check all drivers
            ENDIF
         ENDIF

         IF This.iSelectedDriver=0
            FOR m.lii=1 TO This.iDriver
                m.lihdbc=SQLSTRINGCONNECT(This.GetConnectionString(This.aDriver(m.lii)))
                IF m.lihdbc>=0
                   This.iSelectedDriver=m.lii
                   EXIT
                ENDIF
            NEXT
         ENDIF
         IF m.lihdbc<=0
            =AERROR(This.aErrors)
            RETURN __DBF2ANY_ErrConnectionFailed
         ENDIF

         =SQLSETPROP(m.lihdbc, "Asynchronous", .F.) && Set Asynchronnous 
         =SQLSETPROP(m.lihdbc, "Transactions", 1)   && Set Auto Transactions
         =SQLSETPROP(m.lihdbc, "DispLogin", 3)      && Don't show dialog for login

         This.hdbc=m.lihdbc
      ELSE
         LOCAL m.lii
         This.SetCatch(.T.)
         This.oConnection = CREATEOBJECT('ADODB.Connection')
         IF This.iSelectedDriver>0 && driver exists, maybe
            This.oConnection.Open(This.GetConnectionString(This.aDriver(This.iSelectedDriver)))
            IF This.nErr<>0
               This.iSelectedDriver=0 && hmm, failed, check all drivers
            ENDIF
         ENDIF
         
         IF This.iSelectedDriver=0
            FOR m.lii=1 TO This.iDriver
                This.SetCatch(.T.)
                This.oConnection.Open(This.GetConnectionString(This.aDriver(m.lii)))
                IF This.nErr=0
                   This.iSelectedDriver=m.lii
                   EXIT
                ENDIF
            NEXT
         ENDIF

         IF This.nErr<>0
             This.SetCatch(.NULL.)
            =AERROR(This.aErrors)
            RETURN __DBF2ANY_ErrConnectionFailed
         ENDIF
         This.SetCatch(.NULL.)
      
      ENDIF
      RETURN __DBF2ANY_ErrOK
   ENDPROC


   PROCEDURE CloseRemoteDB
      * 
      * _dbf2any::CloseRemoteDB
      * 

      IF This.nRemoteMode=__DBF2ANY_Mode_ODBC
         IF This.hdbc>0
            =SQLDISCONNECT(This.hdbc)
            This.hdbc=0
         ENDIF
      ELSE
         IF NOT ISNULL(This.oConnection)
            This.oConnection.Close()
            This.oConnection=.NULL.
         ENDIF
      ENDIF

   ENDPROC

ENDDEFINE
