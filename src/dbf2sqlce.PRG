#INCLUDE "dbf2any.h"
#INCLUDE "dbf2any_API.h"



DEFINE CLASS _dbf2sqlce AS _dbf2any && Upload DBF file to MSSQL CE

   * System properties
   Version="0.0.0.2"  && Version
   nRemoteMode=__DBF2ANY_Mode_OLEDB

   iDriver=2          && Drivers count
   DIMENSION aDriver(2)
   aDriver(1)="Microsoft.SQLSERVER.CE.OLEDB.4.0"
   aDriver(2)="Microsoft.SQLSERVER.CE.OLEDB.3.5"

   cDB    =""
   cPWD   =""
   lDBReCreate=.F.
   
   cInsertSQLEx=""


   PROCEDURE OpenRemoteDB
      * 
      * _dbf2sqlce::OpenRemoteDB
      * 

      IF !FILE(This.cDB) OR This.lDBReCreate
         m.liErr=This.CreateRemoteDB(This.cDB)
         IF m.liErr#__DBF2ANY_ErrOK
            RETURN m.liErr
         ENDIF
      ENDIF

      RETURN DODEFAULT()
   ENDPROC
   

   PROCEDURE CreateRemoteDB && Create new SDF file
      * 
      * _dbf2sqlce::CreateRemoteDB
      * 
      LPARAMETERS m.lcDB

      * lcDB  - Target SDF file

      LOCAL m.loCat
      m.loCat=CREATEOBJECT("ADOX.Catalog")

      IF FILE(m.lcDB) && If SDF file exist
         This.DeleteFile(m.lcDB) && delete it
      ENDIF

      This.SetCatch(.T.)
      m.loCat.Create(This.GetConnectionString(This.aDriver(This.iSelectedDriver)))

      IF This.nErr<>0
          This.SetCatch(.NULL.)
         =AERROR(This.aErrors)
         RETURN __DBF2ANY_ErrDBNotCreated
      ENDIF
      This.SetCatch(.NULL.)
      
      RETURN __DBF2ANY_ErrOK
   ENDPROC


   PROCEDURE CreateSQLCommand
      * 
      * _dbf2sqlce::CreateSQLCommand
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lii
      LOCAL ARRAY m.laFields(1)

      STORE "" TO m.lcSQLC, m.lcSQLIV, m.lcSQLI
      * Read field list
      FOR m.lii=1 TO AFIELDS(m.laFields, m.loPar.cAlias)
          m.lcSQLC=m.lcSQLC+[ "]+m.laFields(m.lii, 1)+[" ]+;
                   IIF(m.laFields(m.lii, 2)='C', "nchar("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='V', "nvarchar("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='Q', "varbinary("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='M', "ntext",;
                   IIF(m.laFields(m.lii, 2)='Y', "decimal(20,4)",;
                   IIF(m.laFields(m.lii, 2)='F', "float",;
                   IIF(m.laFields(m.lii, 2)='B', "float",;
                   IIF(m.laFields(m.lii, 2)='I', "integer",;
                   IIF(m.laFields(m.lii, 2)='N', "decimal("+LTRIM(STR(m.laFields(m.lii, 3), 3))+IIF(m.laFields(m.lii, 4)=0,")",","+LTRIM(STR(m.laFields(m.lii, 4), 3))+")") ,;
                   IIF(m.laFields(m.lii, 2)$'D,T', "datetime",;
                   IIF(m.laFields(m.lii, 2)='G', 'image',;
                   IIF(m.laFields(m.lii, 2)='W', 'image', "decimal(1)"))))))))))))+" "+IIF(m.laFields(m.lii, 5),"NULL,","NOT NULL,")

          m.lcSQLI=m.lcSQLI+" ["+m.laFields(m.lii, 1)+"], "
          DO CASE
             CASE INLIST(m.laFields(m.lii, 2), 'C', 'V', 'M')
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyString(]+m.laFields(m.lii, 1)+[)+','+]

             CASE m.laFields(m.lii, 2)='I'
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyInt(]+m.laFields(m.lii, 1)+[)+','+]

             CASE m.laFields(m.lii, 2)='D'
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyDate(]+m.laFields(m.lii, 1)+[)+','+]

             CASE m.laFields(m.lii, 2)='T'
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyDateTime(]+m.laFields(m.lii, 1)+[)+','+]

             CASE m.laFields(m.lii, 2)='L'
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyBool(]+m.laFields(m.lii, 1)+[)+','+]

             CASE INLIST(m.laFields(m.lii, 2), 'Q', 'W')
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyBinary(]+m.laFields(m.lii, 1)+[)+','+]

             CASE m.laFields(m.lii, 2)= 'G'
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyBinary(CAST(]+m.laFields(m.lii, 1)+[ AS W))+','+]

             CASE INLIST(m.laFields(m.lii, 2), 'N', 'F')
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyDecimal(]+m.laFields(m.lii, 1)+[,]+LTRIM(STR(m.laFields(m.lii, 3), 3))+[,]+LTRIM(STR(m.laFields(m.lii, 4), 3))+[)+','+]

             CASE m.laFields(m.lii, 2)='Y'
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyDecimal(]+m.laFields(m.lii, 1)+[,20,4)+','+]

             CASE m.laFields(m.lii, 2)='B'
                  m.lcSQLIV=m.lcSQLIV+[This.SafetyDecimal(]+m.laFields(m.lii, 1)+[,20,]+LTRIM(STR(m.laFields(m.lii, 4), 3))+[)+','+]

          ENDCASE
      NEXT
      This.cCreateSQL="CREATE TABLE "+m.loPar.cTable+" ("+LEFT(m.lcSQLC, LEN(m.lcSQLC)-1)+")"
      This.cInsertSQL="INSERT INTO "+m.loPar.cTable+" ("+LEFT(m.lcSQLI, LEN(m.lcSQLI)-2)+") VALUES (%FIELDS%)"
      This.cInsertSQLEx=LEFT(m.lcSQLIV, LEN(m.lcSQLIV)-5)

      This.cDropSQL=[DROP TABLE "]+m.loPar.cTable+["]
      This.cDeleteSQL=[DELETE FROM  "]+m.loPar.cTable+["]

   ENDPROC 


   PROCEDURE CheckTableIfExists
      * 
      * _dbf2sqlce::CheckTableIfExists
      * 
      LPARAMETERS m.loPar
      
      SELE (m.loPar.cATables)
      LOCATE FOR UPPER(ALLT(TABLE_NAME))==UPPER(m.loPar.cTable)
      RETURN FOUND()
   ENDPROC 


   PROCEDURE Upload2Remote
      * 
      * _dbf2sqlce::Upload2Remote
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.liErr, m.loCMD, m.lii, m.liFields
      LOCAL ARRAY m.laFields(1)

      m.liErr=__DBF2ANY_ErrOK

      m.loCMD = CREATEOBJECT("ADODB.Command")
      m.loCMD.ActiveConnection=This.oConnection
      m.loCMD.CommandType = 1
      
      SELE (m.loPar.cAlias)
      GO TOP

      This.SetCatch(.T.)
      SCAN REST
           m.loCmd.CommandText=STRTRAN(This.cInsertSQL, "%FIELDS%", EVALUATE(This.cInsertSQLEx))
           =loCMD.Execute()
           IF This.nErr<>0
              m.liErr=__DBF2ANY_ErrNotInsert
              =AERROR(m.loPar.laErr)
              EXIT
           ENDIF
      ENDSCAN
      This.SetCatch(.NULL.)

      RETURN m.liErr
   ENDPROC 


   PROCEDURE BeforeUpload2Remote
      * 
      * _dbf2sqlce::BeforeUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE AfterUpload2Remote
      * 
      * _dbf2sqlce::AfterUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      This.lDBReCreate=.F. && reset the flag
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE GetConnectionString
      * 
      * _dbf2sqlce::GetConnectionString
      * 
      LPARAMETERS m.lcDriver
      
      RETURN "Provider="+m.lcDriver+";Data Source="+This.cDB+";ssce:database password="+This.cPWD+";persist security info=false;"+This.cExtendedCS
   ENDPROC


   PROCEDURE SafetyString
      * 
      * _dbf2sqlce::SafetyString
      * 
      LPARAMETERS m.lcText

      IF ISNULL(m.lcText)
         RETURN "NULL"
      ENDIF

      LOCAL m.lii
      m.lcText="N'"+STRTRAN(m.lcText, "'", "''")+"'"
      FOR m.lii=0 TO 31
          m.lcText=STRTRAN(m.lcText, CHR(m.lii), "'+CHAR("+LTRIM(STR(m.lii,2))+")+'")
      NEXT
      RETURN STRTRAN(m.lcText,"+''","")
   ENDPROC


   PROCEDURE SafetyDecimal
      * 
      * _dbf2sqlce::SafetyDecimal
      * 
      LPARAMETERS m.lnVal, m.liLEN, m.liDec
      RETURN IIF(ISNULL(m.lnVal), "NULL", CHRTRAN(LTRIM(STR(m.lnVal, m.liLEN, m.liDec)),',','.'))
   ENDPROC


   PROCEDURE SafetyInt
      * 
      * _dbf2sqlce::SafetyInt
      * 
      LPARAMETERS m.lnVal
      RETURN IIF(ISNULL(m.lnVal), "NULL", LTRIM(STR(m.lnVal, 11)))
   ENDPROC


   PROCEDURE SafetyDate
      * 
      * _dbf2sqlce::SafetyDate
      * 
      LPARAMETERS m.ldDate
      RETURN IIF(ISNULL(m.ldDate), "NULL", "CONVERT(DATETIME,'"+CHRTRAN(TTOC(m.ldDate, 3),"T", " ")+"', 120)")
   ENDPROC


   PROCEDURE SafetyDateTime
      * 
      * _dbf2sqlce::SafetyDateTime
      * 
      LPARAMETERS m.ldDateTime
      RETURN IIF(ISNULL(m.ldDateTime), "NULL", "CONVERT(DATETIME,'"+CHRTRAN(TTOC(m.ldDateTime, 3),"T", " ")+"', 120)")
   ENDPROC


   PROCEDURE SafetyBool
      * 
      * _dbf2sqlce::SafetyBool
      * 
      LPARAMETERS m.llVal
      RETURN IIF(ISNULL(m.llVal), "NULL", IIF(m.llVal, '1', '0'))
   ENDPROC


   PROCEDURE SafetyBinary
      * 
      * _dbf2sqlce::SafetyBinary
      * 
      LPARAMETERS m.lgVal
      RETURN IIF(ISNULL(m.lgVal), "NULL", "0X"+TRANSFORM(m.lgVal))
   ENDPROC


ENDDEFINE
