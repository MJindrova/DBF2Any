#INCLUDE "dbf2any.h"
#INCLUDE "dbf2any_API.h"

DEFINE CLASS _dbf2mysql AS _dbf2any && Upload DBF file to MYSQL

   * System properties
   Version="0.0.0.2"  && Version

   iDriver=5          && Drivers count
   DIMENSION aDriver(5)
   aDriver(1)="MySQL ODBC 5.2 ANSI Driver"
   aDriver(2)="MySQL ODBC 5.2 UNICODE Driver"
   aDriver(3)="MySQL ODBC 5.1 ANSI Driver"
   aDriver(4)="MySQL ODBC 5.1 UNICODE Driver"
   aDriver(5)="MySQL ODBC 3.51 Driver"


   cServer =""
   cDB     =""
   cUser   =""
   cPWD    =""
   iPort   =3306
   iOPTION =131073
   cSTMT   =""


   PROCEDURE CreateSQLCommand
      * 
      * _dbf2mysql::CreateSQLCommand
      * 
      LPARAMETERS m.loPar
      
      LOCAL m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV, m.lii
      LOCAL ARRAY m.laFields(1)

      STORE "" TO m.lcSQLC, m.lcSQLIV, m.lcSQLI, m.lcSQLIBV
      * Read field list
      
      FOR m.lii=1 TO AFIELDS(laFields, m.loPar.cAlias)
          m.lcSQLC=m.lcSQLC+" `"+laFields(m.lii, 1)+"` "+;
                   IIF(m.laFields(m.lii, 2)='C', "char ("+LTRIM(STR(laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='V', "varchar("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='Q', "varbinary("+LTRIM(STR(m.laFields(m.lii, 3), 11))+")",;
                   IIF(m.laFields(m.lii, 2)='M', "longtext",;
                   IIF(m.laFields(m.lii, 2)='Y', "decimal(15,4)",;
                   IIF(m.laFields(m.lii, 2)='F', "float",;
                   IIF(m.laFields(m.lii, 2)='B', "double",;
                   IIF(m.laFields(m.lii, 2)='I', "integer(11)",;
                   IIF(m.laFields(m.lii, 2)='N', "decimal("+LTRIM(STR(laFields(m.lii, 3), 3))+IIF(laFields(m.lii, 4)=0,")",","+LTRIM(STR(laFields(m.lii, 4), 3))+")") ,;
                   IIF(m.laFields(m.lii, 2)$'D,T', "datetime",;
                   IIF(m.laFields(m.lii, 2)='G', "longblob",;
                   IIF(m.laFields(m.lii, 2)='W', "longblob", "decimal(1)"))))))))))))+" "+IIF(laFields(m.lii, 5),"NULL,","NOT NULL,")


          m.lcSQLI=m.lcSQLI+" `"+m.laFields(m.lii, 1)+"`, "
          m.lcSQLIV=m.lcSQLIV+"?"+m.loPar.cAlias+"."+m.laFields(m.lii, 1)+", "
          m.lcSQLIBV=m.lcSQLIBV+"?m.laBF(%ROW%)."+m.laFields(m.lii, 1)+", "
      NEXT


      This.cCreateSQL="CREATE TABLE `"+m.loPar.cTable+"` ("+LEFT(m.lcSQLC, LEN(m.lcSQLC)-1)+")"
      This.cInsertSQL="INSERT INTO `"+m.loPar.cTable+"` ("+LEFT(m.lcSQLI, LEN(m.lcSQLI)-2)+") VALUES ("+LEFT(m.lcSQLIV, LEN(m.lcSQLIV)-2)+")"
      This.cInsertBatchSQL="INSERT INTO `"+m.loPar.cTable+"` ("+LEFT(m.lcSQLI, LEN(m.lcSQLI)-2)+") VALUES ("+LEFT(m.lcSQLIBV, LEN(m.lcSQLIBV)-2)+")"
      
      This.cDropSQL=[DROP TABLE `]+m.loPar.cTable+[`]
      This.cDeleteSQL=[DELETE FROM `]+m.loPar.cTable+[`]
   ENDPROC 


   PROCEDURE CheckTableIfExists
      * 
      * _dbf2mysql::CheckTableIfExists
      * 
      LPARAMETERS m.loPar
      
      SELE (m.loPar.cATables)
      LOCATE FOR UPPER(ALLT(TABLE_NAME))==UPPER(m.loPar.cTable)
      RETURN FOUND()
   ENDPROC 


   PROCEDURE Upload2Remote
      * 
      * _dbf2mysql::Upload2Remote
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
      * _dbf2mysql::BeforeUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      =SQLEXEC(This.hdbc,"SET IDENTITY_INSERT dbo."+m.loPar.cTable+" ON")
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE AfterUpload2Remote
      * 
      * _dbf2mysql::AfterUpload2Remote
      * 
      LPARAMETERS m.loPar
      
      =SQLEXEC(This.hdbc,"SET IDENTITY_INSERT dbo."+m.loPar.cTable+" OFF")
      RETURN __DBF2ANY_ErrOK
   ENDPROC 


   PROCEDURE GetConnectionString
      * 
      * _dbf2mysql::GetConnectionString
      * 
      LPARAMETERS m.lcDriver
      
      RETURN "Driver={"+m.lcDriver+"};DB="+This.cDB+";SERVER="+This.cServer+";PORT="+STR(This.iPort,5)+;
             ";OPTION="+STR(This.iOPTION,20)+";STMT="+This.cSTMT+";uid="+This.cUser+";pwd="+This.cPWD+";"+This.cExtendedCS
   ENDPROC

ENDDEFINE


