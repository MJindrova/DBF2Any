#INCLUDE "..\src\dbf2any.h"

LOCAL m.lcPath, m.loDBF2SQLCE, m.lcAlias, m.lii

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


DO (m.lcPath+"..\src\dbf2any") WITH "create", "sqlce", m.loDBF2SQLCE
IF ISNULL(m.loDBF2SQLCE)
   RETURN
ENDIF


m.loDBF2SQLCE=CREATEOBJECT("_dbf2sqlce")
*m.loDBF2SQLCE.cDB=m.lcPath+"Northwind.sdf"

*!* FOR m.lii=1 TO ALEN(m.loDBF2SQLCE.aExistsDrivers, 1)
*!*     ?m.loDBF2SQLCE.aExistsDrivers(m.lii, 1)+', '+m.loDBF2SQLCE.aExistsDrivers(m.lii, 2)
*!* NEXT

m.loDBF2SQLCE.cDB=m.lcPath+"..\out\test35.sdf"
m.loDBF2SQLCE.lDBReCreate=.T.
m.loDBF2SQLCE.cPwd=""
IF m.loDBF2SQLCE.SelectDriver("Microsoft.SQLSERVER.CE.OLEDB.3.5")
   m.loDBF2SQLCE.lReFill=.T.                             && If table in destination exists, delete data
   m.loDBF2SQLCE.AttachDataSession(1)                    && Attach default data session
   ?m.loDBF2SQLCE.UploadDBF("", m.lcAlias, "")           && Upload cursor
   ?m.loDBF2SQLCE.UploadDBF(m.lcPath+"test.dbf", "", "") && Upload table
ENDIF   


m.loDBF2SQLCE.cDB=m.lcPath+"..\out\test40.sdf"
m.loDBF2SQLCE.lDBReCreate=.T.
m.loDBF2SQLCE.cPwd=""
IF m.loDBF2SQLCE.SelectDriver("Microsoft.SQLSERVER.CE.OLEDB.4.0")
   m.loDBF2SQLCE.lReFill=.T.                             && If table in destination exists, delete data
   m.loDBF2SQLCE.AttachDataSession(1)                    && Attach default data session
   ?m.loDBF2SQLCE.UploadDBF("", m.lcAlias, "")           && Upload cursor
   ?m.loDBF2SQLCE.UploadDBF(m.lcPath+"test.dbf", "", "") && Upload table
ENDIF

DO (m.lcPath+"..\src\dbf2any") WITH "release", "sqlce"

USE IN (m.lcAlias)

