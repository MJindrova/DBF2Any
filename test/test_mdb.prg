#INCLUDE "..\src\dbf2any.h"

LOCAL m.lcPath, m.loDBF2MDB, m.lcAlias, m.lii

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


DO (m.lcPath+"..\src\dbf2any") WITH "create", "mdb", m.loDBF2MDB
IF ISNULL(m.loDBF2MDB)
   RETURN
ENDIF

*?m.loDBF2MDB.cTemplateMDB
m.loDBF2MDB.cTargetMDB=m.lcPath+"..\out\test.mdb"   && Output file
m.loDBF2MDB.iFormat=3                               && MS Access 95/97
m.loDBF2MDB.lMDBReCreate=.T.                        && Create new mdb file for first upload
m.loDBF2MDB.lReFill=.T.                             && If table in destination exists, delete data
m.loDBF2MDB.AttachDataSession(1)                    && Attach default data session
?m.loDBF2MDB.UploadDBF("", m.lcAlias, "")           && Upload cursor
?m.loDBF2MDB.UploadDBF(m.lcPath+"test.dbf", "", "") && Upload table


*?m.loDBF2MDB.cTemplateMDB
m.loDBF2MDB.cTargetMDB=m.lcPath+"..\out\test.accdb" && Output file
m.loDBF2MDB.iFormat=4                               && MS Access 2007
m.loDBF2MDB.lMDBReCreate=.T.                        && Create new mdb file for first upload
m.loDBF2MDB.lReFill=.T.                             && If table in destination exists, delete data
m.loDBF2MDB.AttachDataSession(1)                    && Attach default data session
?m.loDBF2MDB.UploadDBF("", m.lcAlias, "")           && Upload cursor
?m.loDBF2MDB.UploadDBF(m.lcPath+"test.dbf", "", "") && Upload table


DO (m.lcPath+"..\src\dbf2any") WITH "release", "mdb"

USE IN (m.lcAlias)