#IFNDEF VFP_VERSION
 #DEFINE VFP_VERSION VAL(SUBS(VERSION(),LEN("Visual FoxPro ")+1,2))
#ENDIF

#DEFINE __DBF2ANY_Mode_ODBC                0
#DEFINE __DBF2ANY_Mode_OLEDB               1

#DEFINE __DBF2ANY_ErrOK                0 && 
#DEFINE __DBF2ANY_ErrConnectionFailed -1 && Can't open ODBC/OLEDB connection
#DEFINE __DBF2ANY_ErrFewParameters    -3 && A few parameters
#DEFINE __DBF2ANY_ErrNoTables         -4 && Can't get table list
#DEFINE __DBF2ANY_ErrCannotCTable     -5 && Can't create table
#DEFINE __DBF2ANY_ErrCannotDTable     -6 && Can't drop table
#DEFINE __DBF2ANY_ErrNotaTable        -7 && Table or view is system table
#DEFINE __DBF2ANY_ErrNotInsert        -8 && Insert record failed
#DEFINE __DBF2ANY_ErrNotDelete        -9 && Delete record failed
#DEFINE __DBF2ANY_ErrAliasNotOpen    -10 && Alias not is open
#DEFINE __DBF2ANY_ErrDBFNotOpen      -11 && Table is not open
#DEFINE __DBF2ANY_ErrDBCNotOpen      -12 && DBC is not open

#DEFINE __DBF2ANY_ErrDBNotCreated    -102 && Can't create MDB/SDF file



