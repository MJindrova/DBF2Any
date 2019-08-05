LOCAL m.lcPath

m.lcPath=SYS(16)
m.lcPath=IIF(RAT("\", m.lcPath)>0, LEFT(m.lcPath, RAT("\", m.lcPath)), m.lcPath)

USE (lcPath+"test")

testxxx("test", "C_GENERAL")

USE


PROCEDURE testxxx
LPARAMETERS lcAlias, lcFName

SUSP


