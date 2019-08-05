* ODBC
#DEFINE SQL_DRIVER_VER                       7

#DEFINE SQL_SUCCESS                0
#DEFINE SQL_SUCCESS_WITH_INFO      1
#DEFINE SQL_NO_DATA              100
#DEFINE SQL_ERROR                 -1
#DEFINE SQL_INVALID_HANDLE        -2
#DEFINE SQL_FETCH_NEXT             1
#DEFINE SQL_FETCH_FIRST            2
#DEFINE SQL_SUCCESS                0
#DEFINE MAX_STRING               128


#DEFINE REGAPI_Other         -1
#DEFINE REGAPI_SUCCESS       0
#DEFINE REGAPI_MORE_DATA     234  && More data is available.  
#DEFINE REGAPI_NO_MORE_ITEMS 259 && No more data is available. 


#DEFINE SECURITY_ACCESS_MASK 983103     && SAM value KEY_ALL_ACCESS

#DEFINE KEY_QUERY_VALUE         (0x0001)
#DEFINE KEY_SET_VALUE           (0x0002)
#DEFINE KEY_CREATE_SUB_KEY      (0x0004)
#DEFINE KEY_ENUMERATE_SUB_KEYS  (0x0008)
#DEFINE KEY_NOTIFY              (0x0010)
#define KEY_CREATE_LINK         (0x0020)
#DEFINE KEY_WOW64_32KEY         (0x0200)
#DEFINE KEY_WOW64_64KEY         (0x0100)

#DEFINE STANDARD_RIGHTS_READ 0x00020000

#DEFINE REG_NONE                        0 &&  No value type
#DEFINE REG_SZ                          1 &&  Unicode nul terminated string
#DEFINE REG_EXPAND_SZ                   2 &&  Unicode nul terminated string
#DEFINE REG_BINARY                      3 &&  Free form binary
#DEFINE REG_DWORD                       4 && 32-bit number
#DEFINE REG_DWORD_LITTLE_ENDIAN         4 && 32-bit number (same as REG_DWORD)
#DEFINE REG_DWORD_BIG_ENDIAN            5 && 32-bit number
#DEFINE REG_LINK                        6 && Symbolic Link (unicode)
#DEFINE REG_MULTI_SZ                    7 && Multiple Unicode strings
#DEFINE REG_RESOURCE_LIST               8 && Resource list in the resource map
#DEFINE REG_FULL_RESOURCE_DESCRIPTOR    9 && Resource list in the hardware description
#DEFINE REG_RESOURCE_REQUIREMENTS_LIST 10

#DEFINE REG_CREATED_NEW_KEY         0x00000001 && New Registry Key created
#DEFINE REG_OPENED_EXISTING_KEY     0x00000002 && Existing Key opened

#DEFINE HKEY_CLASSES_ROOT           0x80000000
#DEFINE HKEY_CURRENT_USER           0x80000001
#DEFINE HKEY_LOCAL_MACHINE          0x80000002
#DEFINE HKEY_USERS                  0x80000003
#DEFINE HKEY_PERFORMANCE_DATA       0x80000004
#DEFINE HKEY_CURRENT_CONFIG         0x80000005
#DEFINE HKEY_DYN_DATA               0x80000006
