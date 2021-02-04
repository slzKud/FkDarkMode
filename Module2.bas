Attribute VB_Name = "Module2"
Option Explicit


'Download by http://www.NewXing.com
'*************************************************************************
'**模 块 名：RegWork
'**创 建 人：叶帆
'**日    期：2003年01月11日
'**修 改 人：
'**日    期：
'**描    述：注册表操作(不同类型,读写方法有一定区别)
'**版    本：版本1.0
'*************************************************************************+
'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------

'关闭登录关键字
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hkey As Long) As Long

'建立关键字
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long

'打开关键字
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long

'返回关键字的类型和值
Private Declare Function RegQueryValueEx_SZ Lib "advapi32" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_DWORD Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, ByRef lpcbData As Long) As Long

'将文本字符串与指定关键字关联
Private Declare Function RegSetValueEx_SZ Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_DWORD Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

'删除关键字
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hkey As Long, ByVal pszSubKey As String) As Long
'从登录关键字中删除一个值
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long

' 注册表的数据类型
Private Enum REGValueType

    REG_SZ = 1                             ' Unicode空终结字符串
    REG_EXPAND_SZ = 2                      ' Unicode空终结字符串
    REG_BINARY = 3                         ' 二进制数值
    REG_DWORD = 4                          ' 32-bit 数字
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7                       ' 二进制数值串

End Enum

' 注册表创建类型值...
Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留
Const KEY_WOW64_64KEY = &H100

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
      KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
      KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' 注册表关键字根类型...
Private Enum REGRoot

    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004

End Enum

' 返回值...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'- 注册表安全属性类型...
Private Type SECURITY_ATTRIBUTES

    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean

End Type

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByRef lpftLastWriteTime As FILETIME) As Long
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


'*************************************************************************
'**函 数 名：WriteRegKey
'**输    入：ByVal KeyRoot(REGRoot)         - 根
'**        ：ByVal KeyName(String)          - 键的路径
'**        ：ByVal SubKeyName(String)       - 键名
'**        ：ByVal SubKeyType(REGValueType) - 键的类型
'**        ：ByVal SubKeyValue(String)      - 键值
'**输    出：(Boolean) - 成功返回True，失败返回False
'**功能描述：写注册表
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2003年01月10日
'**修 改 人：
'**日    期：
'**版    本：版本1.0
'*************************************************************************

Private Function WriteRegKey(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, ByVal SubKeyType As REGValueType, ByVal SubKeyValue As String) As Boolean

    Dim rc As Long                                      ' 返回代码
    Dim hkey As Long                                    ' 处理一个注册表关键字
    Dim hDepth As Long                                  ' 用于装载下列某个常数的一个变量
    ' REG_CREATED_NEW_KEY——新建的一个子项
    ' REG_OPENED_EXISTING_KEY——打开一个现有的项
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 注册表安全类型
    Dim i As Integer
    Dim bytValue(1024) As Byte

    lpAttr.nLength = 50                                 ' 设置安全属性为缺省值...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '- 创建/打开注册表关键字...
    rc = RegCreateKeyEx(KeyRoot, KeyName, 0, SubKeyType, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, lpAttr, hkey, hDepth)                                                                                          ' 创建/打开//KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理...

    '- 创建/修改关键字值...

    If (SubKeyValue = "") Then SubKeyValue = " "        ' 要让RegSetValueEx() 工作需要输入一个空格...

    Select Case SubKeyType                                        ' 搜索数据类型...

        Case REG_SZ, REG_EXPAND_SZ                                ' 字符串注册表关键字数据类型

        rc = RegSetValueEx_SZ(hkey, SubKeyName, 0, SubKeyType, ByVal SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError         ' 错误处理

        Case REG_DWORD                                            ' 四字节注册表关键字数据类型

        rc = RegSetValueEx_DWORD(hkey, SubKeyName, 0, SubKeyType, Val("&h" + SubKeyValue), 4)

        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError         ' 错误处理

        Case REG_BINARY                                           ' 二进制字符串

        Dim intNum As Integer

        For i = 1 To Len(Trim$(SubKeyValue)) - 1 Step 3

            intNum = intNum + 1
            bytValue(intNum - 1) = Val("&h" + Mid$(SubKeyValue, i, 2))

        Next i

        rc = RegSetValueEx_BINARY(hkey, SubKeyName, 0, SubKeyType, bytValue(0), intNum)

        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理

        Case Else

        GoTo CreateKeyError                                    ' 错误处理

    End Select

    '- 关闭注册表关键字...
    rc = RegCloseKey(hkey)                              ' 关闭关键字

    WriteRegKey = True                                  ' 返回成功

    Exit Function                                       ' 退出

CreateKeyError:

    WriteRegKey = False                                 ' 设置错误返回代码
    rc = RegCloseKey(hkey)                              ' 试图关闭关键字

End Function

'*************************************************************************
'**函 数 名：ReadRegKey
'**输    入：KeyRoot(Long)     - 根
'**        ：KeyName(String)   - 键的路径
'**        ：SubKeyRef(String) - 键名
'**输    出：(String) - 返回键值
'**功能描述：读注册表
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2003年01月10日
'**修 改 人：
'**日    期：
'**版    本：版本1.0
'*************************************************************************

Private Function ReadRegKey(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, Optional flag64 As Boolean = True) As String

    Dim i As Long                                            ' 循环计数器
    Dim rc As Long                                           ' 返回代码
    Dim hkey As Long                                         ' 处理打开的注册表关键字
    Dim hDepth As Long                                       '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                  ' 注册表关键字数据类型
    Dim tmpVal As String                                     ' 注册表关键字的临时存储器
    Dim KeyValSize As Long                                   ' 注册表关键字变量尺寸
    Dim lngValue As Long
    Dim bytValue(1024) As Byte

    ' 在 KeyRoot下打开注册表关键字
    If flag64 Then
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, hkey)    ' 打开注册表关键字(64注册读取)
    Else
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, hkey)    ' 打开注册表关键字
    End If
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError                 ' 处理错误...

    ' 检测键的类型

    rc = RegQueryValueEx(hkey, SubKeyName, 0, lKeyValType, ByVal 0, KeyValSize)  ' 获得/创建关键字的值lKeyValType

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError                 ' 处理错误...

    '读相应的键值

    Select Case lKeyValType                                         ' 搜索数据类型...

        Case REG_SZ, REG_EXPAND_SZ                                  ' 字符串注册表关键字数据类型

        tmpVal = String$(1024, 0)                                   ' 分配变量空间
        KeyValSize = 1024                                           ' 标记变量尺寸

        rc = RegQueryValueEx_SZ(hkey, SubKeyName, 0, 0, tmpVal, KeyValSize)     ' 获得/创建关键字的值
        
        If rc <> ERROR_SUCCESS Then GoTo GetKeyError           ' 错误处理

        If InStr(tmpVal, Chr$(0)) > 0 Then sKeyVal = Left$(tmpVal, InStr(tmpVal, Chr$(0)) - 1)     ' 复制字符串的值,并去除空字符.
        
        Case REG_DWORD                                             ' 四字节注册表关键字数据类型
        
        KeyValSize = 1024                                          ' 标记变量尺寸
        rc = RegQueryValueEx_DWORD(hkey, SubKeyName, 0, 0, lngValue, KeyValSize)     ' 获得/创建关键字的值
        
        If rc <> ERROR_SUCCESS Then GoTo GetKeyError            ' 错误处理
        
        sKeyVal = "0x" + Hex$(lngValue)
        
        Case REG_BINARY                                            ' 二进制字符串
        
        rc = RegQueryValueEx(hkey, SubKeyName, 0, 0, bytValue(0), KeyValSize)       ' 获得/创建关键字的值

        If rc <> ERROR_SUCCESS Then GoTo GetKeyError            ' 错误处理

        sKeyVal = ""
        
        For i = 1 To KeyValSize

            If Len(Hex$(bytValue(i - 1))) = 1 Then
            
                sKeyVal = sKeyVal + "0" + Hex$(bytValue(i - 1)) + " "
                
            Else
            
                sKeyVal = sKeyVal + Hex$(bytValue(i - 1)) + " "
                
            End If

        Next i

        Case Else
        
        sKeyVal = ""
    
    End Select

    ReadRegKey = sKeyVal                                      ' 返回值
    rc = RegCloseKey(hkey)                                    ' 关闭注册表关键字
    
    Exit Function                                             ' 退出

GetKeyError:

    ' 错误发生过后进行清除...

    ReadRegKey = ""                                      ' 设置返回值为错误

    rc = RegCloseKey(hkey)                                    ' 关闭注册表关键字

End Function

'*************************************************************************
'**函 数 名：DelRegKey
'**输    入：KeyRoot(Long)     - 根
'**        ：KeyName(String)   - 键的路径
'**        ：SubKeyRef(String) - 键名
'**输    出：(Long) - 状态码
'**功能描述：删除关键字
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2003年01月11日
'**修 改 人：
'**日    期：
'**版    本：版本1.0
'*************************************************************************

Private Function DelRegKey(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, Optional flag64 As Boolean = False) As Long

    Dim lKeyId          As Long
    Dim lResult         As Long
    'dbglog KeyName + "\" + SubKeyName, "DelRegKey"
    '检测设置的参数
    If Len(KeyName) = 0 And Len(SubKeyName) = 0 Then

        ' 键值没设置则返回相应错误码
        DelRegKey = ERROR_BADKEY

        Exit Function

    End If
    ' 打开关键字并尝试创建它,如果已存在,则返回ID值
    'lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    If flag64 Then
        lResult = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, lKeyId)
    Else
        lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    End If
    If lResult = 0 Then

        '删除关键字
        DelRegKey = RegDeleteKey(lKeyId, ByVal SubKeyName)
        'dbglog "OK", "DelRegKey"
    End If

End Function

'*************************************************************************
'**函 数 名：DelRegValue
'**输    入：KeyRoot(Long)     - 根
'**        ：KeyName(String)   - 键的路径
'**        ：SubKeyRef(String) - 键名
'**输    出：(Long) - 状态码
'**功能描述：从登录关键字中删除一个值
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2003年01月11日
'**修 改 人：
'**日    期：
'**版    本：版本1.0
'*************************************************************************

Private Function DelRegValue(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, Optional flag64 As Boolean = False) As Long

    Dim lKeyId As Long
    Dim lResult As Long

    '检测设置的参数
    If Len(KeyName) = 0 And Len(SubKeyName) = 0 Then

        ' 键值没设置则返回相应错误码
        DelRegValue = ERROR_BADKEY

        Exit Function

    End If

    ' 打开关键字并尝试创建它,如果已存在,则返回ID值
    'lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    If flag64 Then
        lResult = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, lKeyId)
    Else
        lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    End If
    If lResult = 0 Then

        '从登录关键字中删除一个值
        DelRegValue = RegDeleteValue(lKeyId, ByVal SubKeyName)

    End If

End Function


Public Function GetNowToolsVersion() As String
Dim temp As String
temp = ReadRegKey(HKEY_LOCAL_MACHINE, "SOFTWARE\livlab\livlab_tools", "ToolsVersion")
If temp = "" Then GetNowToolsVersion = "1.0": Exit Function
'dbglog "当前工具版本:" + temp, "GetNowToolsVersion"
GetNowToolsVersion = temp
End Function

Private Function FindKeys(ByVal hkey As REGRoot, SubKey As String) As String()
  Dim phkRet     As Long
  Dim Index     As Long, Name       As String, lName       As Long, lReserved       As Long, Class       As String, lClass       As Long, LWT       As FILETIME
  Dim lRet     As Long
  Dim Keys     As String, TempKeys       As String
  Dim temp_arr() As String
  Static Num     As Long
  lReserved = 0&
  Index = 0
  ReDim temp_arr(0)
  lRet = RegOpenKeyEx(hkey, SubKey, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, phkRet)
  'lRet = RegOpenKey(hKey, SubKey, phkRet)
  If lRet = ERROR_SUCCESS Then
    Do
        DoEvents
        Name = String(255, Chr(0))
        lName = Len(Name)
        lRet = RegEnumKeyEx(phkRet, Index, Name, lName, lReserved, Class, lClass, LWT)
        If lRet = ERROR_SUCCESS Then
          'If   SubKey   =   ""   Then
            'Keys   =   Name
          'Else
            Keys = SubKey & "\" & Name
          'End   If
          'TempKeys   =  Keys
          'dbglog Keys, "FindKeys"
          ReDim Preserve temp_arr(Index)
          temp_arr(Index) = Replace(Name, Chr(0), "")
        Else
          Exit Do
        End If
        Index = Index + 1
        
        Loop While lRet = ERROR_SUCCESS
  End If
   
  Call RegCloseKey(phkRet)
  FindKeys = temp_arr
  End Function


Public Function get_path_from_reg(modeltype As String, modelfrom As String)
Dim reg As String, path_reg As String
Select Case modelfrom
    Case "livlab"
        path_reg = ReadRegKey(HKEY_LOCAL_MACHINE, "SOFTWARE\livlab\livlab_tools", "ToolsPath")
    Case "p5d"
        path_reg = ReadRegKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Lockheed Martin\Prepar3D v5", "SetupPath")
    Case "p3d"
        path_reg = ReadRegKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Lockheed Martin\Prepar3D v4", "SetupPath")
    Case "winbuild"
        path_reg = ReadRegKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild")
    Case "updatechannel"
        path_reg = ReadRegKey(HKEY_LOCAL_MACHINE, "SOFTWARE\livlab\livlab_tools", "UpdateChannel")
        If Trim(path_reg) = "" Then path_reg = "ERROR"
        On Error Resume Next
        Dim p As Integer
        p = -1
        p = CInt(path_reg)
        If p = -1 Or p > 3 Then
            #If BETA = 1 Then
                p = 2
            #Else
                p = 1
            #End If
        End If
        path_reg = CStr(p)
    Case Else
        get_path_from_reg = "ERROR"
        Exit Function
End Select
'dbglog "获取了Path:" + path_reg + "," + modeltype + "," + modelfrom, "get_path_from_reg"
get_path_from_reg = path_reg
End Function
