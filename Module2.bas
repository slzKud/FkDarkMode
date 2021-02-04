Attribute VB_Name = "Module2"
Option Explicit


'Download by http://www.NewXing.com
'*************************************************************************
'**ģ �� ����RegWork
'**�� �� �ˣ�Ҷ��
'**��    �ڣ�2003��01��11��
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����ע������(��ͬ����,��д������һ������)
'**��    �����汾1.0
'*************************************************************************+
'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------

'�رյ�¼�ؼ���
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hkey As Long) As Long

'�����ؼ���
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long

'�򿪹ؼ���
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long

'���عؼ��ֵ����ͺ�ֵ
Private Declare Function RegQueryValueEx_SZ Lib "advapi32" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_DWORD Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, ByRef lpcbData As Long) As Long

'���ı��ַ�����ָ���ؼ��ֹ���
Private Declare Function RegSetValueEx_SZ Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_DWORD Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

'ɾ���ؼ���
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" (ByVal hkey As Long, ByVal pszSubKey As String) As Long
'�ӵ�¼�ؼ�����ɾ��һ��ֵ
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long

' ע������������
Private Enum REGValueType

    REG_SZ = 1                             ' Unicode���ս��ַ���
    REG_EXPAND_SZ = 2                      ' Unicode���ս��ַ���
    REG_BINARY = 3                         ' ��������ֵ
    REG_DWORD = 4                          ' 32-bit ����
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7                       ' ��������ֵ��

End Enum

' ע���������ֵ...
Const REG_OPTION_NON_VOLATILE = 0       ' ��ϵͳ��������ʱ���ؼ��ֱ�����
Const KEY_WOW64_64KEY = &H100

' ע���ؼ��ְ�ȫѡ��...
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

' ע���ؼ��ָ�����...
Private Enum REGRoot

    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004

End Enum

' ����ֵ...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'- ע���ȫ��������...
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
'**�� �� ����WriteRegKey
'**��    �룺ByVal KeyRoot(REGRoot)         - ��
'**        ��ByVal KeyName(String)          - ����·��
'**        ��ByVal SubKeyName(String)       - ����
'**        ��ByVal SubKeyType(REGValueType) - ��������
'**        ��ByVal SubKeyValue(String)      - ��ֵ
'**��    ����(Boolean) - �ɹ�����True��ʧ�ܷ���False
'**����������дע���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ҷ��
'**��    �ڣ�2003��01��10��
'**�� �� �ˣ�
'**��    �ڣ�
'**��    �����汾1.0
'*************************************************************************

Private Function WriteRegKey(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, ByVal SubKeyType As REGValueType, ByVal SubKeyValue As String) As Boolean

    Dim rc As Long                                      ' ���ش���
    Dim hkey As Long                                    ' ����һ��ע���ؼ���
    Dim hDepth As Long                                  ' ����װ������ĳ��������һ������
    ' REG_CREATED_NEW_KEY�����½���һ������
    ' REG_OPENED_EXISTING_KEY������һ�����е���
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ע���ȫ����
    Dim i As Integer
    Dim bytValue(1024) As Byte

    lpAttr.nLength = 50                                 ' ���ð�ȫ����Ϊȱʡֵ...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '- ����/��ע���ؼ���...
    rc = RegCreateKeyEx(KeyRoot, KeyName, 0, SubKeyType, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, lpAttr, hkey, hDepth)                                                                                          ' ����/��//KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������...

    '- ����/�޸Ĺؼ���ֵ...

    If (SubKeyValue = "") Then SubKeyValue = " "        ' Ҫ��RegSetValueEx() ������Ҫ����һ���ո�...

    Select Case SubKeyType                                        ' ������������...

        Case REG_SZ, REG_EXPAND_SZ                                ' �ַ���ע���ؼ�����������

        rc = RegSetValueEx_SZ(hkey, SubKeyName, 0, SubKeyType, ByVal SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError         ' ������

        Case REG_DWORD                                            ' ���ֽ�ע���ؼ�����������

        rc = RegSetValueEx_DWORD(hkey, SubKeyName, 0, SubKeyType, Val("&h" + SubKeyValue), 4)

        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError         ' ������

        Case REG_BINARY                                           ' �������ַ���

        Dim intNum As Integer

        For i = 1 To Len(Trim$(SubKeyValue)) - 1 Step 3

            intNum = intNum + 1
            bytValue(intNum - 1) = Val("&h" + Mid$(SubKeyValue, i, 2))

        Next i

        rc = RegSetValueEx_BINARY(hkey, SubKeyName, 0, SubKeyType, bytValue(0), intNum)

        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������

        Case Else

        GoTo CreateKeyError                                    ' ������

    End Select

    '- �ر�ע���ؼ���...
    rc = RegCloseKey(hkey)                              ' �رչؼ���

    WriteRegKey = True                                  ' ���سɹ�

    Exit Function                                       ' �˳�

CreateKeyError:

    WriteRegKey = False                                 ' ���ô��󷵻ش���
    rc = RegCloseKey(hkey)                              ' ��ͼ�رչؼ���

End Function

'*************************************************************************
'**�� �� ����ReadRegKey
'**��    �룺KeyRoot(Long)     - ��
'**        ��KeyName(String)   - ����·��
'**        ��SubKeyRef(String) - ����
'**��    ����(String) - ���ؼ�ֵ
'**������������ע���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ҷ��
'**��    �ڣ�2003��01��10��
'**�� �� �ˣ�
'**��    �ڣ�
'**��    �����汾1.0
'*************************************************************************

Private Function ReadRegKey(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, Optional flag64 As Boolean = True) As String

    Dim i As Long                                            ' ѭ��������
    Dim rc As Long                                           ' ���ش���
    Dim hkey As Long                                         ' ����򿪵�ע���ؼ���
    Dim hDepth As Long                                       '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                  ' ע���ؼ�����������
    Dim tmpVal As String                                     ' ע���ؼ��ֵ���ʱ�洢��
    Dim KeyValSize As Long                                   ' ע���ؼ��ֱ����ߴ�
    Dim lngValue As Long
    Dim bytValue(1024) As Byte

    ' �� KeyRoot�´�ע���ؼ���
    If flag64 Then
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, hkey)    ' ��ע���ؼ���(64ע���ȡ)
    Else
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, hkey)    ' ��ע���ؼ���
    End If
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError                 ' �������...

    ' ����������

    rc = RegQueryValueEx(hkey, SubKeyName, 0, lKeyValType, ByVal 0, KeyValSize)  ' ���/�����ؼ��ֵ�ֵlKeyValType

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError                 ' �������...

    '����Ӧ�ļ�ֵ

    Select Case lKeyValType                                         ' ������������...

        Case REG_SZ, REG_EXPAND_SZ                                  ' �ַ���ע���ؼ�����������

        tmpVal = String$(1024, 0)                                   ' ��������ռ�
        KeyValSize = 1024                                           ' ��Ǳ����ߴ�

        rc = RegQueryValueEx_SZ(hkey, SubKeyName, 0, 0, tmpVal, KeyValSize)     ' ���/�����ؼ��ֵ�ֵ
        
        If rc <> ERROR_SUCCESS Then GoTo GetKeyError           ' ������

        If InStr(tmpVal, Chr$(0)) > 0 Then sKeyVal = Left$(tmpVal, InStr(tmpVal, Chr$(0)) - 1)     ' �����ַ�����ֵ,��ȥ�����ַ�.
        
        Case REG_DWORD                                             ' ���ֽ�ע���ؼ�����������
        
        KeyValSize = 1024                                          ' ��Ǳ����ߴ�
        rc = RegQueryValueEx_DWORD(hkey, SubKeyName, 0, 0, lngValue, KeyValSize)     ' ���/�����ؼ��ֵ�ֵ
        
        If rc <> ERROR_SUCCESS Then GoTo GetKeyError            ' ������
        
        sKeyVal = "0x" + Hex$(lngValue)
        
        Case REG_BINARY                                            ' �������ַ���
        
        rc = RegQueryValueEx(hkey, SubKeyName, 0, 0, bytValue(0), KeyValSize)       ' ���/�����ؼ��ֵ�ֵ

        If rc <> ERROR_SUCCESS Then GoTo GetKeyError            ' ������

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

    ReadRegKey = sKeyVal                                      ' ����ֵ
    rc = RegCloseKey(hkey)                                    ' �ر�ע���ؼ���
    
    Exit Function                                             ' �˳�

GetKeyError:

    ' ����������������...

    ReadRegKey = ""                                      ' ���÷���ֵΪ����

    rc = RegCloseKey(hkey)                                    ' �ر�ע���ؼ���

End Function

'*************************************************************************
'**�� �� ����DelRegKey
'**��    �룺KeyRoot(Long)     - ��
'**        ��KeyName(String)   - ����·��
'**        ��SubKeyRef(String) - ����
'**��    ����(Long) - ״̬��
'**����������ɾ���ؼ���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ҷ��
'**��    �ڣ�2003��01��11��
'**�� �� �ˣ�
'**��    �ڣ�
'**��    �����汾1.0
'*************************************************************************

Private Function DelRegKey(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, Optional flag64 As Boolean = False) As Long

    Dim lKeyId          As Long
    Dim lResult         As Long
    'dbglog KeyName + "\" + SubKeyName, "DelRegKey"
    '������õĲ���
    If Len(KeyName) = 0 And Len(SubKeyName) = 0 Then

        ' ��ֵû�����򷵻���Ӧ������
        DelRegKey = ERROR_BADKEY

        Exit Function

    End If
    ' �򿪹ؼ��ֲ����Դ�����,����Ѵ���,�򷵻�IDֵ
    'lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    If flag64 Then
        lResult = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, lKeyId)
    Else
        lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    End If
    If lResult = 0 Then

        'ɾ���ؼ���
        DelRegKey = RegDeleteKey(lKeyId, ByVal SubKeyName)
        'dbglog "OK", "DelRegKey"
    End If

End Function

'*************************************************************************
'**�� �� ����DelRegValue
'**��    �룺KeyRoot(Long)     - ��
'**        ��KeyName(String)   - ����·��
'**        ��SubKeyRef(String) - ����
'**��    ����(Long) - ״̬��
'**�����������ӵ�¼�ؼ�����ɾ��һ��ֵ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ҷ��
'**��    �ڣ�2003��01��11��
'**�� �� �ˣ�
'**��    �ڣ�
'**��    �����汾1.0
'*************************************************************************

Private Function DelRegValue(ByVal KeyRoot As REGRoot, ByVal KeyName As String, ByVal SubKeyName As String, Optional flag64 As Boolean = False) As Long

    Dim lKeyId As Long
    Dim lResult As Long

    '������õĲ���
    If Len(KeyName) = 0 And Len(SubKeyName) = 0 Then

        ' ��ֵû�����򷵻���Ӧ������
        DelRegValue = ERROR_BADKEY

        Exit Function

    End If

    ' �򿪹ؼ��ֲ����Դ�����,����Ѵ���,�򷵻�IDֵ
    'lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    If flag64 Then
        lResult = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, lKeyId)
    Else
        lResult = RegCreateKey(KeyRoot, KeyName, lKeyId)
    End If
    If lResult = 0 Then

        '�ӵ�¼�ؼ�����ɾ��һ��ֵ
        DelRegValue = RegDeleteValue(lKeyId, ByVal SubKeyName)

    End If

End Function


Public Function GetNowToolsVersion() As String
Dim temp As String
temp = ReadRegKey(HKEY_LOCAL_MACHINE, "SOFTWARE\livlab\livlab_tools", "ToolsVersion")
If temp = "" Then GetNowToolsVersion = "1.0": Exit Function
'dbglog "��ǰ���߰汾:" + temp, "GetNowToolsVersion"
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
'dbglog "��ȡ��Path:" + path_reg + "," + modeltype + "," + modelfrom, "get_path_from_reg"
get_path_from_reg = path_reg
End Function
