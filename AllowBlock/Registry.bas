Attribute VB_Name = "Registry"
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * '
'   Written by Haytham alaa, Egypt.                                                                 '
'                                                                                                                                   '
'   1- Feel free to use this module in any application you are make     '
'       just don't remove all these comments from your application         '
'   2- Feel free to contact me if u experinced any problem                       '
'                                                                                                                                    '
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * '
' Instructions:
' You will need just three functions: QueryValue() ,SetKeyValue() and CreateNewKey()

   Public vValue As Variant
    
   Public Const REG_SZ As Long = 1
   Public Const REG_DWORD As Long = 4

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_ARENA_TRASHED = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   Public Const KEY_ALL_ACCESS = &H3F

   Public Const REG_OPTION_NON_VOLATILE = 0

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long

'###################################################
' This section is for the Drag event
    Public Declare Sub ReleaseCapture Lib "user32" ()
    
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
    Public Const WM_NCLBUTTONDOWN = &HA1
'#################################################

' Use this function to set the value of certin key in the Registry
' Example :
' Call SetKeyValue(HKEY_CURRENT_USER, "Project1", "KeyOne","HelloWorld", REG_SZ)

Public Sub SetKeyValue(hkey As Long, sKeyName As String, sValueName As String, _
   vValueSetting As Variant, lValueType As Long)
       Dim lRetVal As Long         'result of the SetValueEx function

       'open the specified key
       lRetVal = RegOpenKeyEx(hkey, sKeyName, 0, _
                              KEY_ALL_ACCESS, hkey)
       lRetVal = SetValueEx(hkey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hkey)
   End Sub

   Public Function SetValueEx(ByVal hkey As Long, sValueName As String, _
   lType As Long, vValue As Variant) As Long
       Dim lValue As Long
       Dim sValue As String
       Select Case lType
           Case REG_SZ
               sValue = vValue & Chr$(0)
               SetValueEx = RegSetValueExString(hkey, sValueName, 0&, _
                                              lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               SetValueEx = RegSetValueExLong(hkey, sValueName, 0&, _
   lType, lValue, 4)
           End Select
   End Function

   Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
   String, vValue As Variant) As Long
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Error 5

       Select Case lType
           ' For strings
           Case REG_SZ:
               sValue = String(cch, 0)

   lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
   sValue, cch)
               If lrc = ERROR_NONE Then
                   vValue = Left$(sValue, cch - 1)
               Else
                   vValue = Empty
               End If
           ' For DWORDS
           Case REG_DWORD:
   lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
   lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           Case Else
               'all other data types not supported
               lrc = -1
       End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit
   End Function

' use this function to return the value of a certain key

Public Function QueryValue(hkey As Long, sKeyName As String, sValueName As String)
    Dim lRetVal As Long         'result of the API functions
    
    lRetVal = RegOpenKeyEx(hkey, sKeyName, 0, _
KEY_ALL_ACCESS, hkey)
    lRetVal = QueryValueEx(hkey, sValueName, vValue)
    RegCloseKey (hkey)
    
    QueryValue = vValue
End Function

' Use this function to create new key

Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
              vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
              0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Sub
