Attribute VB_Name = "modReg"
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
    "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias _
    "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal dwType As Long, ByVal lpData As String, _
    ByVal cbData As Long) As Long
            ' Return codes from Registration functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

Public Sub RegExt(ext As String)

    Dim sKeyName As String    'Holds Key Name in registry.
    Dim sKeyValue As String   'Holds Key Value in registry.
    Dim lRet As Long          'Holds error status if any from API calls.
    Dim lKey As Long          'Holds created key handle from RegCreateKey.
            
    sKeyName = "ANSIExpress"
    sKeyValue = "ANSI Express File"
    lRet = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lKey)
    lRet = RegSetValue&(lKey, "", REG_SZ, sKeyValue, 0&)
    sKeyName = ext
    sKeyValue = "ANSIExpress"
    lRet = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lKey)
    lRet = RegSetValue&(lKey, "", REG_SZ, sKeyValue, 0&)
    sKeyName = "ANSIExpress"
    sKeyValue = App.Path & "\ANSIExpress.exe %1"
    lRet = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lKey)
    lRet = RegSetValue&(lKey, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
    sKeyValue = App.Path & "\ansidoc.ico"
    lRet = RegSetValue&(lKey, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
End Sub

Public Sub RegApp()
    RegExt ".ANS"
    RegExt ".VT"
    RegExt ".ASC"
End Sub
