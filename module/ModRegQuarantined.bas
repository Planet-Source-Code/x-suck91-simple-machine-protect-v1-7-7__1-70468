Attribute VB_Name = "ModRegQuarantined"
'added By X-Suck91

Const KEY_ALL_ACCESS = &H3F
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const REG_PRIMARY_KEY = "Software\Classes\"
Const REG_SHELL_KEY = "Shell\"
Const REG_SHELL_OPEN_KEY = "Open\"
Const REG_SHELL_OPEN_COMMAND_KEY = "Command"
Const REG_ICON_KEY = "DefaultIcon"
Const REG_SZ = 1
Const REG_OPTION_NON_VOLATILE = 0
Const ERROR_SUCCESS = 0&
Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As Any, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long


Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
Private Function OpenKey(lhKey As Long, SubKey As String, ulOptions As Long) As Long
Dim lhKeyOpen As Long
Dim lResult As Long

lhKeyOpen = 0
lResult = RegOpenKeyEx(lhKey, SubKey, 0, ulOptions, lhKeyOpen)

If lResult <> ERROR_SUCCESS Then
OpenKey = 0
Else
OpenKey = lhKeyOpen
End If
End Function

Private Function CreateKey(lhKey As Long, SubKey As String, NewSubKey As String) As Boolean
Dim lhKeyOpen As Long
Dim lhKeyNew As Long
Dim lDisposition As Long
Dim lResult As Long
Dim Security As SECURITY_ATTRIBUTES

lhKeyOpen = OpenKey(lhKey, SubKey, KEY_CREATE_SUB_KEY)
lResult = RegCreateKeyEx(lhKeyOpen, NewSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, lhKeyNew, lDisposition)

If lResult = ERROR_SUCCESS Then
CreateKey = True
RegCloseKey (lhKeyNew)
Else
CreateKey = False
End If

RegCloseKey (lhKeyOpen)
End Function

Private Function SetValue(lhKey As Long, SubKey As String, sValue As String) As Boolean
Dim lhKeyOpen As Long
Dim lResult As Long
Dim lTyp As Long
Dim lByte As Long

lByte = Len(sValue)
lTyp = REG_SZ
lhKeyOpen = OpenKey(lhKey, SubKey, KEY_SET_VALUE)
lResult = RegSetValue(lhKey, SubKey, lTyp, sValue, lByte)

If lResult <> ERROR_SUCCESS Then
SetValue = False
Else
SetValue = True
RegCloseKey (lhKeyOpen)
End If
End Function

Public Function RegisterFile(sFileExt As String, sFileDescr As String, sAppID As String, sOpenCmd As String, sIconFile As String) As Boolean
Dim hKey As Long
Dim bSuccess As Boolean
Dim bSuccess2 As Boolean
    
bSuccess = False
hKey = HKEY_LOCAL_MACHINE
  
If CreateKey(hKey, REG_PRIMARY_KEY, sFileExt) Then
 If SetValue(hKey, REG_PRIMARY_KEY & sFileExt, sAppID) Then
  If CreateKey(hKey, REG_PRIMARY_KEY, sAppID) Then
   If SetValue(hKey, REG_PRIMARY_KEY & sAppID, sFileDescr) Then
    If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY) Then
        bSuccess = SetValue(hKey, REG_PRIMARY_KEY & sAppID & "\" & REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY, sOpenCmd)
     If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, REG_ICON_KEY) Then
        bSuccess2 = SetValue(hKey, REG_PRIMARY_KEY & sAppID & "\" & REG_ICON_KEY, sIconFile)
     End If
    End If
   End If
  End If
 End If
End If

RegisterFile = (bSuccess = bSuccess2)
End Function
