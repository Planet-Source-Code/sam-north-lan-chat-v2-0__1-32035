Attribute VB_Name = "ModDecs"
'Flashing window decleration
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long

'My code to write a key to Run key in the regisrty.
'Function definitions come from www.vbapi.com - very useful site!

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal _
    hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired _
    As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" _
    () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" _
    (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
    
'Always on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Always on top constants
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Const KEY_READ = &H20019
Const REG_SZ = 1
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_DYN_DATA = &H80000006
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_USERS = &H80000003
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_DWORD_BIG_ENDIAN = 5
Const REG_DWORD_LITTLE_ENDIAN = 4
Const REG_EXPAND_SZ = 2
Const REG_LINK = 6
Const REG_MULTI_SZ = 7
Const REG_NONE = 0
Const REG_RESOURCE_LIST = 8
Const KEY_WRITE = &H20006

Public Sub WriteReg() 'writes to the key below to make sure it runs on start up
'no error checking, no point becuase if it fails you cant return it to the
'client because UDP is one way! + I can't be bothered!

On Error Resume Next
  Dim hregkey As Long
  Dim secattr As SECURITY_ATTRIBUTES
  Dim subkey As String
  Dim neworused As Long
  Dim stringbuffer As String
  Dim retval As Long

  subkey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
  secattr.nLength = Len(secattr)
  secattr.lpSecurityDescriptor = 0
  secattr.bInheritHandle = True


  retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, "", 0, KEY_WRITE, secattr, hregkey, neworused)
  If retval <> 0 Then
    Exit Sub
  End If


  stringbuffer = App.Path & "\" & App.EXEName & ".exe" & vbNullChar
  retval = RegSetValueEx(hregkey, "LanChat", 0, REG_SZ, ByVal stringbuffer, Len(stringbuffer))
  
  retval = RegCloseKey(hregkey)
End Sub

'make a form normal
Public Sub MakeNormal(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

'make a form topmost
Public Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
