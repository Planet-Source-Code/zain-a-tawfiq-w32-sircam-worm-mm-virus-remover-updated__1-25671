Attribute VB_Name = "Functions"
'some declarations for the editing of the registry
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
    ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String) As Long
' Reg Key ROOT Types...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_WRITE = &H20006
Public Const KEY_ALL_ACCESS = &H2003F
Public Const KEY_READ = _
((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
' Registry location
Public Const gREGKEYLocation = "SOFTWARE\Your Company Name\Your App Name\Your Current Version"
Public Const gREGKEYXPos = "XPos"
Public Const gREGKEYYPos = "YPos"
Public Const gREGKEYWidth = "Width"
Public Const gREGKEYHeight = "Height"
Public Const gREGKEYWindowState = "WindowState"
Public Const ERROR_SUCCESS = 0&
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function EnumWindows Lib "USER32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private BBOX_DummyStr As String
Private BBOX_MySeparator As String

Function ExecWait(ByVal Cmdline As String) As String
  Dim PINF As PROCESS_INFORMATION, SUI As STARTUPINFO
  Dim ret1 As Long, ret2 As Long, ret3 As Long
  SUI.cb = Len(SUI)
  Cmdline = Cmdline + vbNullChar$ + vbNullChar$
  ret1 = CreateProcessA(0&, Cmdline, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, SUI, PINF)
  ret2 = WaitForSingleObject(PINF.hProcess, INFINITE)
  ret3 = CloseHandle(PINF.hProcess)
  ExecWait = Format$(ret1&) + " " + Format$(ret2&) + " " + Format$(ret3&)
End Function

Function Exec(ByVal Cmdline As String, Iconized As Boolean) As Long
  Cmdline = Cmdline + vbNullChar$
  If Iconized Then
    Exec = WinExec(Cmdline, 0&)
  Else
    Exec = WinExec(Cmdline, 1&)
  End If
End Function

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
  Dim sSave As String, Ret As Long
  
  Ret = GetWindowTextLength(hWnd)
  sSave = Space(Ret)
  GetWindowText hWnd, sSave, Ret + 1
  BBOX_DummyStr = BBOX_DummyStr + Format$(hWnd) + BBOX_MySeparator + sSave + vbCrLf$
  EnumWindowsProc = True
End Function


Function GetWindowList(Optional Separator As String = ":") As String
  Dim old As String
  BBOX_MySeparator = Separator
  old = BBOX_DummyStr
  BBOX_DummyStr = ""
  EnumWindows AddressOf EnumWindowsProc, ByVal 0&
  GetWindowList = BBOX_DummyStr
  BBOX_DummyStr = old
End Function





