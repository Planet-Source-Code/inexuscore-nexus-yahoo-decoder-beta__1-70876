Attribute VB_Name = "mod_Global"
Option Explicit

'// public win32 api declarations
Public Declare Function InitCommonControls Lib "ComCtl32.dll" () As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

'// public constant declarations
Public Const SW_SHOWNORMAL = 1  '// for shell execute api
Public Const InstanceCode = "A8F500EA:D54F:210F:ED0A:F4A5A20C037B" '// App Instance Code
'// for SHBrowseForFolder api
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const MAX_PATH = 260
'// constants for text coloring
Public Const lngColorSent As Long = &H808080
Public Const lngColorRecv As Long = &H8000000D

'// public data types declaration (for SHBrowseForFolder api)
Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'// public variables declarations
Public bLoading As Boolean  '// decoding status
Public strFolder As String  '// profiles root path
Public strNodeData() As String  '// selected node's data (profile, buddy, message)
Public strOutput As String      '// decoded archive data
