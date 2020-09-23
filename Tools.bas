Attribute VB_Name = "Tools"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessageLong& Lib "User32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "User32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal HWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SystemParametersInfo_Rect Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal HWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function OffsetRect Lib "User32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Type PointAPI
X As Long
Y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const WM_MOVING = &H216
Private Const WM_SIZING = &H214
Private Const WM_ENTERSIZEMOVE = &H231
Private Const WM_EXITSIZEMOVE = &H232
Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const SPI_GETWORKAREA = 48
Private Const WMSZ_LEFT = 1
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPRIGHT = 5
Private Enum SnapFormMode
Moving = 1
Sizing = 2
End Enum
Private Type DockingLog
HWnd As Long
oldProc As Long
End Type
Private m_hMasterWnd As Long
Private Logs() As DockingLog, LogCount As Integer, MaxLogs As Integer
Private MouseX As Long, MouseY As Long
Public SnappedX As Boolean, SnappedY As Boolean
Public Rects() As RECT
Private Const SnapWidth = 10
Private Const DoSubClass As Boolean = True
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_ALIAS = &H10000
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ALIAS_START = 0
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_VALID = &H1F
Private Const SND_NOWAIT = &H2000
Private Const SND_VALIDFLAGS = &H17201F
Private Const SND_RESERVED = &HFF000000
Private Const SND_TYPE_MASK = &H170007
Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1
Const REG_DWORD = 4
Public Enum HKeyTypes
HKEY_CLASSES_ROOT = &H80000000
HKEY_CURRENT_USER = &H80000001
HKEY_LOCAL_MACHINE = &H80000002
HKEY_USERS = &H80000003
HKEY_PERFORMANCE_DATA = &H80000004
End Enum
Private m_snd() As Byte
Public Ren, From, Who, Too, Message, Imv, Name As String
Public Lp, Lc As Integer
Public User(1 To 100) As String
Public Pass(1 To 100) As String
Public Temp(1 To 100) As String
Public Data(1 To 100) As String
Public LCount As String
Public Const Port = 80
'Removes a name/item from list
'Example:
'Private Sub Command1_Click()
'RemoveName Text1.text, List1
'End Sub
Function RemoveName(Name, List As ListBox)
Dim X As Integer
For X = 0 To List.ListCount - 1
If Name = List.List(X) Then List.RemoveItem (X)
Next
End Function
Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strData As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(hKey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
r = RegCloseKey(keyhand)
End Sub


Public Function GetString(hKey As HKeyTypes, strPath As String, strValue As String)
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim lValueType As Long
r = RegOpenKey(hKey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
strBuf = String(lDataBufSize, " ")
lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
intZeroPos = InStr(strBuf, Chr$(0))
If intZeroPos > 0 Then
GetString = Left$(strBuf, intZeroPos - 1)
Else
GetString = strBuf
End If
End If
End If
End Function



Private Function GetFrmRects(ByVal HWnd As Long)
Dim frm As Form
Dim I As Integer
ReDim Rects(0 To 0)
SystemParametersInfo_Rect SPI_GETWORKAREA, vbNull, Rects(0), 0
I = 1
For Each frm In Forms
If frm.Visible And Not frm.HWnd = HWnd Then
ReDim Preserve Rects(0 To I)
GetWindowRect frm.HWnd, Rects(I)
I = I + 1
End If
Next frm
End Function


'This changes your Yahoo status.
'Example:
'Form1_Load()
'ChangeStatus "Program 1.0 Loaded'

Sub ChangeStatus(Status As String)
Dim Name As String
Name = GetString(HKEY_CURRENT_USER, "Software\Yahoo\Pager", "Yahoo! user id")
Call SaveString(HKEY_CURRENT_USER, "Software\Yahoo\Pager\profiles\" + Name + "\Custom Msgs", 1, Status)
Dim X As Long
On Error Resume Next
X = FindWindow("YahooBuddyMain", vbNullString)
SendMessageLong X, &H111, 388, 1&
End Sub

Function Session(Str As String)
Session = Chr(Len(Str) + 5) & Chr(Len(Str) + 2) & Chr(Len(Str) + 7) & Chr(Len(Str) + 9)
End Function
Sub MoveForm(TheForm As Form)
'Example
'Form_Mousemove()
'MoveForm Me
'End Sub
ReleaseCapture
Call SendMessage(TheForm.HWnd, &HA1, 2, 0&)
End Sub
Public Function OpenURL(ByVal URL As String) As Long
OpenURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Function RemoveNumbers(Text) As String
Text = Replace(Text, "1", "")
Text = Replace(Text, "2", "")
Text = Replace(Text, "3", "")
Text = Replace(Text, "4", "")
Text = Replace(Text, "5", "")
Text = Replace(Text, "6", "")
Text = Replace(Text, "7", "")
Text = Replace(Text, "8", "")
Text = Replace(Text, "9", "")
Text = Replace(Text, "0", "")
RemoveNumbers = Text
End Function

Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

'The 'Pause' Sub was created to allow you to have custom
'intervals practically doing anything. Example:
'
'Private Sub Form_Load()
'Me.Hide
'Pause 1.0
'Me.Show
'End Sub
