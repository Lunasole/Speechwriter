Attribute VB_Name = "modFileAPI"
Option Explicit

Declare Function IsDebuggerPresent Lib "kernel32.dll" () As Long

Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias _
"SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Global Const FOF_ALLOWUNDO = &H40&
Global Const FOF_NOCONFIRMATION = &H10&
'global Const FO_COPY = &H2&
'global Const FO_MOVE = &H1&
Global Const FO_DELETE = &H3&

Declare Function GetTickCount Lib "kernel32" () As Long

Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd&, ByVal lpOperation$, ByVal lpFile$, ByVal lpParameters$, ByVal lpDirectory$, ByVal nShowCmd As VbAppWinStyle)
'global Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName$, ByVal lpNewFileName$) As Long
'global Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName$) As Long
Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath$) As Long

Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal Format As Long, ByVal hMem As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal Flags As Long, ByVal iLen As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Global Const CF_UNICODETEXT = &HD&
Global Const GMEM_MOVEABLE = &O2&
Global Const GMEM_ZEROINIT = &O40&

Global Const OF_READ = &H0
'Private Const OF_READWRITE = &H2
'Private Const OF_REOPEN = &H8000
'Private Const OF_SHARE_COMPAT = &H0
'Private Const OF_SHARE_DENY_NONE = &H40
'Private Const OF_SHARE_DENY_READ = &H30
'Private Const OF_SHARE_DENY_WRITE = &H20
'Private Const OF_SHARE_EXCLUSIVE = &H10
'Private Const OF_VERIFY = &H400
'Private Const OF_WRITE = &H1
Global Const OFS_MAXPATHNAME = 128
Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'**************************************
'This module is used for API declares
'**************************************

'Visit my Homepage at
'http://www.geocities.com/marskarthik
'http://marskarthik.virtualave.net
'Email: marskarthik@angelfire.com
Function FileDate(ByRef inFile As String) As String
Dim hFile As Long, rval As Long
Dim buff As OFSTRUCT
Dim ctime As FILETIME, atime As FILETIME, wtime As FILETIME
Dim ftime As SYSTEMTIME
'Open the File for Reading
hFile = OpenFile(inFile, buff, OF_READ)
If hFile Then
    'Get File time
    rval = GetFileTime(hFile, ctime, atime, wtime)
    'Convert File Time Zone to Local
'    rval = FileTimeToLocalFileTime(ctime, ctime)
    'Convert File Time Format to System Time Format
'    rval = FileTimeToSystemTime(ctime, ftime)
    
    'Convert File Time Zone to Local
    rval = FileTimeToLocalFileTime(wtime, wtime)
    'Convert File Time Format to System Time Format
    rval = FileTimeToSystemTime(wtime, ftime)
    FileDate = ftime.wYear & "-" & ftime.wMonth & "-" & ftime.wDay & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
End If

'Close the File Handle
rval = CloseHandle(hFile)
End Function

