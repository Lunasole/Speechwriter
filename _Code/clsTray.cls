VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Const WM_MOUSEMOVE = &H200
'Private Const WM_LBUTTONDOWN = &H201
'Private Const WM_LBUTTONUP = &H202
'Private Const WM_LBUTTONDBLCLK = &H203
'Private Const WM_RBUTTONDOWN = &H204
'Private Const WM_RBUTTONUP = &H205
'Private Const WM_RBUTTONDBLCLK = &H206
'Private Const WM_MBUTTONDOWN = &H207
'Private Const WM_MBUTTONUP = &H208
'Private Const WM_MBUTTONDBLCLK = &H209

'Private Const WM_TrayAction = &H888888

Private Const NIM_ADD = 0
'Private Const NIM_MODIFY = 1
Private Const NIF_MESSAGE = 1
Private Const NIM_DELETE = 2
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4

Private Declare Function Shell_NotifyIconA Lib "shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
    Private Type NOTIFYICONDATA
      cbSize              As Long
      hwnd                As Long
      uID                 As Long
      uFlags              As Long
      uCallbackMessage    As Long
      hIcon               As Long
      szTip               As String * 64
    End Type
Private ThisTray As NOTIFYICONDATA
Private TrayExists As Boolean

Public Function Add(ByVal Handle As Long, ByVal hIcon As Long, ByVal s_Tip As String, ByVal uID As Long) As Long
If Not TrayExists Then
    With ThisTray
        .cbSize = Len(ThisTray)
        .hIcon = hIcon
        .hwnd = Handle
        .szTip = s_Tip & Chr$(&O0)
        .uCallbackMessage = WM_TrayAction
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uID = uID
    End With

    Add = Shell_NotifyIconA(NIM_ADD, ThisTray)
    TrayExists = Add = 1&
End If
End Function

'Public Function ChangeTip(ByVal New_Tip As String) As Long
'If TrayExists Then
'    With ThisTray
'        .szTip = New_Tip & Chr(&O0)
'    End With
'    ChangeTip = Shell_NotifyIconA(NIM_MODIFY, ThisTray)
'End If
'End Function

'Public Function ChangeIcon(ByVal New_Icon As Long) As Long
'If TrayExists Then
'    With ThisTray
'        .hIcon = New_Icon
'    End With
'    ChangeIcon = Shell_NotifyIconA(NIM_MODIFY, ThisTray)
'End If
'End Function
Public Function Reload() As Long
If TrayExists Then
    TrayExists = False
    Call Shell_NotifyIconA(NIM_DELETE, ThisTray)
    Call Add(ThisTray.hwnd, ThisTray.hIcon, ThisTray.szTip, ThisTray.uID)
End If
End Function

Public Function Remove() As Long
If TrayExists Then
    Remove = Shell_NotifyIconA(NIM_DELETE, ThisTray)
    TrayExists = False
        With ThisTray
            .cbSize = 0&
            .hIcon = 0&
            .hwnd = 0&
            .szTip = vbNullString
            .uCallbackMessage = 0&
            .uFlags = 0&
            .uID = 0&
        End With
End If
End Function

Public Function UserID() As Long
    UserID = ThisTray.uID
End Function
