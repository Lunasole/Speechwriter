Attribute VB_Name = "modSysT"
'**********************************************************
'EXAMPLE CODE HERE
'Move it to one of your modules
'
'Then call hook when form starts
'Create new clsTray for each form that needs tray icon
'Use class methods to deal with tray
'Call unhook when form unloads
'**********************************************************

Option Explicit

Public MainTray As New clsT


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public Const WM_TrayAction = &H18894 'Custom message for tray actions

Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public WM_TrayReload As Long 'Tray message

Public OldProc As Long

'call this when the form is loading
Public Sub hook(ByRef hwnd&)
OldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
WM_TrayReload = RegisterWindowMessage("TaskbarCreated")
End Sub

'call this when form unloads
Public Sub unhook(ByRef hwnd&)
SetWindowLong hwnd, GWL_WNDPROC, OldProc
End Sub

'this funct shall receive any window messages, since hook has been called
Public Function NewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If msg = WM_TrayAction Then 'Tray message
    'Wparam is UID here
    'LParam is mouse action like move-click-button down, etc
    'And so strange, but hwnd is hwnd here

    If wParam = 1& Then 'Check UID
     Select Case lParam
        Case &H200: '200 - mouse move
            'Your code for tray mouse move

        Case &H201: '201 - ldown
            'Your code for tray left mouse down
        Case &H202: '202 - lup
            'Your code for tray left mouse up
            MainTray.Remove
            frmMain.Visible = True
            
        Case &H203: '203 - ldblclick
            'Your code for tray left mouse doubleclick
            
        Case &H204: '204 - rdown
            'Your code for tray right mouse down
        Case &H205: '205 - rup
'            MainTray.Remove
'            frmMain.Visible = True
            With frmMain
                .mnuTrayGen.Visible = .AutoGen = 0
                .mnuTrayGenA.Visible = Not .mnuTrayGen.Visible
'                .mnuTrayGen.Caption = IIf(.AutoGen = 0, "Текст", "[авто-текст]")
                If .AutoGen = 0 Then
                    Dim TInt As CommandButton
                    Dim TMnu As Integer
                      For Each TInt In .cmdS()
                        .mnuTrayGenT(TMnu).Caption = TInt.Caption
                        .mnuTrayGenT(TMnu).Tag = TInt.Index
                        TMnu = TMnu + 1
                      Next TInt
                End If
            End With
            frmMain.PopupMenu frmMain.mnuTray, vbPopupMenuRightButton, , , frmMain.mnuTrayShow
            
        Case &H206: '206 - rdblclick
            'Your code for tray right mouse doubleclick
            
        '... and so on, you'll need to capture lParam for other actions
     End Select
    End If
    Exit Function
ElseIf msg = WM_TrayReload Then
    'Tray has been reloaded, for example when explorer.exe restarted
    'It is recommended to redraw all your tray icons
    MainTray.Reload
End If

    'send other messages to its owner
    NewWindowProc = CallWindowProc(OldProc, hwnd, msg, wParam, lParam)
End Function

