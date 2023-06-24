Attribute VB_Name = "modSysS"
Option Explicit
Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (xpSClass As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200
    
Public Sub InitXPstyler()
Dim xpSClass As tagInitCommonControlsEx
With xpSClass
  .lngSize = Len(xpSClass)
  .lngICC = ICC_USEREX_CLASSES
End With
InitCommonControlsEx xpSClass
End Sub
