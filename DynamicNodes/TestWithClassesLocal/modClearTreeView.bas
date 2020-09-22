Attribute VB_Name = "modClearTreeView"
Option Explicit
'~modClearTreeView.bas;
'Quickly clear a Treeview Control
'*******************************************************************************
' modClearTreeView - The ClearTreeView() subroutine clears a Treeview control
'                    very quickly, much faster than the .Clear method. In a
'                    TreeVIew with 5000 nodes, it will take 3.765 to 3.876
'                    seconds to clear with the .Clear Method. With the
'                    ClearTreeView subroutine, it will clear consistently
'                    in 0.08 to 0.09 seconds.
'*******************************************************************************

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETREDRAW = &HB
Private Const TV_FIRST As Long = &H1100
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Private Const TVGN_ROOT As Long = &H0

Public Sub ClearTreeView(TV As TreeView)
  Dim hItem As Long, hwnd As Long
  
  hwnd = TV.hwnd
  SendMessageByNum hwnd, WM_SETREDRAW, False, 0&  'lock window updates
  Do                                             'clear the treeview
    hItem = SendMessageByNum(hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
    If hItem <= 0 Then Exit Do
    SendMessageByNum hwnd, TVM_DELETEITEM, 0&, hItem
  Loop
  SendMessageByNum hwnd, WM_SETREDRAW, True, 0&   'unlock window updates
End Sub
