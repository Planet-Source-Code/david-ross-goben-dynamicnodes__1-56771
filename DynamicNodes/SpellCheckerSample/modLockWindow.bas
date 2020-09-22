Attribute VB_Name = "modLockWindow"
Option Explicit
'~modLockWindow.bas;
'Prevent updateson a single window/control
'****************************************************
' modLockWindow():
' Privides the LockWindow() function which will prevent updates
' on a single window/control, such as a TreeView control. This
' function can only lock updates on one item at a time. Unlock the
' window/control by calling the LockWindow() function without a
' parameter. Unlocking the control refreshes the control display image.
'
'EXAMPLE:
'  LockWindow (TreeView1.hWnd) 'lock screen updates for this control
'' process updates for the TreeView control.
'  LockWindow                  'unlock the control and refresh it
'****************************************************

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Function LockWindow(Optional mLNGhWnd As Long = 0) As Long
  On Error Resume Next
  LockWindow = LockWindowUpdate(mLNGhWnd)
End Function

'*******************************************************************************
' LockControlRepaint(): Prevent repaints on specified control
'*******************************************************************************
Public Sub LockControlRepaint(uControl As Control)
  On Error Resume Next
  LockWindowUpdate uControl.hWnd
  If Err <> 0 Then MsgBox uControl.Name & " does not have an hWnd value"
End Sub

'*******************************************************************************
' UnlockControlRepaint(): turn off repaint locking
'*******************************************************************************
Public Sub UnlockControlRepaint(uControl As Control)
  On Error Resume Next
  LockWindowUpdate 0
  uControl.Refresh
End Sub

