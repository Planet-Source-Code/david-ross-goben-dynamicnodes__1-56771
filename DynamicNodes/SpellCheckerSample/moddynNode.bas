Attribute VB_Name = "moddynNode"
Option Explicit
'*******************************************************************************
' Class PUBLIC stuff
'
' Note that though these variables are declared as Public, they are actually treated like
' Friend objects in that anything not declared within the classes will not be exposed to
' in the DLL interface (though within this project we can see them just fine (as we of
' course can with Friend objects).
'*******************************************************************************
Public g_gSize As Long
Public g_Tag() As Boolean               'True if the global slot in use
Public g_Root() As dynNode              'store root node reference
Public g_ID() As Long                   'globally unique ID number storage
Public g_LastError() As dynErrorCodes   'last error code
Public g_LastErrorNode() As dynNode     'last error node
Public g_MsgBoxErrors() As Boolean      'True if error should be reported in a MsgBox

