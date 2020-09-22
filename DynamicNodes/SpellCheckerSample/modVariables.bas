Attribute VB_Name = "modVariables"
Option Explicit
'*******************************************************************************
' API Stuff
'*******************************************************************************
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Const LB_FINDSTRING = &H18F

'*******************************************************************************
' Global variables
'*******************************************************************************
Public Fso As FileSystemObject    'standard File System Object from Windows Scripting Host
Public RootNode As dynNode        'DynNode root-level node
Public IsDirty As Boolean         'if textbox contents have changed
Public TextFile As String         'name of load file
Public DicFile As String          'name of dictionary file
Public DicChanged As Boolean      'if dictionary has changed
Public MyCancel As Boolean        'general cancellation flag
Public OrgTitle As String         'original title for form

Public ShowingCorrect As Boolean  'true when the user interface for correcting text is displayed
Public CorrectFlag As Long        'the responses from the user interface

Public LtrRef As dynNode          'a reference to a node with the required Letter key
Public SndxRef As dynNode         'a reference to a node with the required Soundx key
Public WordRef As dynNode         'a reference to a node with the searched for word
Public WordIndex As Long          'index of WordRef in LtrRef list

Public SoundexNode As dynNode     'reference to "soundex" key onder root
Public SoundexLtr As dynNode      'ref to letter-soundex key
Public SoundexRef As dynNode      'reference to Soundex entry

Public ChangingText As Boolean    'true when we are tempy changing main form text

Public Punctuation As String
