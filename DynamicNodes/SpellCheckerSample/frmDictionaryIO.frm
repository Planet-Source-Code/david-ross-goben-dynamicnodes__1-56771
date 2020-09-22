VERSION 5.00
Begin VB.Form frmDictionaryIO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dictionary Maintenance"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import a word list..."
      Height          =   375
      Left            =   2100
      TabIndex        =   2
      ToolTipText     =   "Import words from a text file to this dictionary"
      Top             =   960
      Width           =   1755
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7140
      TabIndex        =   0
      Top             =   960
      Width           =   1755
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&xport to Text File..."
      Height          =   375
      Left            =   180
      TabIndex        =   3
      ToolTipText     =   "Export the dictionary word list to a text file"
      Top             =   960
      Width           =   1755
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit/View  Word List"
      Height          =   375
      Left            =   4020
      TabIndex        =   1
      ToolTipText     =   "Edit the current dictionary's word list"
      Top             =   960
      Width           =   1755
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   315
      Left            =   8100
      TabIndex        =   4
      ToolTipText     =   "Create a new dictionary file"
      Top             =   480
      Width           =   795
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Default         =   -1  'True
      Height          =   315
      Left            =   7620
      TabIndex        =   5
      ToolTipText     =   "Browse for a dictionary file"
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblCurrent 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Not Loaded)"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   480
      Width           =   7335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Dictionary:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "frmDictionaryIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MeCaption As String

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Form Setup
'*******************************************************************************
Private Sub Form_Load()
  MeCaption = Me.Caption
  Me.Icon = frmSpellCheck.Icon
  If CBool(Len(DicFile)) Then               'if a dictionary file has been loaded...
    Me.lblCurrent.Caption = DicFile         'display its name
  Else
    Me.lblCurrent.Caption = "(Not Loaded)"  'otherwise, show that nothing is loaded
    Me.cmdExport.Enabled = False            'disable any exporting
    Me.cmdEdit.Enabled = False              'and disable the edit word list
  End If
  UpdateDictionaryCount                     'show dictionary word count
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClose_Click
' Purpose           : Get out of Dodge. Close the dialog
'*******************************************************************************
Private Sub cmdClose_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdBrowse_Click
' Purpose           : Open a dictionary file
'*******************************************************************************
Private Sub cmdBrowse_Click()
  Dim TxtFile As String
  Dim Ts As TextStream
'
' save off any changes to the current dictionary, if there are any
'
  If DicChanged Then SaveDictionary  'if dictionary changed, save it
'
' set up Open Dialog properties, then select a file
'
  With frmSpellCheck.CommonDialog1
    .Flags = cdlOFNLongNames Or cdlOFNExplorer Or cdlOFNFileMustExist 'MUST EXIST
    .DefaultExt = "txt"
    .FileName = vbNullString
    .Filter = "Dictionary File (*.dct)|*.dct"
    .DialogTitle = "Open an Existing  Dictionary"
    .CancelError = True
    On Error Resume Next
    .ShowOpen
    If Err.Number Then Exit Sub           'user cancelled
    On Error GoTo 0
    TxtFile = Trim$(.FileName)            'get user selection
    If Len(TxtFile) = 0 Then Exit Sub     'secondary check
  End With
  Me.lblCurrent.Caption = TxtFile         'show path
'
' now load the file
'
  Call LoadDictionary(TxtFile)          'load it
'
' reset word counts
'
  Call ResetWordCounts
  Me.cmdExport.Enabled = Not SoundexNode Is Nothing 'if words were added
  Me.cmdEdit.Enabled = Me.cmdExport.Enabled
  UpdateDictionaryCount                     'show dictionary word count
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCreate_Click
' Purpose           : Create a new dictionary
'*******************************************************************************
Private Sub cmdCreate_Click()
  Dim TxtFile As String
  Dim Ts As TextStream
'
' save off any changes to the current dictionary, if there are any
'
  If DicChanged Then SaveDictionary  'if dictionary changed, save it
'
' set up Open Dialog properties, then select a file
'
  With frmSpellCheck.CommonDialog1
    .Flags = cdlOFNLongNames Or cdlOFNExplorer Or cdlOFNOverwritePrompt 'OVER_WRITE PROMPT
    .DefaultExt = "txt"
    .FileName = vbNullString
    .Filter = "Dictionary File (*.dct)|*.dct"
    .DialogTitle = "Create a Dictionary"
    .CancelError = True
    On Error Resume Next
    .ShowOpen
    If Err.Number Then Exit Sub
    On Error GoTo 0
    TxtFile = Trim$(.FileName)
    If Len(TxtFile) = 0 Then Exit Sub
  End With
'
' if file is new, then intialize with just header
' (we had already asked if we want to over-write it
' with the ShowOpen dialog)
'
  Set Ts = Fso.CreateTextFile(TxtFile, True)
  Ts.Write "NODEDIC"
  Ts.Close
'
' finally load it (even if new) and update word counts
'
  Me.lblCurrent.Caption = TxtFile
  Call LoadDictionary(TxtFile)                  'load the dictionary
'
' reset word counts
'
  Call ResetWordCounts
  Me.cmdExport.Enabled = Not SoundexNode Is Nothing 'if words were added
  Me.cmdEdit.Enabled = Me.cmdExport.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : cmdExport_Click
' Purpose           : Export the dictionary to a text file. This will create a file with
'                   : each word in the dictionary on its own line
'*******************************************************************************
Private Sub cmdExport_Click()
  Dim Ts As TextStream
  Dim Ary() As String, TxtFile As String
'
' get path to file to save information to
'
  With frmSpellCheck.CommonDialog1
    .Flags = cdlOFNLongNames Or cdlOFNExplorer Or cdlOFNOverwritePrompt
    .DefaultExt = "txt"
    .FileName = "Untitled"
    .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    .DialogTitle = "Export Dictionary Word List"
    .CancelError = True
    On Error Resume Next
    .ShowSave
    If Err.Number Then Exit Sub
    On Error GoTo 0
    If Len(Trim$(.FileName)) = 0 Then Exit Sub
    TxtFile = .FileName
  End With
'
' we have the filem now get the complete list of words
'
  Ary = GetDictionaryList()
'
' write this list to the file
'
  Set Ts = Fso.OpenTextFile(TxtFile, ForWriting, True)
  Ts.Write Join(Ary, vbCrLf)
  Ts.Close
End Sub

'*******************************************************************************
' Subroutine Name   : cmdImport_Click
' Purpose           : Import a list of words from a text file
'*******************************************************************************
Private Sub cmdImport_Click()
  Dim TxtFile As String, Str As String, Ary() As String, NewWords() As String
  Dim Index As Long, Count As Long, NewIdx As Long
  Dim Ts As TextStream
'
' set up Open Dialog properties, then select a file
'
  With frmSpellCheck.CommonDialog1
    .Flags = cdlOFNLongNames Or cdlOFNExplorer Or cdlOFNFileMustExist 'MUST EXIST
    .DefaultExt = "txt"
    .FileName = vbNullString
    .Filter = "Text File (*.txt)|*.txt"
    .DialogTitle = "Open an Existing  Dictionary"
    .CancelError = True
    On Error Resume Next
    .ShowOpen
    If Err.Number Then Exit Sub
    On Error GoTo 0
    TxtFile = Trim$(.FileName)
    If Len(TxtFile) = 0 Then Exit Sub
  End With
'
' show that we are busy
'
  Screen.MousePointer = vbHourglass
  DoEvents
'
' load the file
'
  Set Ts = Fso.OpenTextFile(TxtFile, ForReading, False)
  Str = Trim$(Ts.ReadAll)
  Ts.Close
  NewIdx = -1                                 'init new word index
'
' if the file was empty...
'
  If Len(Str) = 0 Then
    Screen.MousePointer = vbDefault
    CenterMsgBoxOnForm frmSpellCheck, "No words to Import.", vbOKOnly Or vbInformation, "No Words"
    Exit Sub
  End If
'
' break the text up into unique words
'
  Ary() = SplitText(Str)                      'get list of unique words from all
  On Error Resume Next                        'don't explode if array not dimmed...
  Count = UBound(Ary)                         'get upper bounds
  If Err.Number = 0 Then
    If CBool(Count) Then                      'if list exists
      Ary = GetUniqueWords(Ary)               'reduce to unique
      On Error Resume Next                    'don't explode if array not dimmed...
      Count = UBound(Ary)                     'get upper bounds
    End If
    If Err.Number = 0 Then
'
' now go through and gather the list of new words
'
      ReDim NewWords(Count) As String
      For Index = 0 To Count
        If Not QuickFindMatch(Ary(Index)) Then  'found this word in the dictionary?
          NewIdx = NewIdx + 1                   'no, so add it to the list
          NewWords(NewIdx) = Ary(Index)
        End If
      Next Index
    End If
  End If
'
' if no new words exist...
'
  If NewIdx = -1 Then
    Screen.MousePointer = vbDefault
    CenterMsgBoxOnForm frmSpellCheck, "No NEW words to Import.", vbOKOnly Or vbInformation, "No New Words"
    Exit Sub
  End If
  On Error GoTo 0
'
' add new words to the dictionary
'
  For Index = 0 To NewIdx
    AddWordToDictionary Ary(Index)       'add it to the dictionary
  Next Index
'
' reset word counts
'
  Call ResetWordCounts
  Me.cmdExport.Enabled = Not SoundexNode Is Nothing 'if words were added
  Me.cmdEdit.Enabled = Me.cmdExport.Enabled
'
' no longer busy
'
  Screen.MousePointer = vbDefault
'
' issue our report to the board...
'
  Str = CStr(NewIdx + 1) & " imported word"
  If CBool(NewIdx) Then Str = Str & "s"
  CenterMsgBoxOnForm frmSpellCheck, Str & " added to the dictionary.", vbOKOnly Or vbInformation, "Words Imported"
  UpdateDictionaryCount                     'show dictionary word count
End Sub

'*******************************************************************************
' Subroutine Name   : cmdEdit_Click
' Purpose           : Edit the word list in the dictionary
'*******************************************************************************
Private Sub cmdEdit_Click()
  frmEditDictionary.Show vbModal, Me
  UpdateDictionaryCount                     'show dictionary word count
End Sub

Private Sub UpdateDictionaryCount()
  Dim Ary() As String
  Dim Count As Long
  
  Ary = GetDictionaryList()
  On Error Resume Next
  Count = UBound(Ary)
  If Err.Number Then Count = -1
  Me.Caption = MeCaption & " - Words in Dictionary: " & CStr(Count + 1)
End Sub
