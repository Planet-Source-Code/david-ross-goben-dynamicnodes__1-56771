VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSpellCheck 
   Caption         =   "Simple Sample Spelling Checker"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9765
   Icon            =   "frmSpellCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8460
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpellCheck.frx":01CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpellCheck.frx":0324
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpellCheck.frx":0776
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpellCheck.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpellCheck.frx":0D22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Key"
            Object.ToolTipText     =   "Open a text file"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save file"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Check"
            Object.ToolTipText     =   "Check Spelling"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Dictionary Maintenance"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh Word Counts"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   9551
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmSpellCheck.frx":0E7C
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1296
            MinWidth        =   882
            Text            =   "Words: 0"
            TextSave        =   "Words: 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2275
            MinWidth        =   882
            Text            =   "Unique Words: 0"
            TextSave        =   "Unique Words: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2593
            MinWidth        =   882
            Text            =   "Unknown Words: 0"
            TextSave        =   "Unknown Words: 0"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open text file..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save text file"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save text file &As..."
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuCheck 
         Caption         =   "Check &Spelling..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDictMaint 
         Caption         =   "Dictionary Maintenance..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Change Font..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh word counts"
      End
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Do not expect too much out of this editor (that's your job).
' This is just a sample of what you can do with spell-checking.
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Set up XP-buttons if enabled
'*******************************************************************************
Private Sub Form_Initialize()
  Call FormInitialize
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Initialize form for display
'*******************************************************************************
Private Sub Form_Load()
  Dim Path As String
  Dim I As Long
  
  OrgTitle = Me.Caption                 'save form original title
  Set Fso = New FileSystemObject
  Me.mnuCheck.Enabled = False           'nothing to check yet
  Me.mnuSave.Enabled = False            'nothing to save
  Me.mnuSaveAs.Enabled = False          'or save as
'
' set up non-text punctuation, less space
'
  For I = 0 To 31                       'control characters
    Punctuation = Punctuation & Chr$(I)
  Next I
  Punctuation = Punctuation & "`~!@#$%^&*()_-+={}[]|\:;""'<>,.?/" 'punctuation
  For I = 127 To 159                    'high-non-text
    Punctuation = Punctuation & Chr$(I)
  Next I
'
' get the selected font to Textbox
'
  ChangingText = True                   'prevent wearing out ChangeText event
  With Me.rtbText
    .Font.Name = GetSetting(App.Title, "Settings", "FontName", CStr(.Font.Name))
    .Font.Size = GetSetting(App.Title, "Settings", "FontSize", CStr(.Font.Size))
    .Font.Bold = CBool(GetSetting(App.Title, "Settings", "FontBold", "0"))
    .Font.Italic = CBool(GetSetting(App.Title, "Settings", "FontItal", "0"))
  End With
  ChangingText = False
'
' auto-load last vaid dictionary file
'
  Path = GetSetting(App.Title, "Settings", "CurrentDictionary", vbNullString)
'
' if a path supplied, see if it still exists
'
  If Len(Path) Then                 'dictionary path present?
    If Fso.FileExists(Path) Then    'yes, but is the file still there?
      Call LoadDictionary(Path)     'it is, so load it
    Else
      Path = vbNullString           'else it is bogus, so reset it to a blank string
    End If
  End If
  Call ResetWordCounts              'set initial word counts to zero
  IsDirty = False                   'start with a clean slate
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : User closing form from "X" button
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    If HandleExit() = vbCancel Then Cancel = 1  'if user canceled, then do not unload form
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the text box to fir the shape of the form
'*******************************************************************************
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub   'do nothing if minimized
  Me.rtbText.Width = Me.ScaleWidth                'fill form with text box
  Me.rtbText.Height = Me.ScaleHeight - Me.StatusBar1.Height - Me.rtbText.Top
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Get rif of created resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Dim Resp As VbMsgBoxResult
  
  If DicChanged Then SaveDictionary   'if dictionary changed, save it
  MyCancel = False                    'init result flag
  If IsDirty Then
    Resp = CenterMsgBoxOnForm(Me, "Save changes to file?", vbYesNoCancel Or vbQuestion, "Save Changes?")
    If Resp = vbCancel Then Exit Sub  'user cancelled
    If Resp = vbYes Then SaveFile     'save file if they want that sort of thing
    If MyCancel Then                  'if user cancelled, then do not unload
      Cancel = 1
      Exit Sub
    End If
  End If
'
' release resources
'
  Call KillCurrentDictionaryData      'release any dictionary resources
  Set Fso = Nothing                   'this one we definitely need to release
End Sub

'*******************************************************************************
' Subroutine Name   : Text1_Change
' Purpose           : Set flags based upon contents of text
'*******************************************************************************
Private Sub rtbText_Change()
  Dim Bol As Boolean
  
  If ChangingText Then Exit Sub           'ignore if temp changing data
  IsDirty = True                          'always dirty if things change
  Bol = CBool(Len(Trim$(Me.rtbText.Text))) 'grab the length of the text
  Me.mnuSaveAs = Bol                      'Save as, if data to save
  Me.mnuSave.Enabled = Bol                'save contents, if something to save
  Me.mnuCheck.Enabled = Bol               'spell checking
End Sub

'*******************************************************************************
' Subroutine Name   : mnuDictMaint_Click
' Purpose           : Perform dictionary maintenance
'*******************************************************************************
Private Sub mnuDictMaint_Click()
  frmDictionaryIO.Show vbModal, Me
End Sub

'*******************************************************************************
' Subroutine Name   : mnuExit_Click
' Purpose           : User chose to exit the program
'*******************************************************************************
Private Sub mnuExit_Click()
  If HandleExit() <> vbCancel Then Unload Me  'whew! All done, so exit
End Sub

'*******************************************************************************
' Function Name     : HandleExit
' Purpose           : Save the dictionary if it has changed.
'                   : Propt for saving the file if it is dirty
'*******************************************************************************
Private Function HandleExit() As VbMsgBoxResult
  If DicChanged Then SaveDictionary             'if dictionary changed, save it
  MyCancel = False                              'init result flag
  If IsDirty Then                               'prompt for save if the file is dirty
    HandleExit = CenterMsgBoxOnForm(Me, "Save changes to file?", vbYesNoCancel Or vbQuestion, "Save Changes?")
    If HandleExit = vbCancel Then Exit Function 'if user decided to cancel
    If HandleExit = vbYes Then SaveFile         'they want to save and exit
    If MyCancel Then
      HandleExit = vbCancel                     'a subsequent cancel was issued
    End If
    IsDirty = False                             'turn off dirty if no desire to save
  End If
End Function

'*******************************************************************************
' Subroutine Name   : mnuFont_Click
' Purpose           : Change the font
'*******************************************************************************
Private Sub mnuFont_Click()
  With Me.CommonDialog1
    .Flags = cdlCFScreenFonts               'allow just screen fonts
    .FontName = Me.rtbText.Font.Name        'set default pont to current
    .FontSize = Me.rtbText.Font.Size
    .FontBold = Me.rtbText.Font.Bold
    .FontItalic = Me.rtbText.Font.Italic
    .DialogTitle = "Select Font"            'set title for dialog
    .CancelError = True                     'allow user cancel
    On Error Resume Next                    'trap the cancel
    .ShowFont                               'show the font dialog
    If Err.Number Then Exit Sub             'user cancelled
    On Error GoTo 0
    ChangingText = True                     'prevent setting IsDirty during the following...
    Me.rtbText.Font.Name = .FontName        'set new font info
    Me.rtbText.Font.Size = .FontSize
    Me.rtbText.Font.Bold = .FontBold
    Me.rtbText.Font.Italic = .FontItalic
    ChangingText = False
'
' save new setting to registry
'
    Call SaveSetting(App.Title, "Settings", "FontName", .FontName)
    Call SaveSetting(App.Title, "Settings", "FontSize", CStr(.FontSize))
    Call SaveSetting(App.Title, "Settings", "FontBold", CStr(.FontBold))
    Call SaveSetting(App.Title, "Settings", "FontItal", CStr(.FontItalic))
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : mnuRefresh_Click
' Purpose           : Refresh word counts
'*******************************************************************************
Private Sub mnuRefresh_Click()
  ResetWordCounts
End Sub

'*******************************************************************************
' Subroutine Name   : mnuSave_Click
' Purpose           : Save the Text File
'*******************************************************************************
Private Sub mnuSave_Click()
  If IsDirty Then Call SaveFile
End Sub


'*******************************************************************************
' Function Name     : mnuSaveAs
' Purpose           : Save the file to a specified file
'*******************************************************************************
Private Sub mnuSaveAs_Click()
  Call SaveAs
End Sub

'*******************************************************************************
' Subroutine Name   : mnuOpen_Click
' Purpose           : Open and load a Text File
'*******************************************************************************
Private Sub mnuOpen_Click()
  Dim Ts As TextStream
  Dim Str As String
  
  With Me.CommonDialog1
    .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNLongNames Or cdlOFNExplorer
    .DefaultExt = "txt"
    .FileName = vbNullString
    .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    .DialogTitle = "Open Text File"
    .CancelError = True
    On Error Resume Next                                      'trap user cancel
    .ShowOpen                                                 'show the dialog
    If Err.Number Then Exit Sub                               'the user cancelled
    On Error GoTo 0
    If Len(Trim$(.FileName)) = 0 Then Exit Sub
    TextFile = Trim$(.FileName)
  End With
'
' read the file into the text box
'
  Set Ts = Fso.OpenTextFile(TextFile, ForReading, False)
  Me.rtbText.Text = Ts.ReadAll
  Ts.Close
'
' show filename in form caption
'
  Me.Caption = OrgTitle & " - " & TextFile
'
' reset the word counts
'
  Call ResetWordCounts
'
' indicate a clean state
'
  IsDirty = False
End Sub

'*******************************************************************************
' Subroutine Name   : mnuCheck_Click
' Purpose           : Check the spelling of the file
'*******************************************************************************
Private Sub mnuCheck_Click()
  Dim Ary() As String, Test As String, Sample As String, ChgAll() As String, C As String
  Dim Text As String, Str As String, Tmp As String, Word As String, Chg() As String
  Dim Index As Long, Count As Long, I As Long, WordSt As Long
  Dim J As Long, K As Long, L As Long, AddedWords As Long, Offset As Long
  Dim Inword As Boolean
  Dim NewWords() As String    'the list of words that are not in the dictionary
  Dim NewIdx As Long         'the Ubound of NewWOrds()
'
' if the dictionary object is not loaded, then we can do nothing
'
  If Len(DicFile) = 0 Then
    Do
      frmDictionaryIO.Show vbModal, Me
      If Not RootNode Is Nothing Then Exit Do
      If Len(DicFile) <> 0 Then Exit Do
      If CenterMsgBoxOnForm(Me, "A dictionary file (even a new empty one) is not present." & vbCrLf & _
             "Cannot perform a spell check, or add words to it, without" & vbCrLf & _
             "a dictionary file.", vbRetryCancel Or vbExclamation, "No Dictionary") = vbCancel Then Exit Sub
    Loop
  End If
  
'
' prepare the operating room
'
  Me.Enabled = False                          'disable user-editing
  Screen.MousePointer = vbHourglass           'show that we are busy
  DoEvents
  
  With Me.rtbText
    If .SelLength Then
      Ary() = SplitText(.SelText)             'get list of unique words from selection
    Else
      Ary() = SplitText(.Text)                'get list of unique words from all
    End If
  End With
  
  On Error Resume Next                        'don't explode if array not dimmed...
  Count = UBound(Ary)                         'get upper bounds
  If Err.Number = 0 Then
    If Count <> 0 Then
      Ary = GetUniqueWords(Ary)               'reduce to unique
      On Error Resume Next                    'don't explode if array not dimmed...
      Count = UBound(Ary)                     'get upper bounds
    End If
  End If
'
' if there is an error, then there are no words to check for. This is a good thing...
'
  If Err.Number Then
    CenterMsgBoxOnForm Me, "No words to check.", vbOKOnly Or vbInformation, "No Words"
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
  End If
  On Error GoTo 0
'
' now go through and gather the list of new words
'
  ReDim NewWords(Count) As String
  NewIdx = -1
  For Index = 0 To Count
    If Not QuickFindMatch(Ary(Index)) Then
      NewIdx = NewIdx + 1
      NewWords(NewIdx) = Ary(Index)
    End If
  Next Index
'
' show not busy anymore
'
  Screen.MousePointer = vbDefault               'show no longer busy
  DoEvents
'
' if all the words were found in the dictionary, then we are REALLY OK!
'
  If NewIdx = -1 Then
    CenterMsgBoxOnForm Me, "Congradulations. All words are in the dictionary.", vbOKOnly Or vbInformation, "No Unfound Words"
    Me.Enabled = True
    Exit Sub
  End If
  
  If NewIdx <> Count Then
    ReDim Preserve NewWords(NewIdx) As String   'set proper array size
  End If
  ReDim ChgAll(NewIdx) As String                'set change all array
'
' well, time to process the errors
'
  CorrectFlag = 0                               'init user response flag
  ShowingCorrect = False                        'we have not shown the correction window
'
' if a block of text is selected, use that block only. Otherwise use whole text
'
  With Me.rtbText
    If .SelLength Then
      Text = LCase$(.SelText)                   'grab LC version of text
      Offset = .SelStart                        'adjustment offset
      .SelLength = 0                            'remove selection
    Else
      Text = LCase$(Me.rtbText.Text)            'grab LC version of text
      Offset = 0
    End If
  End With
  Test = " " & Punctuation
  Count = Len(Text)                             'length of text file
  Inword = False                                'init not being inside a word
'
' loop through the text data and find words
'
  AddedWords = 0
  I = 0                                         'init the index into the string
  Do While I <= Count
    If I = Count Then                           'if at max...
      If Not Inword Then Exit Do                'and not in word, then done
    End If
    I = I + 1                                   'new index position
    If I > Count Or InStr(1, Test, Mid$(Text, I, 1)) <> 0 Then 'a pinctuation char?
      If Inword Then                            'yes, are we in a word?
        Word = Mid$(Text, WordSt, I - WordSt)   'yes, so grab full word
        Inword = False                          'and mark no longer in a word
        '
        ' check through the new word list for a matching word
        '
        For Index = 0 To NewIdx
          If CBool(Len(NewWords(Index))) Then   'if a word is still there...
            If StrComp(NewWords(Index), Word, vbBinaryCompare) = 0 Then
 '
 ' see if we are simply adding all unfound words to the dictionary
 '
              If CorrectFlag = 6 Then           'it is THE word. are we ADDING ALL?
                AddWordToDictionary Word       'yes, so add it to the dictionary
                NewWords(Index) = vbNullString 'remove from further checks
                AddedWords = AddedWords + 1
'
' are we changing ALL of this word?
'
              ElseIf Len(ChgAll(Index)) > 0 Then    'change ALL of this word?
                With Me.rtbText                     'yes...
                  .SelStart = WordSt - 1 + Offset   'first select it in in text
                  .SelLength = Len(Word)            'length of current word
                  Tmp = ChgAll(Index)               'get replacement text
                  K = Len(Tmp)                      'length of new data
                  Me.rtbText.SelText = Tmp          'change the word
                  I = I + (K - Len(Word))           'point beyond it and punct. mark
                End With
                IsDirty = True                      'text has changed
                Text = LCase$(Me.rtbText.Text)      'grab LC version of updated text
'
' doing normal checks with user interface
'
              Else
                With Me.rtbText
                  ChangingText = True                   'rprevent some auto-updates
                  LockControlRepaint Me.rtbText         'prevent textbox control jitters
                  C = Mid$(.Text, WordSt + Offset, 1)   'get first character of word
                  .SelStart = WordSt - 1 + Offset       'put tag in first character
                  .SelLength = 1
                  .SelText = Chr$(255)
                  J = WordSt - 150                      'compute grab start
                  If J < 1 Then J = 1                   'too low, to set to 1
                  Sample = Mid$(.Text, J + Offset, 500) 'grab sample data
                  .SelStart = WordSt - 1 + Offset       'put character back
                  .SelLength = 1
                  .SelText = C
                  .SelStart = WordSt - 1 + Offset       'mark in text
                  .SelLength = Len(Word)
                  ChangingText = False                  'rprevent some auto-updates
                  UnlockControlRepaint Me.rtbText       'now allow refreshes
                  '
                  ' if in a word, then find end of it
                  '
                  K = 0
                  If InStr(1, Test, Mid(Sample, K + 1, 1)) = 0 And J - Offset > 1 Then
                    Do
                      K = K + 1                   'now find the beginning of a word
                    Loop While InStr(1, Test, Mid(Sample, K, 1)) = 0
                  End If
                  '
                  ' then skip past any punctuation to the beginning of a word
                  '
                  Do
                      K = K + 1                   'now find the beginning of a word
                  Loop While InStr(1, Test, Mid(Sample, K, 1))
                  '
                  ' now grab the sample of text (up to 300 characters...
                  '
                  If K > 1 Then Sample = Mid$(Sample, K)
                  J = InStr(1, Sample, Chr$(255))   'find marker
                  Mid$(Sample, J, 1) = C            'put original character back
                  '
                  ' now set up correction form
                  '
                  If Not ShowingCorrect Then Load frmCorrect
                  With frmCorrect
                   .txtFix.Text = Mid$(Me.rtbText.Text, WordSt + Offset, I - WordSt)
                   .cmdChange.Enabled = False       'init change buttons disabled
                   .cmdChangeAll.Enabled = False
                   .SelectionStart = J - 1
                   .SelectionLength = Len(Word)
                    With .txtSample
                      .Text = Sample                'stuff sample
                      .SelStart = J - 1             'select word in question
                      .SelLength = Len(Word)
                    End With
                    '
                    ' show any suggested word we may have found
                    '
                    .lstSuggest.Enabled = True
                    .lstSuggest.Clear               'clear list regardless
                    Call QuickFindMatch(Word)       'see if suggestions possible
                    If Not SndxRef Is Nothing Then
                      With SndxRef.Nodes
                        J = .Count
                        For K = 1 To J
                          frmCorrect.lstSuggest.AddItem .Item(K).Text
                        Next K
                      End With
                      .lstSuggest.ListIndex = -1
                    ElseIf QuickFindSndX(Word) Then        'found Soundex key for it?
                      With SoundexRef.Nodes                'yes, so grab the list
                        J = .Count
                        For K = 1 To J
                          frmCorrect.lstSuggest.AddItem .Item(K).Text
                        Next K
                      End With
                    Else
                      .lstSuggest.AddItem "(no suggestions)" 'no suggestions
                      .lstSuggest.Enabled = False
                    End If
                    '
                    ' show the correction form if not presently being shown
                    '
                    .SelFromUser = False
                    If Not ShowingCorrect Then
                      ShowingCorrect = True
                      .Show vbModeless, Me
                    End If
                    .txtFix.SetFocus
                    '
                    ' set the response flag to 0 and wait for it to change
                    '
                    CorrectFlag = 0
                    Do While CorrectFlag = 0
                      Sleep 100                     'this keeps the resource meter happy
                      DoEvents                      'let user interact
                    Loop
                    '
                    ' now check what the user decided to do
                    '
                    Select Case CorrectFlag
                      Case -1                             'cancel
                        Unload frmCorrect                 'they cancelled, so ensure form unloaded
                        ShowingCorrect = False            'flag the form not seen
                        Exit For
                        
                      Case 1                              'ignore once
                      
                      Case 2                              'ignore all
                        NewWords(Index) = vbNullString    'remove from checks
                        
                      Case 3                              'add to dictionary
                          AddWordToDictionary Word        'add the word
                          NewWords(Index) = vbNullString  'remove from checks
                      
                      Case 4, 5                           'change/change all
                          Tmp = frmCorrect.txtFix.Text    'text to replace word with
                          K = Len(Tmp)                    'get length
                          Me.rtbText.SelText = Tmp        'stuff word
                          I = I + (Len(Tmp) - Len(Word))  'point past it and punctuation
                          
                          If CorrectFlag = 5 Then         'change all?
                            ChgAll(Index) = Tmp           'yes, so stuff flag with new text
                          End If
                          IsDirty = True                  'text has changed
                          Text = LCase$(Me.rtbText.Text)  'grab LC version of updated text
                      
                      Case 6                              'add all
                        Unload frmCorrect                 'remove user-interface
                        ShowingCorrect = False            'flag it not shown
                        Screen.MousePointer = vbHourglass
                        DoEvents                          'we may be busy for a while...
                        RootNode.KeyChecks = False        'all of these will be new
                        AddWordToDictionary Word          'add current word
                        NewWords(Index) = vbNullString    'remove from further checks
                        AddedWords = 1
                        Load frmMsg
                        With frmMsg
                          .lblMsg.Caption = "Adding All Unknown Words..."
                          .Show vbModeless, Me
                        End With
                        Screen.MousePointer = vbHourglass
                        DoEvents
                      Case 7                        'review unknown words
                        ReDim Chg(NewIdx) As String 'define temp array
                        J = -1
                        For K = 0 To NewIdx
                          If Len(NewWords(K)) > 0 And Len(ChgAll(K)) = 0 Then
                            J = J + 1
                            Chg(J) = NewWords(K)
                          End If
                        Next K
                        If J < 0 Then
                          CenterMsgBoxOnForm Me, "No more unknown words to review.", vbOKOnly Or vbInformation, "Nothing To Review"
                        Else
                          If J <> NewIdx Then
                            ReDim Preserve Chg(J) As String
                          End If
                              
                          Load frmReview
                          With frmReview
                            .lstWords.Clear     'init main list
                            .lstAdd.Clear       'init ADD list
                            .lstIgnore.Clear    'init REMOVE list
                            For K = 0 To J
                              .lstWords.AddItem Chg(K)
                            Next K
                            .lblbWordCount.Caption = "Word Count: " & CStr(.lstWords.ListCount)
                            CorrectFlag = 0
                            .Show vbModal, Me
                            If (.lstAdd.ListCount > 0 Or .lstIgnore.ListCount > 0) And CorrectFlag = 8 Then
                                With .lstAdd
                                  For K = 0 To .ListCount
                                    Tmp = .List(K)
                                    For L = 0 To NewIdx
                                      If StrComp(Tmp, NewWords(K), vbBinaryCompare) Then
                                        NewWords(K) = vbNullString  'remove selection
                                        AddWordToDictionary Tmp     'add these to dict
                                      End If
                                    Next L
                                  Next K
                                End With
                                With .lstIgnore
                                  For K = 0 To .ListCount
                                    Tmp = .List(K)
                                    For L = 0 To NewIdx
                                      If StrComp(Tmp, NewWords(K), vbBinaryCompare) Then
                                        NewWords(K) = vbNullString  'remove selection
                                      End If
                                    Next L
                                  Next K
                                End With
                            End If
                            Unload frmReview
                            I = I - Len(Word) - 1   'point back to start of word - 1
                            CorrectFlag = 0
                            Exit For
                          End With
                        End If
                    End Select
                  End With
                End With
                Exit For                            'found our word, so exit loop
              End If
            End If
          End If
        Next Index                                  'check for next word
      End If
'
' if we are not inside a word, and we found text
'
    ElseIf Not Inword Then
      WordSt = I                                    'start new word
      Inword = True                                 'indicate we know where we are
    End If
    
    If CorrectFlag < 0 Then Exit Do                 'if the user is aborting
  Loop                                              'check entire text
  RootNode.KeyChecks = True                         'make sure Key checking turned on
'
' unload the user interface if it is still displayed
'
  If ShowingCorrect Then Unload frmCorrect
'
' ensure we are not busy and allow user interaction again
'
  With Me.rtbText
    .SelStart = 0
    .SelLength = 0
  End With
  Call ResetWordCounts            'reset word counts
  Screen.MousePointer = vbDefault
'
' if Adding ALL, then report how many words added to the dictionary
'
  If CorrectFlag = 6 Then
    Unload frmMsg
    If CBool(AddedWords) Then
      Tmp = CStr(AddedWords) & " word"
      If AddedWords <> 1 Then Tmp = Tmp & "s"
      Screen.MousePointer = vbDefault
      CenterMsgBoxOnForm Me, Tmp & " added to the dictionary.", vbOKOnly Or vbInformation, "Add All"
    End If
  End If
  Me.Enabled = True                 'allow user to interact with the text
End Sub

'*******************************************************************************
' Subroutine Name   : Toolbar1_ButtonClick
' Purpose           : Process Toolbar button selects
'*******************************************************************************
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1
      Call mnuOpen_Click
    Case 2
      Call mnuSave_Click
    Case 4
      Call mnuCheck_Click
    Case 5
      Call mnuDictMaint_Click
    Case 6
      Call mnuRefresh_Click
  End Select
End Sub
