Attribute VB_Name = "modDictionaryIO"
Option Explicit


'*******************************************************************************
' Subroutine Name   : SaveDictionary
' Purpose           : Save the contents of the dictionary data
'*******************************************************************************
Public Sub SaveDictionary()
  Dim Ts As TextStream
  Dim Str As String
  Dim Node As dynNode
  Dim Index As Long, Count As Long, Idx As Long, Cnt As Long
  Dim I As Long, C As Long
  
  If RootNode Is Nothing Then Exit Sub        'nothing to do
  With frmMsg
    .lblMsg.Caption = "Saving Dictionary..."
    .Show vbModeless, frmSpellCheck
  End With
  Screen.MousePointer = vbHourglass
  DoEvents
  Set Ts = Fso.CreateTextFile(DicFile, True)
  Ts.Write "NODEDIC"                          'save identifying header
  
  RootNode.SortChildren True                  'sort everything
  With RootNode.Nodes
    Count = .Count
    For Index = 1 To Count                      'parse Letter Children
      If StrComp(.Item(Index).Key, "soundex") <> 0 Then 'Skip SoundEx lists
        With .Item(Index).Nodes                 'process grandchildren to file
          Cnt = .Count                          'get grandchild count
          For Idx = 1 To Cnt                    'process them all
            Str = .Item(Idx).Key                'get Soundex value
            With .Item(Idx).Nodes               'parse word list
              C = .Count                        'get number of words under this soundex key
              For I = 1 To C
                Str = Str & "|" & .Item(I).Text
              Next I
              Ts.WriteLine Str                  'write entry to file
            End With
          Next Idx                              'do all grandchildren
        End With
      End If
    Next Index                                  'loop through all children
  End With
  Ts.Close
  DicChanged = False
  Unload frmMsg
  Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Subroutine Name   : ResetWordCounts
' Purpose           : Reset word counts in status bar
'*******************************************************************************
Public Sub ResetWordCounts()
  Dim Ary() As String, NewWords() As String
  Dim Index As Long, Count As Long, NewIdx As Long
  
  If CBool(Len(Trim$(frmSpellCheck.rtbText.Text))) And CBool(Len(DicFile)) Then
    Screen.MousePointer = vbHourglass           'show that we are busy
    DoEvents
    
    Ary() = SplitText(frmSpellCheck.rtbText.Text) 'get list of unique words
    On Error Resume Next                        'don't explode if array not dimmed...
    Count = UBound(Ary)                         'get upper bounds
    If Err.Number = 0 Then
      frmSpellCheck.StatusBar1.Panels(1).Text = "Words: " & CStr(Count + 1) & " "
      Ary = GetUniqueWords(Ary)                 'reduce to unique
      On Error Resume Next                      'don't explode if array not dimmed...
      Count = UBound(Ary)                       'get upper bounds
      If Err.Number = 0 Then
        On Error GoTo 0
        frmSpellCheck.StatusBar1.Panels(2).Text = "Unique Words: " & CStr(Count + 1) & " "
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
        If NewIdx = -1 Then
          frmSpellCheck.StatusBar1.Panels.Item(3).Text = "Unknown Words: 0 "
          frmSpellCheck.mnuCheck.Enabled = False
        Else
          frmSpellCheck.StatusBar1.Panels.Item(3).Text = "Unknown Words: " & CStr(NewIdx + 1) & " "
          frmSpellCheck.mnuCheck.Enabled = True
        End If
      Else
        With frmSpellCheck.StatusBar1.Panels
          .Item(2).Text = "Unique Words: 0 "
          .Item(3).Text = "Unknown Words: 0 "
        End With
      End If
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
  End If
'
' default here if nothing to use
'
  Screen.MousePointer = vbDefault
  With frmSpellCheck.StatusBar1.Panels
    .Item(1).Text = "Words: 0 "
    .Item(2).Text = "Unique Words: 0 "
    .Item(3).Text = "Unknown Words: 0 "
  End With
End Sub

'*******************************************************************************
' Function Name     : GetDictionaryList
' Purpose           : Get a list of words stored in the dictionary
'*******************************************************************************
Public Function GetDictionaryList() As String()
  Dim Str As String, Ary() As String
  Dim Node As dynNode
  Dim Index As Long, Count As Long, Idx As Long, Cnt As Long
  Dim I As Long, ArySize As Long

  If RootNode Is Nothing Then Exit Function   'nothing to do
  ArySize = 0
  
  RootNode.SortChildren True                  'sort everything
  With RootNode.Nodes
    Count = .Count                            'get count of letter items
    For Index = 1 To Count                    'parse Letter Children
      If CBool(StrComp("soundex", .Item(Index).Key, vbBinaryCompare)) Then
        With .Item(Index).Nodes               'process grandchildren to file
          For Idx = 1 To .Count                'process them all
            With .Item(Idx).Nodes             'parse word list
              Cnt = .Count                    'get number of words under this soundex key
              If CBool(Cnt) Then                'if words to process
                ReDim Preserve Ary(ArySize + Cnt - 1) 'resize array to hold new stuff
                For I = 1 To Cnt
                  Ary(ArySize) = .Item(I).Text 'grab actual word
                  ArySize = ArySize + 1       'bump array index
                Next I                        'do while word list
              End If
            End With
          Next Idx                            'do all grandchildren
        End With
      End If
    Next Index                                'loop through all children
  End With
  GetDictionaryList = Ary                     'return list of words
End Function

'*******************************************************************************
' Function Name     : SaveFile
' Purpose           : Save the file. If no Filename, then call SaveAs()
'*******************************************************************************
Public Function SaveFile() As Boolean
  If Len(TextFile) Then
    Dim Ts As TextStream
    Set Ts = Fso.OpenTextFile(TextFile, ForWriting, False)  'open file
    Ts.Write frmSpellCheck.rtbText.Text                       'write data
    Ts.Close                                                'close file
    MyCancel = False                                        'turn off cancel
    IsDirty = False                                         'no longer dirty
    frmSpellCheck.mnuSave.Enabled = False
    frmSpellCheck.mnuSaveAs.Enabled = False
    SaveFile = True
  Else
    SaveFile = SaveAs                                       'else SAVE AS...
  End If
  frmSpellCheck.mnuSave.Enabled = False
End Function

'*******************************************************************************
' Function Name     : SaveAs
' Purpose           : Save the file to a specified file
'*******************************************************************************
Public Function SaveAs() As Boolean
  Dim Ts As TextStream
  
  MyCancel = False
  With frmSpellCheck.CommonDialog1
    .Flags = cdlOFNLongNames Or cdlOFNExplorer Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    .DefaultExt = "txt"
    If Len(TextFile) = 0 Then
      .FileName = "Untitled"
    Else
      .FileName = TextFile
    End If
    .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    .DialogTitle = "Save Text File"
    .CancelError = True
    On Error Resume Next
    .ShowOpen
    If Err.Number Then Exit Function
    On Error GoTo 0
    If Len(Trim$(.FileName)) = 0 Then Exit Function
    TextFile = .FileName
    Set Ts = Fso.OpenTextFile(TextFile, ForWriting, True)
    Ts.Write frmSpellCheck.rtbText.Text
    Ts.Close
  End With
  frmSpellCheck.mnuSave.Enabled = False
  frmSpellCheck.mnuSaveAs.Enabled = False
  IsDirty = False
  SaveAs = True
End Function

'*******************************************************************************
' Subroutine Name   : LoadDictionary
' Purpose           : Load a dictionarty file into a dynNode tree
'*******************************************************************************
Public Sub LoadDictionary(Path As String)
  Dim Ts As TextStream
  Dim Str As String, Ary() As String, WordAry() As String, DctName As String
  Dim Index As Long, Idx As Long, Cnt As Long, I As Long, C As Long
  Dim Node As dynNode, SndxNode As dynNode
'
' show that we are busy
'
  Screen.MousePointer = vbHourglass
  DoEvents
  Call KillCurrentDictionaryData                'delete any current resources
  
'
' load file to Str string
'
  Set Ts = Fso.OpenTextFile(Path, ForReading, False)
  DctName = Ts.Read(7&)
'
' if not a dictionary file, then error
'
  If StrComp(DctName, "NODEDIC", vbBinaryCompare) <> 0 Then
    Ts.Close
    Screen.MousePointer = vbDefault
    CenterMsgBoxOnForm frmSpellCheck, "Invalid Dictionary File. Not Loaded", vbOKOnly Or vbExclamation, "Error"
    Exit Sub
  End If
'
' create a new dictionary
'
  Set RootNode = New dynNode  'create new root-level node
  RootNode.Init Path, Path    'initialize it
  RootNode.KeyChecks = False  'we know we will not run into duplicates...
'
' see if there is anything to add to the array
'
  If Not Ts.AtEndOfStream Then
    DctName = Path
    If InStrRev(DctName, "\") <> 0 Then DctName = Mid$(Path, InStrRev(Path, "\") + 1)
    If InStrRev(DctName, ".") <> 0 Then DctName = Left$(DctName, InStrRev(DctName, ".") - 1)
    'split the text into a string array
    Ary = Split(Ts.ReadAll, vbCrLf)
    Ts.Close
    'add each entry to the dictionary
    For Index = 0 To UBound(Ary)
      If Len(Ary(Index)) Then
        WordAry = Split(Ary(Index), "|")          'slit to array of words (0=Soundex)
        Str = WordAry(0)                              'get Soundex key
        'find the Letter node
        Set Node = LetterFromRoot(Left$(Str, 1))  'check if KEY for Character found
        'create it if it is not found
        If Node Is Nothing Then
          Set Node = RootNode.Nodes.Add(, dynNodeChild, Left$(Str, 1), Left$(Str, 1))
        End If                                    'create grandchild with Soundex
        'Find Soundex key from letter Node
        Set SndxNode = SoundexFromLetter(Node, Str) 'letter node already exists?
        If SndxNode Is Nothing Then                'no...
          Set SndxNode = Node.Nodes.Add(, dynNodeChild, Str, Str) 'so create it
        End If
        '
        ' now parse each word in the list and add nodes for them (they are unique)
        '
        With SndxNode.Nodes
          C = UBound(WordAry)
          For I = 1 To C
            Str = WordAry(I)
            .Add , dynNodeChild, , Str            'add word
            AddSoundexWordToDictionary Str        'add to standard SoundEx list
          Next I
        End With
      End If
    Next Index
'''    Unload frmMsg
  Else
    Ts.Close
  End If
'
' now clean house before leaving
'
  RootNode.KeyChecks = True                     'reset key checks to CHECK
  DicChanged = False                            'no changes to dictionary
  DicFile = Path                                'the path is valid
  SaveSetting App.Title, "Settings", "CurrentDictionary", Path
  frmSpellCheck.mnuCheck.Enabled = IsDirty
  Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************
' Function Name     : LetterFromRoot
' Purpose           : Find a node directly under RootNode with the specified Letter
'*******************************************************************************
Public Function LetterFromRoot(Letter As String) As dynNode
  Dim Index As Long, Count As Long
  Dim Node As dynNode
  
  With RootNode.Nodes
    Count = .Count                    'get chiild node count
    For Index = 1 To Count            'check all
      If StrComp(.Item(Index).Key, Letter) = 0 Then 'found letter?
        Set Node = .Item(Index)       'yes
        Exit For                      'no need to keep looking
      End If
    Next Index
    If Not Node Is Nothing Then Set LetterFromRoot = Node
  End With
End Function

'*******************************************************************************
' Function Name     : SoundexFromLetter
' Purpose           : Find a node with the specified Soundex value under the
'                   : letter node
'*******************************************************************************
Public Function SoundexFromLetter(LetterNode As dynNode, Sndx As String) As dynNode
  Dim Index As Long, Count As Long
  Dim Node As dynNode
  
  With LetterNode.Nodes
    Count = .Count                    'get chiild node count
    For Index = 1 To Count            'check all
      If StrComp(.Item(Index).Key, Sndx) = 0 Then 'found match?
        Set Node = .Item(Index)       'yes
        Exit For                      'no need to keep looking
      End If
    Next Index
    If Not Node Is Nothing Then Set SoundexFromLetter = Node
  End With
End Function

'*******************************************************************************
' Function Name     : QuickFindMatch
' Purpose           : This function will speed an already fast search by checking the
'                   : children directy under the root for the starting letter, and then
'                   : search for the soundex entry under than. Once found, the
'                   : Text of the found node will be checked for a match with the
'                   : search text and return TRUE if a match was found
'*******************************************************************************
Public Function QuickFindMatch(Text As String) As Boolean
  Dim Sndx As String, Test As String, C As String, Ary() As String
  Dim Index As Long, Count As Long
  Dim Node As dynNode
  
  Test = LCase$(Trim$(Text))                        'normalize the search text
  Sndx = GetSoundPlus(Test)                         'get its Soundex value
  C = Left$(Sndx, 1)                                'get Soundex letter (1st char)
'
' ensure the ltrRef and SndxRef node references are Nothing to begin with in
' we fail further up before redefinition, and old still lingers on...
'
  Set LtrRef = Nothing
  Set SndxRef = Nothing
'
' use this letter as the search key to heck all the nodes directly under the
' root node for the node that contains that key...
'
  With RootNode.Nodes
    Count = .Count
    For Index = 1 To Count
      If StrComp(.Item(Index).Key, C, vbBinaryCompare) = 0 Then Exit For
      If .Item(Index).Key = C Then Exit For
    Next Index
    If Index > Count Then Exit Function             'no key was matched
'
' we found the node. so search that node's children for a match of the Soundex
' value as a KEY...
'
    Set LtrRef = .Item(Index)                       'save node with matched Letter
  End With
  With LtrRef.Nodes
    Count = .Count
    For Index = 1 To Count
      If StrComp(.Item(Index).Key, Sndx, vbBinaryCompare) = 0 Then Exit For  'found the key!
    Next Index
    If Index > Count Then Exit Function             'key was NOT found
'
' we matched the key, so parse that node's children for the matching word
'
    Set SndxRef = .Item(Index)                      'save node with matched Soundex
  End With
  With SndxRef.Nodes
    Count = .Count
    For Index = 1 To Count
      If StrComp(Text, .Item(Index).Text, vbBinaryCompare) = 0 Then
        Set WordRef = .Item(Index)
        WordIndex = Index                           'save index into LtrRef list
        Exit For                                    'found a match
      End If
    Next Index
  End With
  
  QuickFindMatch = Index <= Count                   'TRUE if we found a matched word
End Function

'*******************************************************************************
' Function Name     : QuickFindSndX
' Purpose           : Find the SoundEx code key
'*******************************************************************************
Public Function QuickFindSndX(Text As String) As Boolean
  Dim Sndx As String, Test As String, C As String, Ary() As String
  Dim Index As Long, Count As Long
  Dim Node As dynNode
'
' first see if we have a soundex value to work with
'
  Test = LCase$(Trim$(Text))                        'normalize the search text
  Sndx = GetSoundex(Test)                           'get its Soundex value
  If Len(Sndx) = 0 Then                             'no Soundex for this word?
    Sndx = UCase$(Left$(Test, 1)) & "000"           'no, so force one
  End If
  C = Left$(Sndx, 1) & "soundex"                    'else get Soundex letter (1st char)
'
' ensure the SoundexLtr and SoundexRef node references to Nothing to begin with in
' we fail further up before redefinition, and old still lingers on...
'
  Set SoundexLtr = Nothing
  Set SoundexRef = Nothing
'
' First find the main SoundEx if it is not yet defined...
'
  If SoundexNode Is Nothing Then
    With RootNode.Nodes
      Count = .Count
      For Index = 1 To Count
        If StrComp(.Item(Index).Key, "soundex", vbBinaryCompare) = 0 Then Exit For
      Next Index
      If Index > Count Then Exit Function             'no main SoundEx key was matched
      Set SoundexNode = .Item(Index)
    End With
  End If
'
' we have the SoundEx node. Search that node's children for a match of the Soundex
' letter as a KEY...
'
  If SoundexLtr Is Nothing Then
    With SoundexNode.Nodes
      Count = .Count
      For Index = 1 To Count
        If StrComp(.Item(Index).Key, C, vbBinaryCompare) = 0 Then Exit For 'found!
      Next Index
      If Index > Count Then Exit Function               'no Letter key was matched
      Set SoundexLtr = .Item(Index)
    End With
  End If
'
' we have the SoundEx letter node. Search that node's children for a match of the
' Soundex value as a KEY...
'
  If SoundexRef Is Nothing Then
    With SoundexLtr.Nodes
      Count = .Count
      For Index = 1 To Count
        If StrComp(.Item(Index).Key, Sndx, vbBinaryCompare) = 0 Then Exit For 'found!
      Next Index
      If Index > Count Then Exit Function               'no Letter key was matched
      Set SoundexRef = .Item(Index)
    End With
  End If
'
' return TRUE if we found a matched Soundex key
'
  QuickFindSndX = True
End Function

'*******************************************************************************
' Subroutine Name   : AddWordToDictionary
' Purpose           : Add a word to our dictionary
'*******************************************************************************
Public Sub AddWordToDictionary(NewWord As String)
  Dim Sndx As String, Test As String, C As String
  Dim Index As Long, Count As Long
'
' if the dictionary object is not loaded, then we can do nothing
'
  If RootNode Is Nothing Then Exit Sub        'nothing to do
'
' see if the word is already in the dictionary (safety net)
'
  If QuickFindMatch(NewWord) Then Exit Sub          'the word was found
'
' set up for exmploration
'
  Test = LCase$(Trim$(NewWord))                     'normalize the search text
  Sndx = GetSoundPlus(Test)                         'get its Soundex value
  C = Left$(Sndx, 1)                                'get Soundex letter (1st char)
'
' see if the Letter reference Node exists for the word
'
  If LtrRef Is Nothing Then                         'not found, so crate it
    Set LtrRef = RootNode.Nodes.Add(, dynNodeChild, Left$(Sndx, 1), Left$(Sndx, 1))
    DoEvents
  End If
'
' see if the Soundex reference exists
'
  If SndxRef Is Nothing Then                        'not found, so create it
    Set SndxRef = LtrRef.Nodes.Add(, dynNodeChild, Sndx, Sndx)
    DoEvents
  End If
'
' now add the actual word to the dictionary
'
  Call SndxRef.Nodes.Add(, dynNodeChild, , Test)
'
' see if we can add the standard Soundex value as well
'
  AddSoundexWordToDictionary Test
'
' indicate that the dictionary has been altered
'
  DicChanged = True
End Sub

'*******************************************************************************
' Subroutine Name   : AddSoundexWordToDictionary
' Purpose           : Add a word to the SoundEx list
'*******************************************************************************
Public Sub AddSoundexWordToDictionary(NewWord As String)
  Dim Sndx As String, C As String

  Sndx = GetSoundex(NewWord)                      'so get its Soundex value
  If Len(Sndx) = 0 Then                           'no Soundex for this word?
    Sndx = UCase$(Left$(NewWord, 1)) & "000"      'no, so force one
  End If
  If Not QuickFindSndX(NewWord) Then              'SoundEx key exists?
    C = Left$(Sndx, 1) & "soundex"                'else get Soundex letter (1st char)
'
' did not find Soundex key, so create SoundexNode if it is not defined...
'
  If SoundexNode Is Nothing Then                    'SoundEx node defined?
    Set SoundexNode = RootNode.Nodes.Add(, dynNodeChild, "soundex", "SoundEx")
  End If
'
' if Soundex letter does not exist, create it as well..
'
  If SoundexLtr Is Nothing Then
    Set SoundexLtr = SoundexNode.Nodes.Add(, dynNodeChild, C, C)
  End If
'
' then create SoundEx key...
'
    Set SoundexRef = SoundexLtr.Nodes.Add(, dynNodeChild, Sndx, Sndx)
  End If
'
' now create an entry entry for teh actual word
'
  Call SoundexRef.Nodes.Add(, dynNodeChild, , NewWord)
End Sub

'*******************************************************************************
' Subroutine Name   : KillCurrentDictionaryData
' Purpose           : Release any existing dictionary resources
'*******************************************************************************
Public Sub KillCurrentDictionaryData()
  WordIndex = 0
  Set WordRef = Nothing
  Set SndxRef = Nothing
  Set LtrRef = Nothing
  Set SoundexLtr = Nothing
  Set SoundexRef = Nothing
  Set SoundexNode = Nothing
  If Not RootNode Is Nothing Then RootNode.Nodes.Clear
  Set RootNode = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : SortChildren
' Purpose           : Shell/Metzner Sort the child nodes (VERY FAST)
'*******************************************************************************
Public Sub SortList(SortArray() As String, Optional ReverseOrder As Boolean = False)
  Dim IndexLo As Long, IncIndex As Long, Count As Long
  Dim HalfUp As Long, IndexHi As Long
  Dim HalfDown As Long, NumberofItems As Long
  Dim CompFlag As Integer
  Dim Tmp As String
'
' if not enough items to sort, then nothing to do
'
  On Error Resume Next
  Count = UBound(SortArray) + 1           'get number of items
  If Count = 0 Then Exit Sub              'no need if we have 0
  On Error GoTo 0
'
' if we have only one child, it is not required to sort the current collection
' but we will try sorting the single child node
'
  If Count > 1 Then                       'if more than one child
'
' set up comparison flag for either Descending or Ascending sort order
'
    If ReverseOrder Then
      CompFlag = -1                       'sort in reverse order
    Else
      CompFlag = 1                        'sort in ascending order (Default)
    End If
'
' sort initialization                                                         Original Algorythm (1-Based)
'                                                                             ----------------------------
    NumberofItems = Count                 'get number if items to sort            (N=Number of Items)
    HalfDown = NumberofItems              'number of items to sort                (M=N)
'
' perform the sort
'
    Do While CBool(HalfDown \ 2)          'while counter can be halved        A:  IF(M\2)=0 THEN STOP
      HalfDown = HalfDown \ 2             'back down by 1/2                       (M=M\2)
      HalfUp = NumberofItems - HalfDown   'look in upper half                     (K=N-M)
      IncIndex = 0                        'init index to start of array           (J=1)
      Do While IncIndex < HalfUp          'do while we can index range
        IndexLo = IncIndex                'set base                           B:  I=J
        Do
          IndexHi = IndexLo + HalfDown  'if (IndexLo) > (IndexHi), then swap  C:  L=I+M
          If StrComp(SortArray(IndexLo), _
                     SortArray(IndexHi), vbTextCompare) = CompFlag Then       '   IF D(I)>D(L) THEN GOTO D
            Tmp = SortArray(IndexLo) 'swap nodes                                  T=D(I)
            SortArray(IndexLo) = SortArray(IndexHi)                           '   D(I)=D(L)
            SortArray(IndexHi) = Tmp                                          '   D(L)=T
            IndexLo = IndexLo - HalfDown  'back up index                      '   I=I-M
          Else                                                                '   IF I>=1 THEN GOTO C
            IncIndex = IncIndex + 1       'else bump counter                   D: J=J+1
            Exit Do                                                           '   IF J>K THEN GOTO A
          End If                                                              '   GOTO B
        Loop While IndexLo >= 0            'while more things to check
      Loop
    Loop
  End If
End Sub

