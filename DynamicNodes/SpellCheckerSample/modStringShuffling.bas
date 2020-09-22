Attribute VB_Name = "modStringShuffling"
Option Explicit

'*******************************************************************************
' Function Name     : GetUniqueWords
' Purpose           : Split word list into an array of unique words. Remove
'                   : special punctuation
'*******************************************************************************
Public Function GetUniqueWords(WordList() As String) As String()
  Dim Ary() As String
  Dim Index As Long, Count As Long, I As Long
  Dim Col As Collection
  
  On Error Resume Next
  Count = UBound(WordList)        'get high cell number
  If Err.Number Then Exit Function
'
' create out collection object
'
  Set Col = New Collection
'
' now fill the collection with the words. By including the word as a key, the
' collection will generate an error and not add duplicate. Hence, we will take
' advanatage of the On Error Resume Next command to ignore those errors. The
' result will be a collection with all uniqur values
'
  On Error Resume Next            'ignore duplication keys (this way we get rid of dupes)
  For Index = 0 To Count
    I = Len(WordList(Index))                              'get length of word
    If I > 1 Then                                         'ignore 1-character data
      If Not IsNumeric(Left$(WordList(Index), 1)) Then    'and numeric data
        If Not IsNumeric(Right$(WordList(Index), 1)) Then
          Col.Add WordList(Index), WordList(Index)        'passed tests, so add it
        End If
      End If
    End If
  Next Index
'
' now see if we still have something to do
'
  With Col
    If .Count = 0 Then            'something in collection?
      Set Col = Nothing           'nope, so release collection object resources
      Exit Function               'and exit the function
    End If
'
' resize array for the number of unique entries
'
    ReDim Ary(.Count - 1) As String   'else set array size to file contents
'
' now fill the array with the entries
'
    Do While .Count
      Ary(.Count - 1) = .Item(1)  'get a unique word
      .Remove 1                   'collapse collection as we go
    Loop
  End With
  Set Col = Nothing               'release collection object resources
  GetUniqueWords = Ary            'and return it
End Function

'*******************************************************************************
' Function Name     : SplitText
' Purpose           : Split the text file into an array of unque words. Remove
'                   : special punctuation
'*******************************************************************************
Public Function SplitText(SrcText As String) As String()
  Dim Text As String, Ary() As String, Test As String, C As String
  Dim Index As Long, Count As Long, I As Long
'
' if the trimmed, lowercase version of the text
'
  Text = LCase$(Trim$(SrcText))
  If Len(Text) = 0 Then Exit Function   'nothing to do
'
' set up for removal of punctuation
'
  Test = Punctuation
  Count = Len(Test)               'length of test string
'
' remove punctuation
'
  For Index = 1 To Count
    C = Mid$(Test, Index, 1)      'get a punctuation character
    I = InStr(1, Text, C)         'see if the text contains at lease 1 of it
    Do While I
      Mid$(Text, I, 1) = " "      'remove it while one exists in the text
      I = InStr(I + 1, Text, C)
    Loop
  Next Index                      'process all special characters
'
' now trim the word list my remove double spaces
'
  Do While InStr(1, Text, "  ")   'has double spaces?
    Ary = Split(Text, "  ")       'split into an array, bounded on them
    Text = Join(Ary, " ")         'rebuild text, replace 2 with 1
  Loop                            'keep removing them
'
' all done. Now see if there is still something to do
'
  Text = Trim$(Text)
  If Len(Text) = 0 Then Exit Function
'
' there is, so split it one the single space separators, and get a count of them
'
  If CBool(InStr(1, Text, " ")) Then  'if multiple words present
    Ary = Split(Text, " ")            'break them up
    SortArray Ary                     'sort the array
  Else
    ReDim Ary(0)                      'make a single entry
    Ary(0) = Text                     'and put the word there
  End If
  SplitText = Ary                 'and return it
End Function

'*******************************************************************************
' Subroutine Name   : SortArray
' Purpose           : Sort an array of string into ascending order
'*******************************************************************************
Public Sub SortArray(Str() As String)
  Dim IndexLo As Long, IncIndex As Long
  Dim HalfUp As Long, IndexHi As Long
  Dim HalfDown As Long, NumberofItems As Long
  Dim Tmp As String
  
  NumberofItems = UBound(Str) + 1       'if number of strings to do (from 0)
  HalfDown = NumberofItems              'number of items to sort
  Do While HalfDown \ 2                 'while counter can be halved
    HalfDown = HalfDown \ 2             'back down by 1/2
    HalfUp = NumberofItems - HalfDown   'look in upper half
    IncIndex = 0                        'init index to start of array
    Do While IncIndex < HalfUp          'do while we can index range
      IndexLo = IncIndex                'set base
      Do
        IndexHi = IndexLo + HalfDown
        If StrComp(Str(IndexLo), Str(IndexHi), vbBinaryCompare) = 1 Then 'check strings
          Tmp = Str(IndexLo)         'swap strings
          Str(IndexLo) = Str(IndexHi)
          Str(IndexHi) = Tmp
          IndexLo = IndexLo - HalfDown  'back up index
        Else
          IncIndex = IncIndex + 1       'else bump counter
          Exit Do
        End If
      Loop While IndexLo >= 0           'while more things to check
    Loop
  Loop
End Sub

