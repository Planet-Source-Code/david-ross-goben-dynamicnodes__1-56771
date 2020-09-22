VERSION 5.00
Begin VB.Form frmCorrect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spell Check"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReview 
      Caption         =   "Review Unknown Words"
      Height          =   315
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Review all words in the text that are not in the dictionary"
      Top             =   3420
      Width           =   2235
   End
   Begin VB.CommandButton cmdAddAll 
      Caption         =   "Add ALL to &Dictionary"
      Height          =   315
      Left            =   3240
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Add all unfound words to the dictionary"
      Top             =   3420
      Width           =   1995
   End
   Begin VB.TextBox txtFix 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1740
      Width           =   3435
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Cancel spell checking"
      Top             =   3420
      Width           =   1995
   End
   Begin VB.CommandButton cmdChangeAll 
      Caption         =   "Change Al&l"
      Height          =   315
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Change all subsequent encounters with this word to the selected text"
      Top             =   2460
      Width           =   1995
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Default         =   -1  'True
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Change the word to the selected text"
      Top             =   2040
      Width           =   1995
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add to Dictionary"
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      ToolTipText     =   "Add this word to the dictionary"
      Top             =   1260
      Width           =   1995
   End
   Begin VB.CommandButton cmdIgnoreAll 
      Caption         =   "I&gnore All"
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Ignore this and all subsequent encounters with this word"
      Top             =   810
      Width           =   1995
   End
   Begin VB.CommandButton cmdIngnoreOnce 
      Caption         =   "&Ignore Once"
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   "Ignore this word just once"
      Top             =   360
      Width           =   1995
   End
   Begin VB.ListBox lstSuggest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   5055
   End
   Begin VB.TextBox txtHidefocus 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Lock"
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtSample 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   5115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suggestions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   11
      Top             =   1800
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Not in Dictionary:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "frmCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectionStart As Long     'start of highlighted text in sample
Public SelectionLength As Long    'of of selection text in sample
Public SelFromUser As Boolean     'true when the user types something
'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Init form layout
'*******************************************************************************
Private Sub Form_Load()
  Me.Icon = frmSpellCheck.Icon
  Me.txtHidefocus.Left = -2880  'hide the locked textbox that gets focus when the
                                'user selects the displayed locked textbox
  Me.cmdChange.Enabled = CBool(Me.lstSuggest.ListCount)
  Me.cmdChangeAll.Enabled = Me.cmdChange.Enabled
  Me.cmdChange.Default = True
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : The user is closing the window using the X button
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then CorrectFlag = -1
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Cancel the spell check
'*******************************************************************************
Private Sub cmdCancel_Click()
  CorrectFlag = -1
End Sub

'*******************************************************************************
' Subroutine Name   : cmdIngnoreOnce_Click
' Purpose           : Ignore this word this once
'*******************************************************************************
Private Sub cmdIngnoreOnce_Click()
  CorrectFlag = 1
  Me.cmdIngnoreOnce.Default = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdIgnoreAll_Click
' Purpose           : Ignore all subsequent encounters with this word
'*******************************************************************************
Private Sub cmdIgnoreAll_Click()
  CorrectFlag = 2
  Me.cmdIgnoreAll.Default = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAdd_Click
' Purpose           : Add the unknown word to teh dictionary
'*******************************************************************************
Private Sub cmdAdd_Click()
  CorrectFlag = 3
  Me.cmdAdd.Default = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdChange_Click
' Purpose           : The user wants this word to be changed to the selected text
'*******************************************************************************
Private Sub cmdChange_Click()
  ChangeChangeAll 4
End Sub

'*******************************************************************************
' Subroutine Name   : ChangeChangeAll
' Purpose           : Support change and change all
'*******************************************************************************
Private Sub ChangeChangeAll(Cmd45 As Long)
  Dim Text As String, Ary() As String, Test As String, Word As String
  Dim Index As Long, Count As Long, I As Long, Words As Long, WordSt As Long
  Dim Inword As Boolean
  
  If SelFromUser Then                             'if the user entered something
    Text = Trim$(Me.txtFix.Text)                  'get text
'
' break it down into an array of potentially separate words
'
    Test = " " & Punctuation
    Count = Len(Text)                             'length of text file
    Inword = False                                'init not being inside a word
    I = 0
    Words = 0
    Do While I < Count
      I = I + 1
      If InStr(1, Test, Mid$(Text, I, 1)) Then    'a pinctuation character?
        If Inword Then                            'yes, are we in a word?
          Word = Mid$(Text, WordSt, I - WordSt)   'yes, so grab full word
          Inword = False                          'and mark no longer in a word
          ReDim Preserve Ary(Words) As String
          Ary(Words) = Word
          Words = Words + 1
        End If
      ElseIf Not Inword Then
        WordSt = I                                    'start new word
        Inword = True                                 'indicate we know where we are
      End If
    Loop
    
    If Inword Then                            'yes, are we in a word?
      Word = Mid$(Text, WordSt, I - WordSt + 1) 'yes, so grab full word
      Inword = False                          'and mark no longer in a word
      ReDim Preserve Ary(Words) As String
      Ary(Words) = Word
      Words = Words + 1
    End If
'
' if we have a list of words, we will remove from this array any that exist
' in the dictionary
'
    Count = 0                                         'init count for words to add
    If Words Then                                     'if words exist
      Test = vbNullString                             'init message list
      For Index = 0 To Words - 1                      'parse each word(s)
        If QuickFindMatch(Ary(Index)) Then            'in dictionary?
          Ary(Index) = vbNullString 'word was found, so purge from list
        Else
          Count = Count + 1                           'count an unfound word
          Test = Test & ", " & Ary(Index)             'add word to report
        End If
      Next Index
    End If
'
' if any words were not found, ask the user if they want to add them to the dictionary
'
    If Count Then                                     'words are still in list
      Test = Mid$(Test, 3)                            'strip initial ", "
      If Count = 1 Then                               'only 1 word to add
        Test = "The following word was not found in the dictionary:" & vbCrLf & Test & vbCrLf & vbCrLf & "Add it to the dictionary?"
      Else                                            'more than one word
        Test = "The following words were not found in the dictionary:" & vbCrLf & Test & vbCrLf & vbCrLf & "Add them to the dictionary?"
      End If
'
' prompt the user for a response. Cancel remains in the form
'
      Select Case CenterMsgBoxOnForm(Me, Test, vbYesNoCancel, "Add To Dictionary?")
        Case vbNo                                     'do not add, just replace
          CorrectFlag = Cmd45                         'set 4 or 5
        Case vbYes
          For Index = 0 To Words - 1                  'add all words
            Word = Ary(Index)                         'grab a word
            If CBool(Len(Word)) Then                  'still exists?
              AddWordToDictionary Word                'add to dictionary if so
            End If
          Next Index                                  'parse all words
          CorrectFlag = Cmd45                         'set command 4 or 5
        Case Else
          Me.txtFix.SetFocus                          'CANCEL. So remain in form
          Exit Sub
      End Select
    End If
'
' here when the user has not typed a replacement
'
  Else
    CorrectFlag = Cmd45
  End If
  Me.cmdChange.Default = True   'never see change all as a default button
End Sub

'*******************************************************************************
' Subroutine Name   : cmdChangeAll_Click
' Purpose           : The user wants all subsequent encountered with this word
'                   : to be changed to the selected text
'*******************************************************************************
Private Sub cmdChangeAll_Click()
  ChangeChangeAll 5
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAddAll_Click
' Purpose           : The user wants to simply add all unfound words to the dictionary
'*******************************************************************************
Private Sub cmdAddAll_Click()
  If CenterMsgBoxOnForm(Me, "Are you SURE that you want to do this? You should only use this command" & vbCrLf & _
            "when you are positive that all words are correct. You may wish to first" & vbCrLf & _
            "browse the unknown word list first before answering YES. Go ahead?", _
            vbYesNo Or vbQuestion, "Add All Verification") = vbNo Then Exit Sub
  CorrectFlag = 6
End Sub

'*******************************************************************************
' Subroutine Name   : cmdReview_Click
' Purpose           : The user wants to review unknown words
'*******************************************************************************
Private Sub cmdReview_Click()
  CorrectFlag = 7
End Sub

'*******************************************************************************
' Subroutine Name   : lstSuggest_Click
' Purpose           : Something seelcted in the list
'*******************************************************************************
Private Sub lstSuggest_Click()
  Dim Str As String
'
' ignore if the selected text is a "(no suggestions)" note
'
  Str = Me.lstSuggest.List(Me.lstSuggest.ListIndex)
  If Left$(Str, 1) = "(" Then Exit Sub 'a no suggestion, so skip out
  Me.txtFix.Text = Str                  'else set the change text
  SelFromUser = False                   'selection from list
End Sub

'*******************************************************************************
' Subroutine Name   : txtFix_Change
' Purpose           : When the text field is altered, update the Change and
'                   : Change All buttons as required.
'*******************************************************************************
Private Sub txtFix_Change()
  Me.cmdChange.Enabled = CBool(Len(Me.txtFix.Text))
  Me.cmdChangeAll.Enabled = Me.cmdChange.Enabled
  Me.cmdChange.Default = Me.cmdChange.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : txtFix_GotFocus
' Purpose           : Highlight the entire text field when the textbox is selected
'*******************************************************************************
Private Sub txtFix_GotFocus()
  With Me.txtFix
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : txtFix_KeyPress
' Purpose           : User typed something
'*******************************************************************************
Private Sub txtFix_KeyPress(KeyAscii As Integer)
  SelFromUser = True
End Sub

'*******************************************************************************
' Subroutine Name   : txtSample_GotFocus
' Purpose           : When the user click on the sample text, the highlight will
'                   : be affected. This will reset the highlight and move the
'                   : focus away from the sample
'*******************************************************************************
Private Sub txtSample_GotFocus()
  Me.txtHidefocus.SetFocus
  With Me.txtSample
    .SelStart = Me.SelectionStart
    .SelLength = Me.SelectionLength
  End With
End Sub
