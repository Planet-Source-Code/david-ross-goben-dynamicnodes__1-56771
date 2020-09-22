VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditDictionary 
   Caption         =   "Edit Dictionary Word List"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picRight 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   6960
      ScaleHeight     =   1455
      ScaleWidth      =   2115
      TabIndex        =   8
      Top             =   4440
      Width           =   2115
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "Reset the word list and all choices"
         Top             =   0
         Width           =   2115
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Accept changes and close this dialog"
         Top             =   540
         Width           =   2115
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "Cancel all changes and close this dialog"
         Top             =   1080
         Width           =   2115
      End
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   4440
      Width           =   6675
      Begin VB.CommandButton cmdFindExact 
         Caption         =   "Find Exact Match..."
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Find an exact match for a word"
         Top             =   1020
         Width           =   2115
      End
      Begin VB.CommandButton cmdFindMatch 
         Caption         =   "Find Match..."
         Height          =   375
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Find the nearest match for a word"
         Top             =   1020
         Width           =   2115
      End
      Begin VB.CommandButton cmdInvert 
         Caption         =   "In&vert All Selections"
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         ToolTipText     =   "Invert the checks of all items in the list"
         Top             =   540
         Width           =   2115
      End
      Begin VB.CommandButton cmdUnselectAll 
         Caption         =   "&Unselect All"
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "Uncheck all items in the list"
         Top             =   540
         Width           =   2115
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Check all items in the list"
         Top             =   540
         Width           =   2115
      End
      Begin VB.CommandButton cmdClipboard 
         Caption         =   "Save Checked to Clipboard"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         ToolTipText     =   "Dave checked items to the clipboard"
         Top             =   0
         Width           =   2115
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Checked Word(s)..."
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "Remove all checked words from the dictionary"
         Top             =   0
         Width           =   2115
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Word(s)..."
         Height          =   375
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Add a word to the dictionary"
         Top             =   0
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Double-click word to place copy in the clipoboard"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   4620
         TabIndex        =   14
         Top             =   1020
         Width           =   1995
      End
   End
   Begin MSComctlLib.ListView lvwWords 
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5847
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblClicked 
      AutoSize        =   -1  'True
      Caption         =   "Clicked Item"
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
      Left            =   120
      TabIndex        =   15
      Top             =   5820
      Width           =   1065
   End
End
Attribute VB_Name = "frmEditDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ColAdd As Collection        'collection to store words to be added to the dict.
Dim ColDelete As Collection     'collection to store words to be removed from the dict.
Dim OrgTitle As String          'original title of the form

Private Sub Form_Load()
  OrgTitle = Me.Caption
  
  Me.lblClicked.Caption = vbNullString
  Me.lblClicked.ForeColor = vbBlue
  Set ColAdd = New Collection
  Set ColDelete = New Collection
  Call Reset
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resizing the form. SHape everything to fit
'*******************************************************************************
Private Sub Form_Resize()
  Static Resizing As Boolean
  
  If Resizing Then Exit Sub
  Resizing = True
  If Me.Width < 9270 Then Me.Width = 9270
  If Me.Height < 4000 Then Me.Height = 4000
  Me.lblClicked.Top = Me.ScaleHeight - Me.lblClicked.Height
  Me.picRight.Top = Me.ScaleHeight - Me.picRight.Height - 120
  Me.picRight.Left = Me.ScaleWidth - Me.picRight.Width - 120
  Me.lvwWords.Width = Me.ScaleWidth - Me.lvwWords.Left * 2
  Me.lvwWords.Height = Me.picRight.Top - Me.lvwWords.Top - 120
  Me.picLeft.Top = Me.lvwWords.Top + Me.lvwWords.Height + 120
  Resizing = False
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Remove created resources when leaving
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set ColAdd = Nothing
  Set ColDelete = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : cmdAdd_Click
' Purpose           : Add word(s) to dictionary
'*******************************************************************************
Private Sub cmdAdd_Click()
  Dim Str As String, Ary() As String, Words As String
  Dim Index As Long, Count As Long, NewIdx As Long
  
  Str = Trim$(InputBox("Enter word(s) to add. Separate multiple" & vbCrLf & _
                       "words with spaces or commas:", "Add Word(s) To Dictionary", vbNullString))
  If Len(Str) = 0 Then Exit Sub
  Ary = SplitText(Str)
  On Error Resume Next
  Count = UBound(Ary)
  If Err.Number = 0 Then
    Ary = GetUniqueWords(Ary)               'reduce to unique
    On Error Resume Next                    'don't explode if array not dimmed...
    Count = UBound(Ary)                     'get upper bounds
'
' now go through and gather the list of new words
'
    If Err.Number = 0 Then
      NewIdx = -1                               'init ubound on NEW words
      ReDim NewWords(Count) As String
      Words = vbNullString
      For Index = 0 To Count
        If Not QuickFindMatch(Ary(Index)) Then  'found this word in the dictionary?
          NewIdx = NewIdx + 1                   'no, so add it to the list
          On Error Resume Next                  'in case user added repeats...
          ColAdd.Add Ary(Index)
          Me.lvwWords.ListItems.Add , Ary(Index), Ary(Index)
          If Err.Number = 0 Then Words = Words & ", " & Ary(Index)
        End If
      Next Index
      On Error GoTo 0
'
' if no new words exist...
'
      If NewIdx = -1 Then
        Screen.MousePointer = vbDefault
        CenterMsgBoxOnForm frmSpellCheck, "No NEW words to Import.", vbOKOnly Or vbInformation, "No New Words"
        Exit Sub
      Else
        Str = CStr(NewIdx + 1) & " word"
        If CBool(NewIdx) Then Str = Str & "s"
        CenterMsgBoxOnForm frmSpellCheck, Str & " will be added when you select OK to exit this window." & vbCrLf & vbCrLf & _
               "Word(s) to be added: " & Mid$(Words, 3) & vbCrLf & vbCrLf & _
               "Total words added to pending list so far: " & CStr(ColAdd.Count), vbOKOnly Or vbInformation, "Words Added"
        Me.cmdReset.Enabled = True
      End If
    End If
  End If
  Call UpdateCount
  Me.cmdOK.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClipboard_Click
' Purpose           : Save checked items to the clipboard
'*******************************************************************************
Private Sub cmdClipboard_Click()
  Dim Itm As ListItem
  Dim Str As String
  
  Str = vbNullString                                  'init list
  For Each Itm In Me.lvwWords.ListItems
    If Itm.Checked Then Str = Str & Itm.Text & vbCrLf 'add checked items to it
  Next Itm
  If CBool(Len(Str)) Then                             'if something to save
    Clipboard.Clear                                   'save it
    Clipboard.SetText Str
    CenterMsgBoxOnForm frmSpellCheck, "Checked items saved to the clipboard.", vbOKOnly Or vbInformation, "Clipboard"
  Else
    CenterMsgBoxOnForm frmSpellCheck, "No Checked items to save to the clipboard.", vbOKOnly Or vbExclamation, "Clipboard"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDelete_Click
' Purpose           : Delete checked items
'*******************************************************************************
Private Sub cmdDelete_Click()
  Dim Index As Long
  
  With Me.lvwWords.ListItems
    For Index = .Count To 1 Step -1
      If .Item(Index).Checked Then                'item checked?
        ColDelete.Add .Item(Index).Text           'add to delete list
        Me.lvwWords.ListItems.Remove Index        'remove from listview
      End If
    Next Index
  End With
  Me.cmdReset.Enabled = True
  Me.cmdOK.Enabled = True
  Call UpdateCount
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSelectAll_Click
' Purpose           : Mark all items with a check
'*******************************************************************************
Private Sub cmdSelectAll_Click()
  Dim Itm As ListItem
  
  For Each Itm In Me.lvwWords.ListItems
    Itm.Checked = True                  'ensure each item is checked
  Next Itm
  Call UpdateCount
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUnselectAll_Click
' Purpose           : Remove checks from all items
'*******************************************************************************
Private Sub cmdUnselectAll_Click()
  Dim Itm As ListItem
  
  For Each Itm In Me.lvwWords.ListItems
    Itm.Checked = False                   'ensure each item is unchecked
  Next Itm
  Call UpdateCount
End Sub

'*******************************************************************************
' Subroutine Name   : cmdInvert_Click
' Purpose           : Invert checks for all items
'*******************************************************************************
Private Sub cmdInvert_Click()
  Dim Itm As ListItem
  
  For Each Itm In Me.lvwWords.ListItems
    Itm.Checked = Not Itm.Checked         'ensure each items check is inverted
  Next Itm
  Call UpdateCount
End Sub

'*******************************************************************************
' Subroutine Name   : cmdReset_Click
' Purpose           : Reset all items in the list and the collections
'*******************************************************************************
Private Sub cmdReset_Click()
  Call Reset                              'reset any changes made
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFindExact_Click
' Purpose           : Find an exact match for a word
'*******************************************************************************
Private Sub cmdFindExact_Click()
  Dim Str As String, Tmp As String
  Dim Slen As Long
  Dim Itm As ListItem
  
  Str = LCase$(Trim$(InputBox("Enter word (or partial) to find in the list:", "Find Match", vbNullString)))
  Slen = Len(Str)
  If Slen = 0 Then Exit Sub
  For Each Itm In Me.lvwWords.ListItems
    Tmp = Itm.Text
    If Len(Tmp) = Slen Then
      If StrComp(Str, Tmp, vbBinaryCompare) = 0 Then
        Itm.Selected = True
        Itm.EnsureVisible
        Set Me.lvwWords.SelectedItem = Itm
        Me.lvwWords.SetFocus
        Exit Sub
      End If
    End If
  Next Itm
  CenterMsgBoxOnForm Me, "Did not find an exact match for: " & Str, vbOKOnly Or vbExclamation, "Word not found"
End Sub

'*******************************************************************************
' Subroutine Name   : cmdFindMatch_Click
' Purpose           : Find nearest match for word
'*******************************************************************************
Private Sub cmdFindMatch_Click()
  Dim Str As String, Tmp As String
  Dim Slen As Long
  Dim Itm As ListItem
  
  Str = LCase$(Trim$(InputBox("Enter word (or partial) to find in the list:", "Find Match", vbNullString)))
  Slen = Len(Str)
  If Slen = 0 Then Exit Sub
  For Each Itm In Me.lvwWords.ListItems
    Tmp = Itm.Text
    If Len(Tmp) >= Slen Then
      If StrComp(Str, Left$(Tmp, Slen), vbBinaryCompare) = 0 Then
        Itm.Selected = True
        Itm.EnsureVisible
        Set Me.lvwWords.SelectedItem = Itm
        Me.lvwWords.SetFocus
        Exit Sub
      End If
    End If
  Next Itm
  CenterMsgBoxOnForm Me, "Did not find a match for: " & Str, vbOKOnly Or vbExclamation, "Word not found"
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Add words to add, delete words to delete
'*******************************************************************************
Private Sub cmdOK_Click()
  Dim Text As String, Key As String, Report As String, Words As String
  Dim Index As Long, ChangeCount As Long
  Dim Node As dynNode, pNode As dynNode
'
' show busy
'
  Screen.MousePointer = vbHourglass
  DoEvents
'
' check for adding words to the dictionary
'
  Report = vbNullString
  With ColAdd
    If CBool(.Count) Then               'if words in ADD list
      Report = CStr(.Count) & " New word"
      If .Count <> 1 Then Report = Report & "s"
      Words = vbNullString
      Do While .Count                   'loop through them all
        AddWordToDictionary .Item(1)    'add each word in the list
        Words = Words & ", " & .Item(1)
        .Remove 1
      Loop
      Report = Report & " added to the dictionary:" & vbCrLf & Mid$(Words, 3) & vbCrLf
      DicChanged = True
    End If
  End With
'
' to delete, we will delete items and clean house if this item was the last in its
' pool. We really should also do this with the Soundex keys as well, but...
'
  With ColDelete
    If CBool(.Count) Then                                   'items to delete?
      Words = CStr(.Count) & " Word"
      If .Count <> 1 Then Words = Words & "s"
      If Len(Report) > 0 Then
        Report = Report & vbCrLf & Words
      Else
        Report = Words
      End If
    
      Words = vbNullString
      Do While .Count                                       'yes, loop through them
        Text = .Item(1)                                     'get an item
        If QuickFindMatch(Text) Then                        'found key?
          Words = Words & ", " & .Item(1)
          Set Node = WordRef                                'set local node to it
'
' we will do looping here in case the removed item was the last item in its parent's
' list. In this case well will then work our way up, and if each successive parent,
' until the root, has only one child, we will clean house by removing them.
'
          Do                                                'loop for possible houecleaning
            Key = Node.Key                                  'get key for item to remove
            Set pNode = Node.Parent                         'get its parent
            With pNode.Nodes
              For Index = 1 To .Count                       'find its index
                If StrComp(Key, .Item(Index).Key, vbBinaryCompare) = 0 Then Exit For
              Next Index
              .Remove Index                                 'remove items from dict.
              If CBool(.Count) Then Exit Do                 'was not last child
              If pNode.Parent Is Nothing Then Exit Do       'if pnode is root
              If pNode.Parent.Nodes.Count > 1 Then Exit Do  'its paranet has children
              Set Node = pNode                              'set we will remove parent
            End With
          Loop
        End If
        .Remove 1                                           'remove entry from col.
      Loop                                                  'wdo while list
      Report = Report & " removed from the dictionary:" & vbCrLf & Mid$(Words, 3)
      DicChanged = True
    End If
  End With
'
' no longer busy
'
  Screen.MousePointer = vbDefault
'
' issue a report if anything was done
'
  If Len(Report) <> 0 Then CenterMsgBoxOnForm frmSpellCheck, Report, vbOKOnly Or vbInformation, "Edit Action Report"
  
  Unload Me                                                 'all done
End Sub

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : Exit with no action
'*******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : lvwWords_Click
' Purpose           : enable OK when first item checked
'*******************************************************************************
Private Sub lvwWords_Click()
  Me.lblClicked.Caption = "Selected Word: " & Me.lvwWords.SelectedItem.Text
  Call UpdateButtons
End Sub

'*******************************************************************************
' Subroutine Name   : Reset
' Purpose           : Reset the lists
'*******************************************************************************
Private Sub Reset()
  Dim Ary() As String
  Dim Index As Long, Count As Long
  
  With ColAdd
    Do While .Count
      .Remove 1
    Loop
  End With
  
  With ColDelete
    Do While .Count
      .Remove 1
    Loop
  End With
  
  With Me.lvwWords.ListItems
    Do While .Count
      .Remove 1
    Loop
  End With
  
  Me.Caption = OrgTitle
  Me.cmdOK.Enabled = False
  Me.cmdDelete.Enabled = False
  Me.cmdClipboard.Enabled = False
  Me.cmdReset.Enabled = False
  Me.cmdFindMatch.Enabled = False
  Me.cmdFindExact.Enabled = False
'
' now build list an display it
'
  Ary = GetDictionaryList()             'grab the dictionary list
  On Error Resume Next
  Count = UBound(Ary)                   'get the number of entries
  If Err.Number = 0 Then
    Call SortList(Ary)
    With Me.lvwWords.ListItems
      For Index = 0 To Count            'last is always empty
        .Add , Ary(Index), Ary(Index)   'add all words
      Next Index
    End With
    Me.lvwWords.ListItems(0).Selected = True  'mark first entry
    Me.lvwWords.ListItems(0).EnsureVisible
  End If
  Call UpdateCount                      'update the counter
End Sub

'*******************************************************************************
' Subroutine Name   : UpdateCount
' Purpose           : Update items counts. Process selection enablement
'*******************************************************************************
Private Sub UpdateCount()
  Dim Itm As ListItem
  Dim Bol As Boolean
  
  Me.Caption = OrgTitle & " - Words in Dictionary: " & CStr(Me.lvwWords.ListItems.Count)
  Call UpdateButtons
End Sub

Private Sub UpdateButtons()
  Dim Itm As ListItem
  Dim Bol As Boolean

  If CBool(Me.lvwWords.ListItems.Count) Then
    For Each Itm In Me.lvwWords.ListItems
      If Itm.Checked Then
        Bol = True
        Exit For
      End If
    Next Itm
  End If
  Me.cmdClipboard.Enabled = Bol
  Me.cmdDelete.Enabled = Bol
  Me.cmdReset.Enabled = CBool(ColAdd.Count + ColDelete.Count)
  Bol = CBool(Me.lvwWords.ListItems.Count)
  Me.cmdFindMatch.Enabled = Bol
  Me.cmdFindExact.Enabled = Bol
  Me.cmdSelectAll.Enabled = Bol
  Me.cmdUnselectAll.Enabled = Bol
  Me.cmdInvert.Enabled = Bol
End Sub

'*******************************************************************************
' Subroutine Name   : lvwWords_DblClick
' Purpose           : Place copy of double-clicked word in clipboard
'*******************************************************************************
Private Sub lvwWords_DblClick()
  Clipboard.Clear
  Clipboard.SetText Me.lvwWords.SelectedItem.Text
End Sub
