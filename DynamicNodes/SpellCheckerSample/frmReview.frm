VERSION 5.00
Begin VB.Form frmReview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Review Unfound Word List"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Save to Clipboar&d"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Save the entire list to the clipboard"
      Top             =   4980
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3540
      TabIndex        =   9
      ToolTipText     =   "Remove all changes and close this form"
      Top             =   6480
      Width           =   1875
   End
   Begin VB.ListBox lstIgnore 
      Height          =   255
      Left            =   3540
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lstAdd 
      Height          =   255
      Left            =   3540
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   1620
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add / Remove"
      Height          =   1335
      Left            =   3480
      TabIndex        =   13
      Top             =   360
      Width           =   2115
      Begin VB.CommandButton cmdRemoveChecked 
         Caption         =   "&Ignore Checked"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Ignore the checked items from the unknown words list"
         Top             =   840
         Width           =   1875
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "Add All &Checked"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "All all checked items to the dictionary"
         Top             =   300
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection Options"
      Height          =   1815
      Left            =   3540
      TabIndex        =   12
      Top             =   2280
      Width           =   1995
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "Check all items in the list"
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton cmdUnselectAll 
         Caption         =   "&Unselect All"
         Height          =   375
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "Uncheck all items in the list"
         Top             =   780
         Width           =   1875
      End
      Begin VB.CommandButton cmdInvert 
         Caption         =   "In&vert All Selections"
         Height          =   375
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Invert the checks of all items in the list"
         Top             =   1320
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find a Word..."
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      ToolTipText     =   "Find a particular word in the list"
      Top             =   4440
      Width           =   1875
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3540
      TabIndex        =   8
      ToolTipText     =   "Accept all changes and close this form"
      Top             =   5940
      Width           =   1875
   End
   Begin VB.ListBox lstWords 
      Height          =   6585
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   300
      Width           =   3255
   End
   Begin VB.Label lblbWordCount 
      AutoSize        =   -1  'True
      Caption         =   "Word Count: 0"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   6960
      Width           =   1035
   End
   Begin VB.Label lblUnknown 
      AutoSize        =   -1  'True
      Caption         =   "Words not in dictionary:"
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
      Left            =   180
      TabIndex        =   10
      Top             =   60
      Width           =   2040
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  CorrectFlag = 0
  Unload Me
End Sub

Private Sub cmdClipboard_Click()
  Dim S As String
  Dim Index As Long
  
  With Me.lstWords
    For Index = 0 To .ListCount - 1
      S = S & .List(Index) & vbCrLf
    Next Index
  End With
  Clipboard.Clear
  Clipboard.SetText S
  CenterMsgBoxOnForm Me, "List saved to the clipboard", vbOKOnly Or vbInformation, "List Saved"
End Sub

Private Sub Form_Load()
  Me.Icon = frmSpellCheck.Icon
  Me.cmdAddAll.Enabled = False
  Me.cmdRemoveChecked.Enabled = False
End Sub

Private Sub cmdAddAll_Click()
  Dim Index As Long
  
  With Me.lstWords
    If .SelCount Then
      For Index = .ListCount - 1 To 0 Step -1
        If .Selected(Index) Then
          Me.lstAdd.AddItem .List(Index)
          .RemoveItem Index
        End If
      Next Index
      .ListIndex = -1
    End If
  End With
  Call FixButtons
End Sub

Private Sub FixButtons()
  Me.cmdAddAll.Enabled = CBool(Me.lstWords.ListCount)
  Me.cmdRemoveChecked.Enabled = Me.cmdAddAll.Enabled
  Me.cmdClipboard.Enabled = Me.cmdAddAll.Enabled
  Me.cmdSelectAll.Enabled = Me.cmdAddAll.Enabled
  Me.cmdUnselectAll.Enabled = Me.cmdAddAll.Enabled
  Me.cmdInvert.Enabled = Me.cmdAddAll.Enabled
End Sub

Private Sub cmdRemoveChecked_Click()
  Dim Index As Long
  
  With Me.lstWords
    If .SelCount Then
      For Index = .ListCount - 1 To 0 Step -1
        If .Selected(Index) Then
          Me.lstIgnore.AddItem .List(Index)
          .RemoveItem Index
        End If
      Next Index
      .ListIndex = -1
    End If
  End With
  Call FixButtons
End Sub

Private Sub cmdSelectAll_Click()
  Dim Index As Long, Hold As Long
  
  With Me.lstWords
    Hold = .ListIndex
    For Index = 0 To .ListCount - 1
      .Selected(Index) = True
    Next Index
    .ListIndex = Hold
  End With
End Sub

Private Sub cmdUnselectAll_Click()
  Dim Index As Long, Hold As Long
  
  With Me.lstWords
    Hold = .ListIndex
    For Index = 0 To .ListCount - 1
      .Selected(Index) = False
    Next Index
    .ListIndex = Hold
  End With
End Sub

Private Sub cmdInvert_Click()
  Dim Index As Long, Hold As Long
  
  With Me.lstWords
    Hold = .ListIndex
    For Index = 0 To .ListCount - 1
      .Selected(Index) = Not .Selected(Index)
    Next Index
    .ListIndex = Hold
  End With
End Sub

Private Sub cmdFind_Click()
  Dim Str As String
  Dim Idx As Long
  
  Str = LCase$(Trim$(InputBox("Enter the word to find, or the first part of it:", "Find a Word", vbNullString)))
  If Len(Str) = 0 Then Exit Sub
  Idx = SendMessageByString(Me.lstWords.hwnd, LB_FINDSTRING, -1&, Str)
  If Idx < 0 Then
    CenterMsgBoxOnForm Me, "Enter text was not found, even partially.", vbOKOnly Or vbInformation, "Word Not Found"
  Else
    Me.lstWords.ListIndex = Idx
  End If
End Sub

Private Sub cmdClose_Click()
  CorrectFlag = 8
  Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then CorrectFlag = 0
End Sub

Private Sub List1_Click()
  Me.cmdAddAll.Enabled = CBool(Me.lstWords.SelCount)
  Me.cmdRemoveChecked.Enabled = Me.cmdAddAll.Enabled
End Sub
