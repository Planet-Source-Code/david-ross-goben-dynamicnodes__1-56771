VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Test dynNode Tree"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "Drive to test"
      Top             =   2820
      Width           =   1635
   End
   Begin VB.CheckBox chkExpand 
      Caption         =   "Expand Branches"
      Height          =   255
      Left            =   6180
      TabIndex        =   8
      ToolTipText     =   "Automatically expand all tree branches"
      Top             =   4620
      Width           =   1575
   End
   Begin VB.CheckBox chkFiles 
      Caption         =   "Include Files"
      Height          =   255
      Left            =   6180
      TabIndex        =   7
      ToolTipText     =   "Include files in test (it will take longer)"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   6240
      ScaleHeight     =   555
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   3480
      Width           =   1275
      Begin VB.OptionButton Option1 
         Caption         =   "Descending"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Build treeview in reverse alphabetical order"
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ascending"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Build treeview in alphabetical order"
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort Order"
      Height          =   975
      Left            =   6120
      TabIndex        =   3
      Top             =   3240
      Width           =   1635
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6180
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8295
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   14631
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test cNodes"
      Height          =   435
      Left            =   6120
      TabIndex        =   0
      ToolTipText     =   "Begin test"
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label lblSubFolderCount 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   8460
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
' Form-Level Variables
'-------------------------------------------------------------------------------
Private WithEvents RootNode As dynNode        'define new root node
Attribute RootNode.VB_VarHelpID = -1
Private FSO As FileSystemObject               'standard system I/O object
Private AccumD As Long                        'keep track of Directory counts
Private AccumF As Long                        'keep track of File counts

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Set aside FSO
'*******************************************************************************
Private Sub Form_Load()
  Set FSO = New FileSystemObject              'instantiate FSO object
  Me.Picture1.BorderStyle = 0                 'hide border on container
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize the form
'*******************************************************************************
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub 'do nothing if minimizing
  If Me.Height < 3000 Then Me.Height = 3000     'do not get too small
  If Me.Width < 5000 Then Me.Width = 5000
  Me.lblSubFolderCount.Top = Me.ScaleHeight - Me.lblSubFolderCount.Height - 60
  Me.cmdTest.Left = Me.ScaleWidth - Me.cmdTest.Width - 180
  Me.cmdTest.Top = Me.ScaleHeight - Me.cmdTest.Height - 120
  Me.Frame1.Left = Me.cmdTest.Left - 60
  Me.Picture1.Left = Me.Frame1.Left + 45
  Me.chkExpand.Left = Me.Frame1.Left
  Me.chkFiles.Left = Me.Frame1.Left
  Me.Drive1.Left = Me.Frame1.Left
  Me.TreeView1.Height = Me.lblSubFolderCount.Top - Me.TreeView1.Top - 60
  Me.TreeView1.Width = Me.cmdTest.Left - Me.TreeView1.Left * 3
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Release FSO resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  If Not RootNode Is Nothing Then             'if node has been instantiated...
    Set RootNode = Nothing                    'release resources
  End If
  Set FSO = Nothing                           'reslease FSO resources
End Sub

'*******************************************************************************
' Subroutine Name   : cmdTest_Click
' Purpose           : Test the dynNode classes
'*******************************************************************************
Private Sub cmdTest_Click()
  Dim Drv As String                         'drive information
  Dim Nd As Node                            'treeview node
  Dim Ttl As Long, Mkd As Long
  Me.cmdTest.Enabled = False                'disable command button form now
  ClearTreeView Me.TreeView1                'clear the treeview of cany contents
  Dim Nds() As dynNode
  Dim cNode As dynNode
  Dim Tm As Date
  Dim S As String
  Dim Index As Long
'
' step 1
'
  Drv = UCase$(Me.Drive1.Drive)             'get drive to drill through
  Index = InStr(1, Drv, "[")
  If Index <> 0 Then Drv = Trim$(Left$(Drv, Index - 1))
  MsgBox "About to build dynNode tree by scanning ALL folders on drive " & Drv & vbCrLf & vbCrLf & _
         "Note: Drilling a drive for ALL folders will naturally take some" & vbCrLf & _
         "time on a large drive with, say, 10,000+ folders...", , "Step 1"
  
  Me.Enabled = False                        'disable everything
  Me.chkFiles.Enabled = False               'show visible disability
  Me.chkExpand.Enabled = False
  Me.Option1(0).Enabled = False
  Me.Option1(1).Enabled = False
  Me.Drive1.Enabled = False
  
  Screen.MousePointer = vbHourglass         'show that we are busy
  DoEvents                                  'let screen catch up
'
' instantiate our Root-Level Node. Take advantage of the user-defined marker
' to mark nodes we will designate as folders (unmarked nodes will be considered
' as files, or leaves)
'
  Set RootNode = CreateNewRoot(Drv, Drv)    'create new root-level node
  RootNode.Marker = True                    'mark it (we will mark all folders)
  RootNode.ErrorMsgBox = True               'display node errors in a message box
  RootNode.KeyChecks = False                'turn off Unique Key verification because
                                            'we know all keys will be unique, and
                                            'this will also cut building a large list
                                            'several times faster.
  
  Tm = Now()                                'we are going to time this process
  EnumerateDriveNodes RootNode              'now enumerate drive folders into the tree
  Screen.MousePointer = vbDefault           'done enumerating, so we are no longer busy
  Me.ZOrder 0                               'sometime we go behind the IDE...
  MsgBox "Elapsed time for Enumerating Drive: " & Format(Now() - Tm, "HH:MM:SS")
'
' step 2
'
  Ttl = RootNode.NodeCount                  'get total number of nodes
  Mkd = RootNode.MarkerCount                'get number of marked (all folder nodes)
  MsgBox "dynNode tree built. Now Sort the " & CStr(Ttl) & " dynNodes" & vbCrLf & _
         "(" & CStr(Mkd) & " folders and " & CStr(Ttl - Mkd) & " files collected)", , "Step 2"
  
  Screen.MousePointer = vbHourglass         'busy again
  DoEvents
  
  If Me.Option1(1).value Then
    RootNode.SortChildren True ' True       'sort children descending
  Else
    RootNode.SortChildren True, False       'sort children ascending
  End If
  Screen.MousePointer = vbDefault           'no longer busy
'
' step 3
'
  MsgBox "dynNode tree sorted. Now Build TreeView from " & CStr(Ttl) & " dynNodes", , "Step 3"
  Screen.MousePointer = vbHourglass         'busy again
  DoEvents
  Set Nd = Me.TreeView1.Nodes.Add(, tvwFirst, "K1", Drv, 1, 1) 'create root treeview node
  Nd.Expanded = True                        'ensure this one is expanded
  EnumerateTV RootNode, Nd                  'ensurnate out nodes into the treeview
  Nd.EnsureVisible                          'then make sure treeview root visible
  Nd.Selected = True
  ShowSubCount Nd.Index
  Screen.MousePointer = vbDefault           'no longer busy
'
' step 4
'
  MsgBox "TreeView Built. We will now get an array of all nodes" & vbCrLf & _
         "and display the ID and Key in the IDE Immediate Window", , "Step 4"
  Screen.MousePointer = vbHourglass         'busy again
  DoEvents
  Nds = RootNode.NodeList()
  Mkd = UBound(Nds)
  For Ttl = 0 To Mkd
    Debug.Print Format(Nds(Ttl).ID, "####0") & ": " & Nds(Ttl).Key
  Next Ttl
  Erase Nds                                 'erase list
  Screen.MousePointer = vbDefault           'no longer busy
'
' Last Step
'
  MsgBox "All operations completed. You can play with the TreeView if you want to.", , _
         "Last Step"
  Me.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : EnumerateDriveNodes
' Purpose           : Enumerate folders of drive into dynNode tree (recursive)
'*******************************************************************************
Private Sub EnumerateDriveNodes(Parent As dynNode)
  Dim Fld As Folder
  Dim Fil As File
  Dim cNd As dynNode
  Dim Str As String, Ppath As String
  Dim pIndex As Long
  
  Ppath = Parent.Key & "\"                        'get parent path
'
' build folders contained in Parent
'
  On Error Resume Next                            'in case of special protected folders
  AccumD = AccumD + 1
  For Each Fld In FSO.GetFolder(Ppath).SubFolders 'get all folders
    Set cNd = Parent.Nodes.Add(, dynNodeChild, Fld.Path, Fld.Name)   'add a node
    cNd.Marker = True                             'tag it as user-marked
    EnumerateDriveNodes cNd                       'enumerate each folder
  Next Fld
'
' build files contained in Parent
'
' Don't build file list as the TreeView control is limited to 32,000 entries, and will
' hange the system if more than that are added. My system, for example, presently has
' almost 154,000 files and folders on it. Though the dynNode list can easily accomodate
' this, we are stuck with the limitations of the least common denominator.
  If Me.chkFiles.value Then
    For Each Fil In FSO.GetFolder(Ppath).Files
     AccumF = AccumF + 1
      Call Parent.Nodes.Add(, dynNodeChild, Fil.Path, Fil.Name)  'add a node
    Next Fil
    Me.Caption = "Test dynNode Tree - " & CStr(AccumD) & " Folders, " & CStr(AccumF) & " Files"
  Else
    Me.Caption = "Test dynNode Tree - " & CStr(AccumD) & " Folders"
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : EnumerateTV
' Purpose           : Enumerate dynNode tree into the Treeview (recursive)
'*******************************************************************************
Private Sub EnumerateTV(Parent As dynNode, tvNode As Node)
  Dim tvN As Node
  Dim cNd As dynNode
  Dim Idx As Long, pIndex As Long, Count As Long
  
  pIndex = tvNode.Index                   'parent index
  With Parent.Nodes
    Count = .Count                        'speed access to var, not property
    For Idx = 1 To Count
      Set cNd = Parent.Nodes(Idx)         'grab dynNode
      If cNd.Marker Then
        Set tvN = Me.TreeView1.Nodes.Add(pIndex, tvwChild, "K" & CStr(cNd.ID), cNd.Text, 4, 4)
        If pIndex = 1 Then                'if parent is root, then update folder to screen
          tvN.EnsureVisible               'scroll as need to show it
          DoEvents
        End If
        If cNd.Children Then              'if node has children
          If Me.chkExpand.value Then      'expand all nodes?
            tvN.Expanded = True           'yes, so expand it
            tvN.Image = 5                 'set open folder image for display
            tvN.SelectedImage = 5
          End If
          Call EnumerateTV(cNd, tvN)      'enumerate child folders
        End If
      Else  'file
        Call Me.TreeView1.Nodes.Add(pIndex, tvwChild, "K" & cNd.ID, cNd.Text, 6, 6)
      End If
    Next Idx
  End With
End Sub

Private Sub RootNode_dynNodeError(ErrorCode As dynErrorCodes, NodeID As Long)
  Debug.Print "error code: " & CStr(ErrorCode) & ", Node ID = " & CStr(NodeID)
End Sub

'*******************************************************************************
' Subroutine Name   : TreeView1_Click
' Purpose           : Update the number of subfolder under the selected tv node
'*******************************************************************************
Private Sub TreeView1_Click()
  ShowSubCount Me.TreeView1.SelectedItem.Index
End Sub

'*******************************************************************************
' Subroutine Name   : TreeView1_Collapse
' Purpose           : Change ifolder icons for Closed
'*******************************************************************************
Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
  With Node
    If .Image = 5 Then
      .Image = 4
      .SelectedImage = 4
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : TreeView1_Expand
' Purpose           : Change ifolder icons for Opened
'*******************************************************************************
Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
  With Node
    If .Image = 4 Then
      .Image = 5
      .SelectedImage = 5
      .Child.LastSibling.EnsureVisible
      .EnsureVisible
    End If
  End With
  
End Sub

'*******************************************************************************
' Subroutine Name   : ShowSubCount
' Purpose           : Update the number of subfolder under the selected tv node
'*******************************************************************************
Private Sub ShowSubCount(NodeIndex As Long)
  Dim S As String
  Dim Dn As dynNode
  
  S = Me.TreeView1.Nodes(NodeIndex).Key
  S = Mid$(S, 2)
  Set Dn = RootNode.FindID(CLng(S))
  S = "Sub-Folders of this Folder: " & CStr(Dn.NodeCount - 1)
  If Dn.Locked Then S = S & ". LOCKED"
  Me.lblSubFolderCount.Caption = S
  
End Sub

'*******************************************************************************
' Function Name     : FindTag
' Purpose           : Sample function to drill down trough a tree, starting at the
'                   : specified node, and match the text of a tag
'
' Inputs            : Node: A dynNode object to being the search at
'                   : Tag:  The text to search for
'
' Outputs           : Returns Nothing if the tag is not found, otherwise it returns a
'                   : reference to the Node that contains the matching tag
'                   :
'
' Assumes           : The Node is a valid dynNode object
'*******************************************************************************
Function FindTag(Node As dynNode, Tag As Variant) As dynNode
  Dim Index As Long
'
' see if we have a local match
'
  If Node.Tag = Tag Then                            'match?
    Set FindTag = Node                              'yes, so return teh found node
    Exit Function
  End If
'
' else searach through the node list of children, and recurse through them
' If there are no child nodes, then the For-Next loop will not be entered
'
  With Node.Nodes
    For Index = 1 To .Count
      Set FindTag = FindTag(.Item(Index), Tag)      'do recursion (drill down)
      If Not FindTag Is Nothing Then Exit Function  'found a valid node match
    Next Index
  End With
End Function

