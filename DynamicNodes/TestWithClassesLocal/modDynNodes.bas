Attribute VB_Name = "modDynNodes"
Option Explicit

'Uncomment blocks of code as you require

'Private m_MyRootNode As dynNode 'A sample Root-Level Node object that is private
                                'to this file. This variable is used by the
                                'DefineMyRootNode(), MyRootNode(), and
                                'ReleaseMyRootNode() examples

''*******************************************************************************
'' Function Name     : CreateNewRoot
'' Purpose           : Return a new Root-Level Node to the invoker
''                   :
'' Inputs            : Key: An optional text indentifier for the new node
''                   : Text: a name for the node (no unique requirement)
''                   :
'' Outputs           : A new Root-Level Node
''                   :
'' Assumes           : DynamicNodes.DLL is referenced in the project
''*******************************************************************************
Public Function CreateNewRoot(Optional Key As String = vbNullString, _
                              Optional Text As String = vbNullString) As dynNode
  Dim cNd As New dynNode  'declare and instantiate a new Root-Level Node. STEP 1 of 2
  cNd.Init Key, Text      'initialize it (do this before anything else).  STEP 2 of 2
  Set CreateNewRoot = cNd 'return reference to created object (cNd is ALSO only a reference)
End Function
'
''*******************************************************************************
'' Function Name     : FindTag
'' Purpose           : Sample function to drill down trough a tree, starting at the
''                   : specified node, and match the contents of a tag
''
'' Inputs            : Node: A dynNode object to begin the search at
''                   : Tag:  The string data to search for
''
'' Outputs           : Returns Nothing if the tag is not found, otherwise it returns a
''                   : reference to the Node that contains the matching tag
''                   :
''
'' Assumes           : The Node is a valid dynNode object
''*******************************************************************************
'Function FindTag(Node As dynNode, Tag As String) As dynNode
''
'' See if we have a local match
''
'  If Node.Tag = Tag Then       'match found?
'    Set FindTag = Node         'yes, so return the found node
'    Exit Function
'  End If
''
'' Else searach through the node list of children, and recurse through them.
'' If there are no child nodes, then the For-Next loop will not be processed
''
'  With Node.Nodes
'    Dim Index As Long
'    For Index = 1 To .Count
'      Set FindTag = FindTag(.Item(Index), Tag)      'do recursion (drill down)
'      If Not FindTag Is Nothing Then Exit Function  'found a valid node match
'    Next Index
'  End With
'End Function
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'Private FSO As FileSystemObject    'standard system I/O object
'                                   '(Windows Scripting Host Object Model reference)
'Public RootNode As dynNode         'define new root node
''*******************************************************************************
'' Subroutine Name   : DrillDrive
'' Purpose           : Enumerate the folder/files of a drive and set a reference
''                   : to the new node list in the public RootNode variable
''*******************************************************************************
'Private Function DrillDrive(DriveLetter As String) As dynNode
'  Dim Drv As String
'  
'  Set FSO = New FileSystemObject            'Create system I/O object
'  Drv = DriveLetter & ":\"                  'build a full file path
'  Set RootNode = NewNodeList(Drv, Drv)      'create new root node (function defined in moddynNode.bas)
'  RootNode.Marker = True                    'mark it (we will mark all folders)
'  'NOTE: this does not 'make' it a folder, but we are simply using the user-definable
'  '      Marker property as a 'personal' reference.  This can be used for ANYTHING.
'  EnumerateNodes RootNode                   'enumerate drive folders into it
'  Set FSO = Nothing                         'release File I/O resources
'End Sub
'
''*******************************************************************************
'' Subroutine Name   : EnumerateNodes
'' Purpose           : Enumerate folders of drive into dynNode tree (recursive)
''
'' Note that this routine calls itself repeatedly.  This is a technique called
'' recursion, which is one of the coolest methods around for drilling through any
'' branching-type system.  ...Unless you miss-code it and do not debug-step through
'' it to ensure that it works, then the Vulcan Nerve Pinch (Ctrl-Alt-Delete with
'' one hand) may be required.
''*******************************************************************************
'Private Sub EnumerateNodes(Parent As dynNode)
'  Dim Fld As Folder     'Folder object from FSO classes
'  Dim Fil As File       'File object from FSO classes
'  Dim cNd As dynNode    'our local dynamic node
'  Dim Ppath As String   'parent path (speed referencing by not using properties)
'
'  Ppath = Parent.Key    'get parent path
''
'' Build folders contained in Parent.  Use On Error Resume Next in case we
'' encounter special system-level protected folder such as those on NT/XP
'' platforms that will let you see them, but doe-nah thoucha it!
''
'  On Error Resume Next                            'in case of special protected folders
'  For Each Fld In FSO.GetFolder(Ppath).SubFolders 'get all folders
'    Set cNd = Parent.Nodes.Add(, dynNodeChild, Fld.Path & "\", Fld.Name) 'add a node
'    cNd.Marker = True                             'tag it as user-marked (folder)
'    EnumerateNodes cNd                            'enumerate each folder
'  Next Fld
''
'' Build files contained in Parent (doing this into a TreeView node list would hang
'' the system on my computer, as a TreeView is limited to 32,000 entries.  dynNodes
'' allow up to 2,147,483,647 nodes (but who would be insane enough to try an make
'' a tree list that long -- at least until personal computers are replaced by
'' personal super-computers).
''
''SPECIAL NOTE!!!!:  Disable these following three lines (the For-Next loop) if you
''want to keep the wait down to a couple of minutes on a large system...
''
'  For Each Fil In FSO.GetFolder(Ppath).Files
'    Call Parent.Nodes.Add(, dynNodeChild, Fil.Path, Fil.Name)  'add a 'file' node
'  Next Fil
'End Sub
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
''*******************************************************************************
'' Function Name     : DefineMyRootNode
'' Purpose           : Define the Root Node object. This is a one-shot deal that
''                   : will only create the node if it does not exist
''*******************************************************************************
'Public Function DefineMyRootNode(Optional Key As String, Optional Text As String) As dynNode
'  If m_MyRootNode Is Nothing Then     'if the object does not exist...
'    Set m_MyRootNode = New dynNode    'instantiate it...
'    m_MyRootNode.Init Key, Text       'and initialize...
'  End If
'  Set DefineMyRootNode = m_MyRootNode 'return a reference to the Node to the caller, regardless
'End Function
'
''*******************************************************************************
'' Get Name          : MyRootNode
'' Purpose           : Get the MyRootNode object
''*******************************************************************************
'Public Property Get MyRootNode() As dynNode
'  Set MyRootNode = m_MyRootNode
'End Property
'
''*******************************************************************************
'' Subroutine Name   : ReleaseMyRootNode
'' Purpose           : Release the Root Node Object
''*******************************************************************************
'Public Sub ReleaseMyRootNode()
'  If Not m_MyRootNode Is Nothing Then   'if the object is instantiated...
'    Set m_MyRootNode = Nothing          'then release its resources
'  End If
'End Sub

