VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dynNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' dynNode - Node  Class module
' Copyright © 2004 by David Goben. All rights reserved.
'
' This is the main object class DynamicNodes. When each node is instantiated,
' it automatically creates a dyNodes Collection object to act as a container
' for any and all sub-nodes that will be build as subordinate to the root node.
'
' An accompanying module file, modDynNodes.bas, contains intitialization code for
' creating a new Root-Level Node. The code of this file is defined thus:
'
'- - - - - - - - - - - - - - - - - - CUT - - - - - - - - - - - - - - - - - - - -
''*******************************************************************************
'' Function Name     : CreateNewRoot
'' Purpose           : Create a new Root-Level Node
''*******************************************************************************
'Public Function CreateNewRoot(Optional Key As String = vbNullString, _
'                              Optional Text As String = vbNullString) As dynNode
'  Dim cNd As New dynNode  'declare and instantiate a new Root-Level Node. STEP 1 of 2
'  cNd.Init Key, Text      'initialize it (do this before anything else).  STEP 2 of 2
'  Set CreateNewRoot = cNd 'return reference to created object (cNd is ALSO only a reference)
'End Function
'- - - - - - - - - - - - - - - - - - CUT - - - - - - - - - - - - - - - - - - - -
'
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'NOTE: Set Instancing in the dynNode class properties to MultiUse. If you are testing
'      these class files as inclusions in a non-DLL test project, VB will auto-convert
'      their instancing to Private. Hence, when you reload them into the ActiveX DLL
'      project, you will have to manually convert them back.
'_______________________________________________________________________________
' Property:         I/O:  Returns:  'Description:
'   ID              G     Long      'Get unique ID number for the node
'   Key             G/L   String    'Get/Let Key text
'   Text            G/L   String    'Get/Let Text
'   Tag             G/L   Variant   'Get/Let Tag
'
'   UserObject      G/S   Object    'Get/Set a user-defined object
'   UserVar         G/S/L   Variant 'Get/Set/Let a user defined variable or structure
'
'   Sorted          G/L   Boolean   'Get/Let sorting on node's immediate children
'   SortDescending  G/L   Boolean   'False (Default) = Ascending. True = Descending
'   SortChildren    G     Boolean   'Sort all child notes and optionally branches w/o setting
'---
'   Locked          G/L   Boolean   'Get/Let flag to prevent deletion of nodes
'   LockedCount     G     Long      'Get number of nodes locked from this branch
'   LockedList      G     dynNode() 'Get list of dynNode() of lcoked nodes from this branch
'   UnlockAll       ---   ---       'Set all nodes from this branch to locked = False
'---
'   Marker          G/L   Boolean   'Get/Let user-usable boolean marker
'   MarkerCount     G     Long      'Get number of nodes marked from this branch
'   MarkerList      G     dynNode() 'Get list of dynNode() of marked nodes from this branch
'   ResetAllMarkers ---   ---       'Set all nodes to Marked = False
'---
'   Nodes           G     dynNode() 'Get list of child nodes
'   NodeCount       G     Long      'Get a count of nodes from the current node, inclusive
'   NodeList        G     dynNode() 'Get a list of all nodes from the specified node
'   Parent          G     dynNodes  'Get parent node. (This set by Init())
'   Children        G     Long      'Get number of child nodes
'   Child           G     dynNode   'Get First child of this node
'   FirstSibling    G/FL  dynNode   'Get first sibling of this node
'   LastSibling     G/FL  dynNode   'Get last sibling of this node
'   NextSibling     G/FL  dynNode   'Get next consecutive sibling of this node
'   PreviousSibling G/FL  dynNode   'Get previous consecutive sibling of this node
'   IsRoot          G     Boolean   'Get True if this node is a Root Node
'   Root            G     dynNode   'Get the Root Node of this tree
'   IsAncestor      G     Long      'Get 1 if the node is an ancestor of another node
'   IsDescendant    G     Long      'Get 1 if the node is a descendant of another node
'   Move            G     Boolean   'Get True if the node is set to a new parent
'---
'   FindID          G     Long      'Get index of specified ID. Searching from the root
'   FindIDLocal     G     Long      'Get index of specified ID, searching from this node downward
'   FindKey         G     Long      'Get Index of Matching Key (exact or partial)
'   FindKeyLocal    G     Long      'Get index of specified Key, searching from this node downward
'   KeyChecks       G/L   Boolean   'Get/Let Key Duplication checks
'   KeyExists       G     Boolean   'Return True if the specified Key exists in the tree
'--
'   Version         G     String    'DLL Version
'   ErrorCode       G     dynErrorCodes 'Get the last error code
'   ErrorNode       G     dynNode   'Get the last error Node
'   ErrorMsgBox     G/L   Boolean   'Get/Let flag to display a messagebox when errors encountered
'   ErrorDescription G    String    'Get a descriotion of the supplied node error
'   ClearLastError  ---   ---       'Clear the last error encountered
'   Init()          ---   ---       'Special initialization routine ONLY for root nodes
'________________________________________________________________________
'NOTE: G=GET, S=SET, L=LET, FG=Friend GET, FS=Friend SET, FL=Friend LET

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Common Enumerations used for dynNode.Nodes.Add
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
' Relationship Enumeration
'
Public Enum dynNodeRelationship
  dynNodeFirst = 0                        'in front of all siblings
  dynNodeLast = 1                         'after last sibling index
  dynNodeNext = 2                         'next after specified sibling index (default)
  dynNodePrevious = 3                     'previous to specified sibling index
  dynNodeChild = 4                        'child of the specified node index
End Enum
'
' Error Enumeration
'
Public Enum dynErrorCodes
  dynNodeErrSuccess = 0                     'no error
  dynNodeErrNotInitialized = 1              'node not yet initialized (Init())
  dynNodeErrAlreadyInitialized = 2          'node already initialized
  dynNodeErrNodeIsNothing = 3               'A specified node does not exist
  dynNodeErrParentAreSame = 4               'the node and the specified node share parent
  dynNodeErrNodesAreTheSame = 5             'This node is the same as the specified node
  dynNodeErrAncestorDescendantConflict = 6  'there was an ancestor/descendant conflict
  dynNodeErrIndexOutOfRange = 7             'specified index is out of range
  dynNodeErrRelationshipRequiresIndex = 8   'Specified relationship requires an index
  dynNodeErrKeyAlreadyExists = 9            'Key already exists
  dynNodeErrUserKeyIsNumeric = 10           'a user-supplied key cannot be numeric
  dynNodeErrKeyInvalid = 11                 'a user-supplied key is not String or numeric
  dynNodeErrKeyParameterBlank = 12          'a user-supplied key cannot be numeric
  dynNodeErrIdOutOfRange = 13               'the specified ID number is out of range
  dynNodeErrRootCanHaveNoSiblings = 14      'A sibling function on the root was performed
  dynNodeErrInvalidIndex = 15               'The specified index is non-integer, non-long
End Enum

'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
' Varaible local to this class
'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
Private m_g_Index As Long             'this tree's index into the global information
Private m_Index As Long
Private m_MyNodes As dynNodes         'colelction of children
Private m_Initialized As Boolean      'true when this item initialized
Private m_MyID As Long                'unique ID number for each node
Private m_MyKey As String             'key text
Private m_MyKeyLC As String           'lowercase version of Key
Private m_MyText As String            '"display" text
Private m_MyTag As Variant             'user-defined info
Private m_MySorted As Boolean         'TRUE if list to be sorted
Private m_MySortDescending As Boolean  'TRUE if list to be sorted in descending order
Private m_MyMarker As Boolean         'User-defined tagging flag
Private m_MyLocked As Boolean         'Node locked from Deletion when set
Private m_MyDeleting As Boolean       'Set to True when a node is being deleted
Private m_MyClassTerminating As Boolean 'True when the class is terminating
Private m_MyParent As dynNode         'parent of this node
Private m_MyNextSibling As Long       'next sibling of this node
Private m_MyPreviousSibling As Long   'previous sibling of this node
Private m_MyUserObj As Object         'generic user object pointer
Private m_MyUserVar As Variant        'generic user variant storage
Private m_MyKeyChecks As Boolean      'If key checks should be performed in ADDs
'------------------------------------------------------------------------------
' Error EVENT Support
'------------------------------------------------------------------------------
Public Event dynNodeError(ErrorCode As dynErrorCodes, NodeID As Long)

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Subroutine Name   : Class_Initialize
' Purpose           : Initialize dynNode Class
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Sub Class_Initialize()
  Set m_MyNodes = New dynNodes        'create dependent dynNodes
  m_MyNodes.Init Me                   'initialize it
  m_Initialized = False               'but not us until Init() is called
  m_MyKeyChecks = True                'init key checking On
End Sub

'*******************************************************************************
' Subroutine Name   : Init
' Purpose           : Initialize new dynNodes class ROOT with no node reference
'*******************************************************************************
Public Sub Init(Optional Key As String, Optional Text As String)
  If Not m_Initialized Then           'if not initialized, then allow setting parent
    m_Initialized = True              'set initialization flag
'
'  set up Text
'
    m_MyText = Trim$(Text)            'assign text so we can see where we are
'
' first check to see if we are a root-level node. If we are, then we will not
' have a parent. Otherwise, we are being invoked from Friend FInit(), which
' had first set a parent for this node
'
    If m_MyParent Is Nothing Then     'if this is a root node object...
      Call BuildGlobals               'build global information
'
' if not root level, get the gkobal index set for the Root Level Node
'
    Else
      m_g_Index = m_MyParent.GetgIndex 'get the m_g_index value from the root and save locally
      m_MyID = m_MyNodes.GetNewID()    'get unique ID for this node
    End If
'
'  set up Key data
'
    Me.Key = Key                      'assign key
    If Len(m_MyKey) = 0 Then          'if key was not set, then user default ID value
      Call IDKey                      'use ID value if null
    End If
'
' if the user supplied Key generated an error, go ahead and clear it
'
    g_LastError(m_g_Index) = dynNodeErrSuccess  'tag success
'
' if the node is already initialized, issue an error
'
  Else
    SetErrEvent dynNodeErrAlreadyInitialized
  End If
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Subroutine Name   : Class_Terminate
' Purpose           : Terminate dynNode Class
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Sub Class_Terminate()
  m_MyClassTerminating = True     'indicate this class is history
  m_MyNodes.Clear                 'clear any and all child nodes from this branch
  Set m_MyUserObj = Nothing       'make sure any stored object is toasted
  If VarType(m_MyUserVar) = vbObject Then
    Set m_MyUserVar = Nothing
  End If
  If m_MyID = 1 Then              'if root node...
    g_ID(m_g_Index) = 0           'free global slot if this is the parent node
  End If
End Sub

'*******************************************************************************
' Property Name     : ID
' Purpose           : Return the node's Unique ID number
'*******************************************************************************
Public Property Get ID() As Long
  ID = m_MyID
End Property

'*******************************************************************************
' Get Name          : Key
' Purpose           : Get Key value
'*******************************************************************************
Public Property Get Key() As String
  Key = m_MyKey
End Property

'*******************************************************************************
' Let Name          : Key
' Purpose           : Assign key value
'*******************************************************************************
Public Property Let Key(KeyText As String)
  Dim TKey As String, Key As String
  
  TKey = Trim$(KeyText)                   'trim text
  If Len(TKey) = 0 Then                   'if blank...
    Call IDKey                            'use default (ID value) if null
  Else
    Key = LCase$(TKey)                    'get trimmed, lowercase version
    If IsNumeric(Left$(Key, 1)) Then      'avoid user-entered numeric keys
      Call SetErrEvent(dynNodeErrUserKeyIsNumeric)
    Else
      If StrComp(m_MyKeyLC, Key, vbBinaryCompare) <> 0 Then 'check only if they differ
        If KeyChecks Then
          If KeyExists(Key) Then          'returns is True if key found
            SetErrEvent dynNodeErrKeyAlreadyExists
          End If
        End If
        m_MyKey = TKey                    'if key is unique, then add
        m_MyKeyLC = Key                   'keep lowercase version
      End If
    End If
  End If
End Property

'*******************************************************************************
' Get Name          : Text
' Purpose           : Get Text
'*******************************************************************************
Public Property Get Text() As String
Attribute Text.VB_Description = "CNode Default Property"
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = 0
  Text = m_MyText
End Property

'*******************************************************************************
' Let Name          : Text (DEFAULT PROPERTY)
' Purpose           : Assign text
'
' This is the default property for this class. With this Class active in the IDE,
' Select TOOLS/Procedure Attributes.../ADVANCED to see how this property was set:
' Set Text as NAME, and Procedure ID to (Default). Property Category: Text
'*******************************************************************************
Public Property Let Text(TextText As String)
  m_MyText = TextText
End Property

'*******************************************************************************
' Get Name          : Tag
' Purpose           : Get user tag
'*******************************************************************************
Public Property Get Tag() As Variant
  Tag = m_MyTag
End Property

'*******************************************************************************
' Let Name          : Tag
' Purpose           : Assign user tag
'*******************************************************************************
Public Property Let Tag(TagData As Variant)
  m_MyTag = TagData
End Property

'*******************************************************************************
' Get Name          : UserObject
' Purpose           : Get User-defined stored object
'*******************************************************************************
Public Property Get UserObject() As Object
  Set UserObject = m_MyUserObj
End Property

'*******************************************************************************
' Set Name          : UserObject
' Purpose           : Set User-defined stored object
'*******************************************************************************
Public Property Set UserObject(UsrObject As Object)
  Set m_MyUserObj = UsrObject
End Property

'*******************************************************************************
' Get Name          : UserVar
' Purpose           : Get User-defined stored variant
'                   : This can be a store variable or structure
'*******************************************************************************
Public Property Get UserVar() As Variant
  If VarType(m_MyUserVar) = vbObject Then
    Set UserVar = m_MyUserVar
  Else
    UserVar = m_MyUserVar
  End If
End Property

'*******************************************************************************
' Set Name          : UserVar
' Purpose           : Set User-defined stored variant
'                   : This can be a store variable or structure
'*******************************************************************************
Public Property Set UserVar(UserObject As Object)
  Set m_MyUserVar = UserObject
End Property

'*******************************************************************************
' Let Name          : UserVar
' Purpose           : Set User-defined stored variant
'                   : This can be a store variable or structure
'*******************************************************************************
Public Property Let UserVar(UserVariant As Variant)
  m_MyUserVar = UserVariant
End Property

'*******************************************************************************
' Get Name          : Marker
' Purpose           : Get user marker. This is a user-defined boolean that the
'                   : user can take advantage of to tag their items
'*******************************************************************************
Public Property Get Marker() As Boolean
  Marker = m_MyMarker
End Property

'*******************************************************************************
' Let Name          : Marker
' Purpose           : assign user marker
'*******************************************************************************
Public Property Let Marker(ByVal value As Boolean)
  m_MyMarker = value
End Property

'*******************************************************************************
' Function Name     : MarkerCount
' Purpose           : Get a count of  all marked items from this node down
'*******************************************************************************
Public Function MarkerCount() As Long
  If m_Initialized Then
    Dim Index As Long, Cnt As Long
    
    Cnt = 0                                       'init nothing marked
    If m_MyMarker Then Cnt = 1                    'if this node is marked
    With m_MyNodes
      For Index = 1 To .Count                     'process all child nodes
        Cnt = Cnt + .Item(Index).MarkerCount()    'accumulate counts from children
      Next Index
    End With
    MarkerCount = Cnt                             'return value
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : MarkerList
' Purpose           : Get a list of all nodes under this one that are marked
'*******************************************************************************
Public Function MarkerList() As dynNode()
  If m_Initialized Then
    MarkerList = GetAllMarked()    'return a list of ALL descended node's markers
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Subroutine Name   : ResetAllMarkers
' Purpose           : Reset all child markers from this node downward
'*******************************************************************************
Public Sub ResetAllMarkers()
  If m_Initialized Then
    Dim Index As Long
    
    m_MyMarker = False                'reset my marker
    With m_MyNodes
      For Index = 1 To .Count         'process all child nodes
        .Item(Index).ResetAllMarkers  'drill down to all sub-nodes and reset them
      Next Index
    End With
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Sub

'*******************************************************************************
' Get Name          : Sorted
' Purpose           : Return TRUE if the list is sorted
'*******************************************************************************
Public Property Get Sorted() As Boolean
  If m_Initialized Then
    Sorted = m_MySorted
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Property

'*******************************************************************************
' Let Name          : Sorted
' Purpose           : Set the sorting flag. Sort as needed
'*******************************************************************************
Public Property Let Sorted(ByVal SetSort As Boolean)
  If m_Initialized Then
    m_MySorted = SetSort        'set the actual flag
    m_MyNodes.Sorted = SetSort  'pass on to Nodes
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Property

'*******************************************************************************
' Get Name          : SortDescending
' Purpose           : Return TRUE if the list is set for Descending sort order
'*******************************************************************************
Public Property Get SortDescending() As Boolean
  If m_Initialized Then
    SortDescending = m_MySortDescending
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Property

'*******************************************************************************
' Let Name          : SortDescending
' Purpose           : Set the sorting flag. Sort as needed
'*******************************************************************************
Public Property Let SortDescending(ByVal SetSort As Boolean)
  If m_Initialized Then
    m_MySortDescending = SortDescending        'set the actual flag
    m_MyNodes.SortDescending = SortDescending  'pass on to nodes
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Property

'*******************************************************************************
' Subroutine Name   : SortChildren
' Purpose           : Shell/Metzner Sort the child nodes (VERY FAST) W/O setting
'                   : the Sorted Flag. Handy for one-shot sorts
'*******************************************************************************
Public Function SortChildren(Optional SortAllGenerations As Boolean = False, _
                             Optional ReverseOrder As Boolean = False) As Boolean
  If m_Initialized Then
    SortChildren = m_MyNodes.SortChildren(SortAllGenerations, ReverseOrder)
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Get Name          : Locked
' Purpose           : Get the state of the Locked flag
'*******************************************************************************
Public Property Get Locked() As Boolean
  Locked = m_MyLocked
End Property

'*******************************************************************************
' Let Name          : Locked
' Purpose           : Set the state of the Locked flag
'*******************************************************************************
Public Property Let Locked(Flag As Boolean)
  m_MyLocked = Flag
End Property

'*******************************************************************************
' Function Name     : LockedCount
' Purpose           : Get a count of  all locked items from this node down
'*******************************************************************************
Public Function LockedCount() As Long
  If m_Initialized Then
    Dim Index As Long, Cnt As Long
    
    Cnt = 0
    If m_MyLocked Then Cnt = 1                    'if this node is locked
    With m_MyNodes
      For Index = 1 To .Count                     'process all child nodes
        Cnt = Cnt + .Item(Index).LockedCount()    'accumulate counts
      Next Index
    End With
    LockedCount = Cnt                             'return value
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : LockedList
' Purpose           : Get a list of all nodes under this one that are Locked
'*******************************************************************************
Public Function LockedList() As dynNode()
  If m_Initialized Then
    LockedList = GetAllLocked()    'return a list of ALL descended node's markers
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Subroutine Name   : UnlockAll
' Purpose           : Reset all child Locks from this node downward
'*******************************************************************************
Public Sub UnlockAll()
  If m_Initialized Then
    Dim Index As Long
    
    m_MyLocked = False                'reset my lock
    With m_MyNodes
      For Index = 1 To .Count         'process all child nodes
        .Item(Index).UnlockAll        'drill down to all sub-nodes and reset them
      Next Index
    End With
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Sub

'*******************************************************************************
' Get Name          : Nodes
' Purpose           : Get a reference to the child node list
'*******************************************************************************
Public Property Get Nodes() As dynNodes
  Set Nodes = m_MyNodes
End Property

'*******************************************************************************
' Function Name     : NodeCount
' Purpose           : Get a count of nodes from the current branch
'*******************************************************************************
Public Function NodeCount() As Long
  If m_Initialized Then
    Dim Index As Long, Cnt As Long
    
    Cnt = 1                                       'count this one
    With m_MyNodes
      For Index = 1 To .Count                     'process all child nodes
        Cnt = Cnt + .Item(Index).NodeCount()      'accumulate counts
      Next Index
    End With
    NodeCount = Cnt                              'return value
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : NodeList
' Purpose           : Return an array of dynNode that contains all nodes in this branch.
'*******************************************************************************
Public Function NodeList() As dynNode()
  Dim ListNodes() As dynNode, Tmp() As dynNode
  Dim Index As Long, I As Long, LCount As Long, UB As Long
  
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
    Exit Function
  End If
  
  LCount = 1                                'init local size to 0
  ReDim ListNodes(0) As dynNode             'make room for Me
  Set ListNodes(0) = Me                     'ad me to the list
'
' now populate with children of this node
'
  With m_MyNodes
      For Index = 1 To .Count               'process all child nodes
      Tmp = .Item(Index).NodeList           'get a list (will have at leaset 1 item in it)
      UB = UBound(Tmp)                      'get upper bounds
      ReDim Preserve ListNodes(LCount + UB) 'make room in the local array
      For I = 0 To UB
        Set ListNodes(LCount + I) = Tmp(I)  'Append new data to local list
      Next I
      LCount = LCount + UB + 1              'bump to next available index
    Next Index                              'process all child nodes
  End With
  NodeList = ListNodes                      'return the list
End Function

'*******************************************************************************
' Property Name     : Parent
' Purpose           : Return this node's parent node
'*******************************************************************************
Public Property Get Parent() As dynNode
  Set Parent = m_MyParent
End Property

'*******************************************************************************
' Property Name     : Parent
' Purpose           : Set this node's parent node
'*******************************************************************************
Friend Property Set Parent(NewParent As dynNode)
  Set m_MyParent = NewParent
End Property

'*******************************************************************************
' Function Name     : Children
' Purpose           : Return the number of child nodes to this node
'*******************************************************************************
Public Function Children() As Long
  Children = m_MyNodes.Count()
End Function

'*******************************************************************************
' Function Name     : Child
' Purpose           : Return a reference to the first child node
'*******************************************************************************
Public Function Child() As dynNode
  If m_Initialized Then
    If CBool(m_MyNodes.Count) Then Set Child = m_MyNodes(1)
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : FirstSibling
' Purpose           : Return a reference to the node's first sibling
'*******************************************************************************
Public Function FirstSibling() As dynNode
  If m_Initialized Then
    If m_MyID <> 1 Then Set FirstSibling = m_MyParent.Child()
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : LastSibling
' Purpose           : Return a reference to the node's last sibling
'*******************************************************************************
Public Function LastSibling() As dynNode
  If m_Initialized Then
    If m_MyID <> 1 Then Set LastSibling = m_MyParent.gLastSibling()
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : NextSibling
' Purpose           : Return a reference to the node's next sibling
'*******************************************************************************
Public Function NextSibling() As dynNode
  If m_Initialized Then
    If m_MyID <> 1 Then Set NextSibling = gNextSibling()
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : PreviousSibling
' Purpose           : Return a reference to the node's previous sibling
'*******************************************************************************
Public Function PreviousSibling() As dynNode
  If m_Initialized Then
    If m_MyID <> 1 Then PreviousSibling = gPreviousSibling()
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : IsRoot
' Purpose           : Return True if this node is a root node
'*******************************************************************************
Public Function IsRoot() As Boolean
  If m_Initialized Then
    IsRoot = m_MyID = 1
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : Root
' Purpose           : Get the Root node of this node extremely quickly
'*******************************************************************************
Public Function Root() As dynNode
  
  If m_Initialized Then
    Set Root = g_Root(m_g_Index)          'root index
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : IsAncestor
' Purpose           : Find out if the node is an Ancestor of the specified node.
'                   : Return 1 if the node is an Ancestor, -1 if the node is not
'                   : an Ancestor, and 0 if there was an error
'*******************************************************************************
Public Function IsAncestor(OfNode As dynNode) As Long
  Dim Nd As dynNode
  
  If m_Initialized Then
    If OfNode Is Nothing Then                       'if bogus node being checked
      SetErrEvent dynNodeErrNodeIsNothing           'error (also return 0 error code)
    Else
      If OfNode.ID = m_MyID Then                    'if they are the SAME node
        Call SetErrEvent(dynNodeErrNodesAreTheSame) 'error
      Else
        Set Nd = OfNode                             'else begin search
        Do While Not Nd Is Nothing                  'while this node exists
          If Nd.ID = m_MyID Then                    'found match?
            IsAncestor = 1                          'yes, so we are it's ancestor
            Exit Do
          End If
          Set Nd = Nd.Parent                        'else skip back a generation
        Loop
        IsAncestor = -1                             'indicate we are not an ancestor
      End If
    End If
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : IsDescendant
' Purpose           : Find out if the node is a descendant of the specified node.
'                   : Return 1 if the node is a descendant, -1 if the node is not
'                   : a descendant, and 0 if there was an error
'*******************************************************************************
Public Function IsDescendant(OfNode As dynNode) As Long
  Dim ID As Long
  Dim Nd As dynNode
  
  If m_Initialized Then                             'if we are initialized
    If OfNode Is Nothing Then                       'if bogus node being checked
      SetErrEvent dynNodeErrNodeIsNothing           'error (also return 0 error code)
    Else
      If OfNode.ID = m_MyID Then                    'if they are the SAME node
        Call SetErrEvent(dynNodeErrNodesAreTheSame) 'error
      Else
        ID = OfNode.ID                              'get ID to scan for
        Set Nd = Me                                 'else begin search with self
        Do While Not Nd Is Nothing                  'while this node exists
          If Nd.ID = ID Then                        'found match?
            IsDescendant = 1                        'yes, so we are a descendant of the node
            Exit Do
          End If
          Set Nd = Nd.Parent                        'else skip back a generation
        Loop
        IsDescendant = -1                           'indicate we are not a descendant
      End If
    End If
  Else ' if not initialized, then do not let anything happen
    SetErrEvent dynNodeErrNotInitialized
  End If
End Function

'*******************************************************************************
' Function Name     : Move
' Purpose           : Move a node to a new parent location
'*******************************************************************************
Public Function Move(NewParent As dynNode) As Boolean
  Dim OldParent As dynNode
  
  If m_Initialized Then
    If m_MyParent Is NewParent Then               'parent already set?
        Move = True                               'assume success
        Exit Function
    End If
    If Not NewParent Is Nothing Then                            'new not root?
      If IsAncestor(NewParent) = 1 Then                         'are we its ancestor?
        Call SetErrEvent(dynNodeErrAncestorDescendantConflict)  'yes, so oops!
        Exit Function
      End If
    End If
    Set OldParent = m_MyParent                    'save our original parent
    Set m_MyParent = NewParent                    'set our new parent
    If m_MyParent Is Nothing Then                 'if we are to become a ROOT node
      Call BuildGlobals                           'build new global reference
    Else
      m_MyParent.ResetgIndex                      'reset global index, if changed
    End If
'
' reset both parent's child relationships, if they are not nothing
'
    If Not OldParent Is Nothing Then OldParent.Nodes.ResetRelationships
    If Not m_MyParent Is Nothing Then m_MyParent.Nodes.ResetRelationships
    Move = True                               'tag success
  Else
    Call SetErrEvent(dynNodeErrNotInitialized)
  End If
End Function

'*******************************************************************************
' Function Name     : FindID
' Purpose           : Find a specified ID, searching from the root
'                   : Return the dynNode. Return Nothing if not found
'*******************************************************************************
Public Function FindID(ByVal ID As Long) As dynNode
  Dim cNd As dynNode
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
'
' check ID range
'
    If ID < 1 Or ID > g_ID(m_g_Index) Then
      SetErrEvent dynNodeErrIdOutOfRange        'ID is out of range
    Else
'
' not matched in this node, so check children
'
      Set FindID = Root.LocalFindID(ID)
    End If
  End If
End Function

'*******************************************************************************
' Function Name     : FindIDLocal
' Purpose           : Find a specified ID from the current node downward.
'                   : Return the dynNode. Return Nothing if not found
'*******************************************************************************
Public Function FindIDLocal(ByVal ID As Long) As dynNode
  Dim cNd As dynNode
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
'
' check ID range
'
    If ID < 1 Or ID > g_ID(m_g_Index) Then
      SetErrEvent dynNodeErrIdOutOfRange        'ID is out of range
    Else
'
' not matched in this node, so check children
'
      Set FindIDLocal = Me.LocalFindID(ID)
    End If
  End If
End Function

'*******************************************************************************
' Function Name     : FindKey
' Purpose           : Find a specified key, searching from the root
'                   : Return the dynNode. Return Nothing if not found
'*******************************************************************************
Public Function FindKey(Key As String) As dynNode
  Dim lcKey As String
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
'
' check for valid key
'
    lcKey = LCase$(Trim$(Key))                            'strip spaces
    If Len(lcKey) = 0 Then
      SetErrEvent dynNodeErrKeyParameterBlank             'bad key
    Else
'
' not matched in this node, so check children
'
      Set FindKey = Root.LocalFindKey(lcKey)
    End If
  End If
End Function

'*******************************************************************************
' Function Name     : FindKeyLocal
' Purpose           : Find a specified key from the current key downward.
'                   : Return the dynNode. Return Nothing if not found
'*******************************************************************************
Public Function FindKeyLocal(Key As String) As dynNode
  Dim lcKey As String
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
'
' check for valid key
'
    lcKey = LCase$(Trim$(Key))                            'strip spaces
    If Len(lcKey) = 0 Then
      SetErrEvent dynNodeErrKeyParameterBlank             'bad key
    Else
'
' not matched in this node, so check children
'
      Set FindKeyLocal = Me.LocalFindKey(lcKey)
    End If
  End If
End Function

'*******************************************************************************
' Get Name          : KeyChecks
' Purpose           : Get the state of key states
'*******************************************************************************
Public Property Get KeyChecks() As Boolean
  If m_MyID = 1 Then
    KeyChecks = m_MyKeyChecks
  Else
    KeyChecks = m_MyParent.KeyChecks
  End If
End Property

'*******************************************************************************
' Let Name          : KeyChecks
' Purpose           : Set the state of key checking
'*******************************************************************************
Public Property Let KeyChecks(ByVal Enabled As Boolean)
  If m_MyID = 1 Then
    m_MyKeyChecks = Enabled
  Else
    Root.KeyChecks = KeyChecks
  End If
End Property

'*******************************************************************************
' Function Name     : KeyExists
' Purpose           : Find a specified key. Return False if not found
'*******************************************************************************
Public Function KeyExists(Key As String) As Boolean
  Dim lcKey As String
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
'
' check for valid key
'
    lcKey = LCase$(Trim$(Key))                            'strip spaces
    If Len(lcKey) = 0 Then
      SetErrEvent dynNodeErrKeyParameterBlank             'bad key
    Else
'
' no start recursion to check from the root-level node to the end for a match
'
      KeyExists = Root.LocalKeyExists(lcKey) 'called routines assunme LC key
    End If
  End If
End Function

'*******************************************************************************
' Function Name     : Version
' Purpose           : Return Application Version
'*******************************************************************************
Public Function Version() As String
  Version = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function

'*******************************************************************************
' Get Name          : ErrorCode
' Purpose           : Get the last error code
'*******************************************************************************
Public Property Get ErrorCode() As dynErrorCodes
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
    ErrorCode = g_LastError(m_g_Index)
  End If
End Property

'*******************************************************************************
' Get Name          : ErrorNode
' Purpose           : Return the node that the last error occurred within
'*******************************************************************************
Public Property Get ErrorNode() As dynNode
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
    If g_LastError(m_g_Index) <> dynNodeErrSuccess Then
      Set ErrorNode = g_LastErrorNode(m_g_Index)
    End If
  End If
End Property

'*******************************************************************************
' Get Name          : ErrorMsgBox
' Purpose           : Return True if errors are to be displayed in a message box
'*******************************************************************************
Public Property Get ErrorMsgBox() As Boolean
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
    ErrorMsgBox = g_MsgBoxErrors(m_g_Index)
  End If
End Property

'*******************************************************************************
' Let Name          : ErrorMsgBox
' Purpose           : Set True if errors are to be displayed in a message box
'*******************************************************************************
Public Property Let ErrorMsgBox(ByVal ShowErrorsInMsgBox As Boolean)
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
    g_MsgBoxErrors(m_g_Index) = ShowErrorsInMsgBox
  End If
End Property

'*******************************************************************************
' Function Name     : ErrorDescription
' Purpose           : Return a text description of the Node error code
'*******************************************************************************
Public Function ErrorDescription(ErrorNumber As dynErrorCodes) As String
  Dim S As String
  
  Select Case ErrorNumber
    Case dynNodeErrSuccess
      S = "No error"
    Case dynNodeErrNotInitialized
      S = "The Root-Level Node is not yet initialized by the Init() method"
    Case dynNodeErrAlreadyInitialized
      S = "The Root-Level Node is already initialized. The user is attempting to invoke the Init() method on it again"
    Case dynNodeErrNodeIsNothing
      S = "The, or a specified node does not exist"
    Case dynNodeErrParentAreSame
      S = "The node and the specified node share the same parent"
    Case dynNodeErrNodesAreTheSame
      S = "This node is the same as the specified node"
    Case dynNodeErrAncestorDescendantConflict
      S = "There was an ancestor/descendant conflict between two specified nodes"
    Case dynNodeErrIndexOutOfRange
      S = "The specified index is out of range in a Nodes Collection"
    Case dynNodeErrRelationshipRequiresIndex
      S = "The specified Relationship during an Add() method invocation requires an Index or Key be supplied"
    Case dynNodeErrKeyAlreadyExists
      S = "An attempt was made to either change a node' Key value, or add a new Node with a Key value that already exists in the Tree"
    Case dynNodeErrUserKeyIsNumeric
      S = "A user-supplied key cannot be numeric"
    Case dynNodeErrKeyInvalid
      S = "The user-supplied key is either not an existing Key or numeric index"
    Case dynNodeErrKeyParameterBlank
      S = "A Key or Index was not supplied"
    Case dynNodeErrIdOutOfRange
      S = "The specified ID number is out of range"
    Case dynNodeErrRootCanHaveNoSiblings
      S = "A sibling function was attempted on the Root-Level Node"
    Case dynNodeErrInvalidIndex
      S = "The specified Index value is not an integer and not a long integer"
    Case Else
      S = "UNKNOWN Error"
  End Select
  ErrorDescription = S
End Function

'*******************************************************************************
' Subroutine Name   : ClearLastError
' Purpose           : Reset error code
'*******************************************************************************
Public Sub ClearLastError()
'
' error handling
'
  If Not m_Initialized Then
    SetErrEvent dynNodeErrNotInitialized
  Else
    g_LastError(m_g_Index) = dynNodeErrSuccess
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : BuildGlobals
' Purpose           : Build a global entry for a Root-Level Node
'*******************************************************************************
Private Sub BuildGlobals()
'
' if global stuff exists, see if a "slot" open for use
'
  m_g_Index = g_gSize                           'get current size of global arrays
  If CBool(g_gSize) Then                        'if defined, then check for open slot
    For m_g_Index = 0 To g_gSize - 1
      If Not g_Tag(m_g_Index) Then Exit For     'found an open slot, so we will use it
    Next m_g_Index
  End If
'
' if no slot found, then expand the array sizes
'
  If m_g_Index = g_gSize Then
    ReDim Preserve g_Tag(m_g_Index) As Boolean  'make room
    ReDim Preserve g_Root(m_g_Index) As dynNode
    ReDim Preserve g_ID(m_g_Index) As Long
    ReDim Preserve g_LastError(m_g_Index) As dynErrorCodes
    ReDim Preserve g_LastErrorNode(m_g_Index) As dynNode
    ReDim Preserve g_MsgBoxErrors(m_g_Index) As Boolean
    g_gSize = g_gSize + 1                       'bump global size index
  End If
'
' stuff tag to indicate this stot is in use
'
  g_Tag(m_g_Index) = True
'
' initialize other gData
'
  Set g_Root(m_g_Index) = Me                  'stuff root node
  g_ID(m_g_Index) = 0                         'init ID to Next-1
  g_LastError(m_g_Index) = dynNodeErrSuccess  'tag success
  Set g_LastErrorNode(m_g_Index) = Nothing    'no node to reference
  g_MsgBoxErrors(m_g_Index) = False           'no messagebox stuff
  ReID 0                                      'reindex new tree (if MOVE/COPY)
End Sub

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'
'                    FRIEND ROUTINES FOR INTERNAL SUPPORT
'
' These methods and properties are visible to this project, but are not exposed
' outside of it in the user-interface.
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'*******************************************************************************
' Subroutine Name   : ReID
' Purpose           : Reindex nodes when branch becomes separate tree
'*******************************************************************************
Friend Sub ReID(ByVal gIndex As Long)
  m_g_Index = gIndex                    'set new global index
  g_ID(m_g_Index) = g_ID(m_g_Index) + 1 'bump ID index
  m_MyID = g_ID(m_g_Index)              'set new node ID
  m_MyNodes.ReID m_g_Index              'process children
End Sub

'*******************************************************************************
' Function Name     : gLastSibling
' Purpose           : Return a reference to the last child node
'*******************************************************************************
Friend Function gLastSibling() As dynNode
  If m_MyNodes.Count Then Set gLastSibling = m_MyNodes(m_MyNodes.Count)
End Function

'*******************************************************************************
' Function Name     : gNextSibling
' Purpose           : Return a reference to the next child node
'*******************************************************************************
Friend Function gNextSibling() As dynNode
  If m_MyNextSibling > 0 Then Set gNextSibling = m_MyNodes(m_MyNextSibling)
End Function

'*******************************************************************************
' Function Name     : sNextSibling
' Purpose           : Set a reference to the last child node
'*******************************************************************************
Friend Sub sNextSibling(ByVal Index As Long)
  m_MyNextSibling = Index
End Sub

'*******************************************************************************
' Function Name     : gPreviousSibling
' Purpose           : Return a reference to the previous child node
'*******************************************************************************
Friend Function gPreviousSibling() As dynNode
  If m_MyPreviousSibling > 0 Then Set gPreviousSibling = m_MyNodes(m_MyPreviousSibling)
End Function

'*******************************************************************************
' Function Name     : sPreviousSibling
' Purpose           : Set a reference to the previous child node
'*******************************************************************************
Friend Sub sPreviousSibling(ByVal Index As Long)
  m_MyPreviousSibling = Index
End Sub

'*******************************************************************************
' Function Name     : GetAllMarked
' Purpose           : Get all marked items from this node down
'*******************************************************************************
Friend Function GetAllMarked() As dynNode()
  Dim Index As Long, Cnt As Long, I As Long, UB As Long
  Dim Nds() As dynNode, sNds() As dynNode
  
  If m_MyMarker Then                              'is THIS node marked?
    ReDim Nds(0) As dynNode                       'yes, so save it to the list
    Set Nds(0) = Me                               'stuff a reference to self in it
    Cnt = 1                                       'count 1 gathered
  Else
    Cnt = 0                                       'else list still empty
  End If
'
' now scan through child nodes and gather their lists
'
  With m_MyNodes
      For Index = 1 To .Count                     'process all child nodes
      sNds = .Item(Index).GetAllMarked()          'recurse through each for all marked
      On Error Resume Next                        'error trap
      UB = UBound(sNds)                           'get upper bounds
      If Err.Number = 0 Then                      'we have a dimmed array
        ReDim Preserve Nds(Cnt + UB) As dynNode   'make local container bigger
        For I = 0 To UB
          Set Nds(Cnt + I) = sNds(I)              'append new list to local
        Next I
        Cnt = Cnt + UB + 1                        'bump new count
      End If
    Next Index                                    'do all child nodes
  End With
  GetAllMarked = Nds                              'return list
End Function

'*******************************************************************************
' Function Name     : GetAllLocked
' Purpose           : Get all locked items from this node down
'*******************************************************************************
Friend Function GetAllLocked() As dynNode()
  Dim Index As Long, Cnt As Long, I As Long, UB As Long
  Dim Nds() As dynNode, sNds() As dynNode
  
  If m_MyLocked Then                              'is THIS node locked?
    ReDim Nds(0) As dynNode                       'yes, so save it to the list
    Set Nds(0) = Me                               'stuff a reference to self in it
    Cnt = 1                                       'count 1 gathered
  Else
    Cnt = 0                                       'else list still empty
  End If
'
' now scan through child nodes and gather their lists
'
  With m_MyNodes
      For Index = 1 To .Count                     'process all child nodes
      sNds = .Item(Index).GetAllLocked()          'recurse through each for all locked
      On Error Resume Next                        'error trap
      UB = UBound(sNds)                           'get upper bounds
      If Err.Number = 0 Then                      'we have a dimmed array
        ReDim Preserve Nds(Cnt + UB) As dynNode   'make local container bigger
        For I = 0 To UB
          Set Nds(Cnt + I) = sNds(I)              'append new list to local
        Next I
        Cnt = Cnt + UB + 1                        'bump new count
      End If
    Next Index                                    'do all child nodes
  End With
  GetAllLocked = Nds                              'return list
End Function

'*******************************************************************************
' Subroutine Name   : SetErrEvent
' Purpose           : Set error flags
'*******************************************************************************
Friend Sub SetErrEvent(ByVal ErrorCode As dynErrorCodes)
  g_LastError(m_g_Index) = ErrorCode                'set error code
  If ErrorCode <> dynNodeErrSuccess Then            'if an actual error
    Set g_LastErrorNode(m_g_Index) = Me             'set error node
    RaiseEvent dynNodeError(ErrorCode, m_MyID)      'raise error event
'
' if we should display errors in a messagebox, then do so
'
    If g_MsgBoxErrors(m_g_Index) Then               'display msgbox?
      Dim Txt As String                             'yes
      Txt = "dynNode Runtime Error # " & CStr(ErrorCode) & _
      ".    Reference Node ID: " & CStr(m_MyID) & ".     Key: " & m_MyKey & vbCrLf & vbCrLf & _
      "Description:" & vbCrLf & ErrorDescription(ErrorCode) & vbCrLf
      MsgBox Txt, vbOKOnly Or vbExclamation, "dynNode Error"
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : FInit
' Purpose           : Initialize new dynNodes class with parent node reference
'*******************************************************************************
Friend Sub FInit(ParentClass As dynNode, Key As String, Text As String)
  Set m_MyParent = ParentClass    'save the parent node reference
  Init Key, Text                  'now initialize the node
End Sub

'*******************************************************************************
' Get Name          : KeyLC
' Purpose           : Get Lowercase Key value
'*******************************************************************************
Friend Property Get KeyLC() As String
  KeyLC = m_MyKeyLC
End Property

'*******************************************************************************
' Function Name     : LocalFindID
' Purpose           : Find a specified ID. Return Nothing if not found
'                   : This method is initially invoked on the Root Node from the
'                   : FindID() method, and subsequently from a like-named
'                   : node in the Nodes collection. This method is very
'                   : recursive and VERY fast.
'*******************************************************************************
Friend Function LocalFindID(ByVal ID As Long) As dynNode
  If ID = m_MyID Then                           'matches local?
    Set LocalFindID = Me                        'yes, so return ME
  Else
    Set LocalFindID = Me.Nodes.LocalFindID(ID)  'else scan children for it
  End If
End Function

'*******************************************************************************
' Function Name     : LocalFindKey
' Purpose           : Find a specified ID. Return Nothing if not found
'                   : This method is initially invoked on the Root Node from the
'                   : FindKey() method, and subsequently from the LocalFindKey
'                   : method in the Nodes collection. This method is very
'                   : recursive and VERY fast.
'*******************************************************************************
Friend Function LocalFindKey(lcKey As String) As dynNode
  If StrComp(lcKey, m_MyKeyLC, vbBinaryCompare) = 0 Then  'matches THIS node?
    Set LocalFindKey = Me                                 'yes, so return ME
  Else
    Set LocalFindKey = m_MyNodes.LocalFindKey(lcKey)       'else scan children for it
  End If
End Function

'*******************************************************************************
' Function Name     : LocalKeyExists
' Purpose           : Find a specified key. Return False if not found
'                   : This method is initially invoked on the Root Node from the
'                   : KeyExists() method, and subsequently from a like-named method
'                   : in the Nodes collection. This method is very recursive
'                   : and VERY fast.
'*******************************************************************************
Friend Function LocalKeyExists(lcKey As String) As Boolean
  If StrComp(lcKey, m_MyKeyLC, vbBinaryCompare) = 0 Then 'if local match...
    LocalKeyExists = True                                'we found one
  Else
    LocalKeyExists = Me.Nodes.LocalKeyExists(lcKey)      'find in sub-nodes
  End If
End Function

'*******************************************************************************
' Get Name          : Deleting
' Purpose           : Get state of deletion flag
'*******************************************************************************
Friend Property Get Deleting() As Boolean
  Deleting = m_MyDeleting
End Property

'*******************************************************************************
' Let Name          : Deleting
' Purpose           : Mark as being deleted when flag set
'
' Do not allow setting the Deleting Flag if the Locked Flag is set to True.
' Test is FALSE if both m_MyLocked = True And Flag = True
'*******************************************************************************
Friend Property Let Deleting(ByVal Flag As Boolean)
  If Not (m_MyLocked And Flag) Then m_MyDeleting = Flag
End Property

'*******************************************************************************
' Get Name          : ClassTerminating
' Purpose           : Return the class terminating flag
'*******************************************************************************
Friend Property Get ClassTerminating() As Boolean
  ClassTerminating = m_MyClassTerminating
End Property

'*******************************************************************************
' Set Name          : ClassTerminating
' Purpose           : Assign the class terminating flag
'*******************************************************************************
Friend Property Let ClassTerminating(ByVal Flag As Boolean)
  m_MyClassTerminating = Flag
End Property

'*******************************************************************************
' Function Name     : GetRootgIndex
' Purpose           : Get the gArray index value from the root node
'*******************************************************************************
Friend Function GetRootgIndex() As Long
  If m_MyID = 1 Then                      'if we are the root, then get the local index
    GetRootgIndex = m_g_Index
  Else
    GetRootgIndex = Me.Root.GetRootgIndex 'else find the root and get it
  End If
End Function

'*******************************************************************************
' Function Name     : GetgIndex
' Purpose           : Provide m_g_Index to Nodes
'*******************************************************************************
Friend Function GetgIndex() As Long
  GetgIndex = m_g_Index
End Function

Friend Sub SetgIndex(ByVal gIndex As Long)
  m_g_Index = gIndex
End Sub

'*******************************************************************************
' Subroutine Name   : ResetgIndex
' Purpose           : Reset the global index. This is called by the Swap and Move
'                   : functions when changing parent nodes.
'*******************************************************************************
Friend Sub ResetgIndex()
  m_g_Index = GetRootgIndex
End Sub

'*******************************************************************************
' Property Name     : ID
' Purpose           : Set the node's Unique ID number
'*******************************************************************************
Friend Property Let ID(ByVal NewID As Long)
  m_MyID = NewID
End Property

'*******************************************************************************
' Subroutine Name   : IDKey
' Purpose           : Use Defauild ID for key
'*******************************************************************************
Friend Sub IDKey()
    m_MyKey = CStr(m_MyID)                'use default (ID value) if null
    m_MyKeyLC = m_MyKey
End Sub

'*******************************************************************************
' Get Name          : Index
' Purpose           : Return the index of the node is it's parent list of children
'*******************************************************************************
Friend Property Get Index() As Long
  Index = m_Index
End Property


'*******************************************************************************
' Let Name          : Index
' Purpose           : Assign the index of the node is it's parent list of children
'*******************************************************************************
Friend Property Let Index(NewIndex As Long)
  m_Index = NewIndex
End Property

