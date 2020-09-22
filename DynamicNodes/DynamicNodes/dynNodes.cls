VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dynNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' dynNodes - Nodes Collection Class module
' Copyright © 2004 by David Goben. All rights reserved.
'
' This class is a support class for the dynNode class. A dynNodes collection
' object is created by every new dynNode object.
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'NOTE: Set Instancing in the dynNode class properties to MultiUse. If you are testing
'      these class files as inclusions in a non-DLL test project, VB will auto-convert
'      their instancing to Private. Hence, when you reload them into the ActiveX DLL
'      project, you will have to manutally convert them back.
'_______________________________________________________________________________
'  Property:        I/O:  Returns:   'Description:
'  Add              ---   dynNode    'Add a node to this node's dynNodes list
'  Count            G     Long       'Return number of nodes in dynNodes list
'  Clear            ---   ---        'release all branch's nodes
'  InitAutoForNext  ---   ---        'Init Item list for sequential gathering
'* Item             G     dynNode    'access a child of this node (DEFAULT property)
'  Remove           L     ---        'remove a child of this node
'_______________________________________________________________________________
'NOTE: G=GET, S=SET, L=LET, FG=Friend GET, FS=Friend SET, FL=Friend LET

'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
' Varaible local to this class
'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
Private m_MyCount As Long             'number of child nodes
Private m_MyNodes() As dynNode        'list of child nodes
Private m_MyParent As dynNode         'reference to this node collection's parent Node
Private m_Initialized As Boolean      'TRUE when this item initialized
Private m_MySorted As Boolean         'TRUE if list to be sorted
Private m_MySortDescending As Boolean 'TRUE if list to be sorted in descending order
Private m_LocalIndex As Long          'internal indexer, simulate "For Each"

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Subroutine Name   : Class_Initialize
' Purpose           : Initialize dynNodes Class
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Sub Class_Initialize()
  m_LocalIndex = 0
  m_MyCount = 0
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Subroutine Name   : Class_Terminate
' Purpose           : Terminate dynNodes Class
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Sub Class_Terminate()
  m_MyParent.ClassTerminating = True  'force deleting even Locked nodes
  Clear                               'erase all child nodes and branches
End Sub

'*******************************************************************************
' Function Name     : Add
' Purpose           : Add a new entry into the list
'*******************************************************************************
Public Function Add(Optional Relative As Variant = 0&, _
                    Optional Relationship As dynNodeRelationship = dynNodeChild, _
                    Optional Key As String = vbNullString, _
                    Optional Text As String = vbNullString) As dynNode
  Dim pIndex As Long, Index As Long, Count As Long
  Dim Rel As dynNodeRelationship
  Dim Node As dynNode
'
' if the dynNodes object is not initialized, then generate an error
'
  If Not m_Initialized Then
    m_MyParent.SetErrEvent dynNodeErrNotInitialized
    Exit Function
  End If
'
' reset local index for automatic incremental Item() selection
'
  m_LocalIndex = 0
'
' grab the desired relationship for the new node (child/sibling)
'
  Rel = Relationship
'
' now check Variant variable IndexOrKey for being a KEY (String)
'
  If VarType(Relative) = vbString Then
'
' ---KEY SPECIFIED-----------------------------------------------------------------
'
    Set Node = m_MyParent.FindKey(CStr(Relative))     'find the node for KEY
    If Node Is Nothing Then                           'found this key?
      m_MyParent.SetErrEvent dynNodeErrKeyInvalid     'no, so we are in error
      Exit Function
    End If
'
' if we are not creating a child node, then we need to get the index of the node
' in it's parent's collection, so that a sibling command can position itself.
' Variable Node will be set to the parent of the object to be created
'
    If Rel <> dynNodeChild Then                       'if not a child, get Node index
      If Node.ID = 1 Then                             'but if we are the root...
        Node.SetErrEvent dynNodeErrRootCanHaveNoSiblings  'we cannot have siblings...
        Exit Function
      End If
      Index = Node.Index                              'else grab node's index value
      Set Node = Node.Parent                          'Node = parent of siblings
    End If
'
' ---INDEX SPECIFIED (or was blank)------------------------------------------------
' grab the index value, Trap any error types than cannot convert to Long
'
  Else
    On Error Resume Next                              'in case of conversion error
    Index = CLng(Relative)                            'convert to a Long (0=blank)
    If Err.Number Then                                'conversion error?
      m_MyParent.SetErrEvent dynNodeErrKeyInvalid     'yes, so set error...
      Exit Function                                   'and leave
    End If
    On Error GoTo 0                                   'clear error trap
'
' check for a valid range within the collection (0=THIS; specifying Node)
'
    If Index <> 0 Then                                    'in case m_MyCount==0
      If Index < 0 Or Index >= m_MyCount Then             'index has a valid range?
        m_MyParent.SetErrEvent dynNodeErrIndexOutOfRange  'no, so error
        Exit Function
      End If
    End If
'
' check for defaulting to the specifying node (0) (specifying node = Node invoking Add())
'
    If Index = 0 Then                 'if index=0, create in THIS' collection
'
' If creating a child, set Node to the specifying node (m_MyParent)
' If creating a sibling, set Node to the specifying node's parent (m_MyParent.Parent)
'
      If Rel = dynNodeChild Then      'creating child?
        Set Node = m_MyParent         'yes, so we create new node in THIS collection
      Else  'sibling commands...
        Set Node = m_MyParent.Parent  'otherwise, set THIS' parent node
        If Node Is Nothing Then       'if Node was Root-Level Node, then error...
          m_MyParent.SetErrEvent dynNodeErrRootCanHaveNoSiblings
          Exit Function
        End If
      End If
'
' Else index is specified...
' If creating a child, create it as a child of the indexed node,
' otherwise we are creating a sibling of the indexed node.
' Node will be set to the parent of the new node being created
'
    Else
      Set Node = m_MyNodes(Index - 1) 'set Node to indexed node
      If Rel <> dynNodeChild Then     'if sibling of indexed node...
        Set Node = Node.Parent        'then set Node to indexed node's parent
      End If
    End If
  End If
'
' now figure insert position in the collection
'
  Count = Node.Children           'get number of children for container node
  If m_MySorted Then              'if sorting, then positioning does not matter...
    pIndex = Count                'so default to top for append (faster insertion)
  ElseIf Count = 0 Then           'if no data in list, then default to first posn
    pIndex = 0
  Else
    Select Case Rel               'else check relationship
      Case dynNodeFirst           'first in container's list
        pIndex = 0
      Case dynNodeLast
          pIndex = Count          'last in container's list
      Case dynNodeNext            'next after selection
        pIndex = Index            '(index = actual index+1)
      Case dynNodePrevious        'previous to index
        pIndex = Index - 1
      Case dynNodeChild           'child of selected node
        pIndex = Count            'if child, add to last of list (same as last)
    End Select
  End If
'
' if a key is defined, see if the key already exists in the child nodes
'
  If Len(Trim$(Key)) <> 0 Then                            'something to test?
    If m_MyParent.Root.KeyChecks Then                     'if checking keys permitted
      If m_MyParent.KeyExists(Key) Then                   'key found?
        m_MyParent.SetErrEvent dynNodeErrKeyAlreadyExists 'yes, so error
        Exit Function
      End If
    End If
  End If  'if nothing to test, use unique ID# defined below
'
' increment node array by 10's to make room (and speed processing)
'
  If m_MyCount Mod 10 = 0 Then               'bump 10 at a time
    ReDim Preserve m_MyNodes(m_MyCount + 10) 'NOTE: m_MyCount is already 1 higher
  End If
'
' shift items up to make room for the next item
'
  If pIndex < m_MyCount Then                              'if not at end (top)
    For Index = m_MyCount - 1 To pIndex Step -1           'shift others up...
      Set m_MyNodes(Index + 1) = m_MyNodes(Index)         'to make room
    Next Index
  End If
  Set m_MyNodes(pIndex) = New dynNode                     'add a new node
'
' set user data to node, and save parent node
'
  m_MyNodes(pIndex).FInit m_MyParent, Key, Text           'initialize new node
  m_MyCount = m_MyCount + 1                               'bump node count
'
' sort, and if not sort, then reset relationships between nodes
'
  If m_MySorted Then                        'sort?
    SortChildren False, m_MySortDescending  'yes (and call ResetRelationships)
  Else
    Call ResetRelationships                 'else simply reset next/prev indexes
  End If
'
' return new node reference
'
  Set Add = m_MyNodes(pIndex)
End Function

'*******************************************************************************
' Get Name          : Count
' Purpose           : Return the number of items in this collection
'*******************************************************************************
Public Property Get Count() As Long
  m_LocalIndex = 0                  'reset local index
  Count = m_MyCount                 'grab collection counter
End Property

'*******************************************************************************
' Subroutine Name   : Clear
' Purpose           : Clear all child nodes
'*******************************************************************************
Public Sub Clear()
  Dim Index As Long
  Dim Lst() As dynNode
  Dim Cnt As Long
  
  m_LocalIndex = 0                          'reset local index for auto-select
  
  If CBool(m_MyCount) Then                  'if something to clear...
    ReDim Lst(m_MyCount - 1) As dynNode     'set aside temporary list
    Cnt = 0                                 'init counter into temp list
    For Index = 0 To m_MyCount - 1          'loop through all
      With m_MyNodes(Index)
        If Not .Locked Or .ClassTerminating Then
          .Deleting = True                  'tag target to delete
          .Nodes.Clear                      'clear its child nodes
        End If
        If .Deleting Then                   'if nothing was locked
          Set m_MyNodes(Index) = Nothing    'remove item resources
        Else
          Set Lst(Cnt) = m_MyNodes(Index)   'save node reference
          Cnt = Cnt + 1                     'count a reference saved
        End If
      End With
    Next Index
'
' revitalize any locked items (Cnt <> 0)
'
    If CBool(Cnt) Then                      'if something was NOT cleared
      If Cnt <> m_MyCount Then              'if nodes have been cleared...
        m_MyCount = Cnt                     'set new count
        Cnt = Cnt - 1                       'back off for Ubound
        For Index = 0 To Cnt                'reorder uncleared child items
          Set m_MyNodes(Index) = Lst(Index)
        Next Index
        ReDim Preserve m_MyNodes((Cnt \ 10) * 10 + 10) As dynNode 'resize array
        Me.ResetRelationships               'reset their relationships
      End If
      m_MyParent.Deleting = False           'ensure parent turned off from deleting
                                            'because children still exist
    Else
      Erase m_MyNodes                       'else erase array if nothing locked
      m_MyCount = 0                         'and reset the child count to 0
    End If
  End If
End Sub

'*******************************************************************************
' Function Name     : InitAutoForNext
' Purpose           : Initialize Automatic For-Next processing on the Item List
'*******************************************************************************
Public Function InitAutoForNext() As Boolean
  If m_Initialized Then
    m_LocalIndex = 0                          'reset local index for auto-select
  Else
    m_MyParent.SetErrEvent dynNodeErrIndexOutOfRange
  End If
  InitAutoForNext = m_Initialized
End Function

'*******************************************************************************
' Get Name          : Item
' Purpose           : Return a dynNode item from the list
'
' This is the Default property for this class. With this Class active in the IDE,
' Select TOOLS/Procedure Attributes.../ADVANCED to see how this property was set:
' Set Item as NAME, and Procedure ID to (Default). Property Category: List
'*******************************************************************************
Public Property Get Item(Optional IndexOrKey As Variant = 0&) As dynNode
Attribute Item.VB_Description = "Class cNodes Default Property"
Attribute Item.VB_ProcData.VB_Invoke_Property = ";List"
Attribute Item.VB_UserMemId = 0
'
' if not initialized, then do not let anything happen
'
  If Not m_Initialized Then
    m_MyParent.SetErrEvent dynNodeErrNotInitialized
    Exit Property
  End If
  
  If CBool(m_MyCount) Then                                  'if something to do
    If VarType(IndexOrKey) = vbString Then                  'used specified a KEY
      m_LocalIndex = 0                                      'reset local index
      Set Item = m_MyParent.Root.FindKey(CStr(IndexOrKey))  'find it
      If Item Is Nothing Then                               'found?
        m_MyParent.SetErrEvent dynNodeErrKeyInvalid         'no, so we have an error
      End If
      Exit Property                                         'leave string portion
    Else
      Dim Index As Long                                     'numeric
      On Error Resume Next
      Index = CLng(IndexOrKey)                              'make a Long
      If Err.Number Then                                    'data error...
        m_MyParent.SetErrEvent dynNodeErrKeyInvalid         'set error
        Exit Property                                       'and leave
      End If
      On Error GoTo 0
      If Index = 0 Then                                     'use default?
        If m_LocalIndex >= m_MyCount Then Exit Property     'return nothing if out of range
        Set Item = m_MyNodes(m_LocalIndex)                  'get object if in range
        m_LocalIndex = m_LocalIndex + 1                     'point to next in list
        Exit Property
      ElseIf Index > 0 And Index <= m_MyCount Then          'valid range?
        m_LocalIndex = 0                                    'reset local index
        Set Item = m_MyNodes(Index - 1)                     'yes, so return object
        Exit Property
      End If
    End If
  End If
'
' here if invalid ANYTHING
'
  m_MyParent.SetErrEvent dynNodeErrIndexOutOfRange
End Property

'*******************************************************************************
' Subroutine Name   : Remove
' Purpose           : Remove an item from the list
'*******************************************************************************
Public Sub Remove(ByVal IndexOrKey As Variant)
  Dim Idx As Long, Index As Long, Cnt As Long
'
' reset local index for auto-select
'
  m_LocalIndex = 0
'
' if not initialized, then do not let anything happen
'
  If Not m_Initialized Then
    m_MyParent.SetErrEvent dynNodeErrNotInitialized
    Exit Sub
  End If
  
  If CBool(m_MyCount) Then                              'if something to do
'
' handle Key value. We will first have to find the Node, then find its index within
' it's parent's list, and then Remove it from the parent's list
'
    If VarType(IndexOrKey) = vbString Then              'used specified a KEY
      Dim Node As dynNode
      With m_MyParent
        Set Node = .Root.FindKey(CStr(IndexOrKey))      'find it
        If Node Is Nothing Then                         'found?
          .SetErrEvent dynNodeErrKeyInvalid             'no, so we have an error
          Exit Sub
        Else
          With Node
          '
          ' if this is the ROOT node (Parent Is Nothing)...
          '
            If .ID = 1 Then                             'if this is the ROOT Node
              If Not .Locked Or .ClassTerminating Then  'if we can delete this item
                .Deleting = True                        'indicate that we are deleting
                .Nodes.Clear                            'clear sub-nodes
              End If
              If .Deleting Then Set Node = Nothing      'render self to ethers
              Exit Sub                                  'and leave
            End If
          End With
          Node.Parent.Nodes.Remove Node.Index           'remove it by index
          Exit Sub
        End If
      End With
    Else
'
' user specified a numeric value
'
      On Error Resume Next                              'error trap
      Index = CLng(IndexOrKey)                          'make a Long
      If Err.Number Then                                'data error...
        m_MyParent.SetErrEvent dynNodeErrKeyInvalid     'set error
        Exit Sub                                        'and leave
      End If
      On Error GoTo 0                                   'else all is well
    End If
  End If
'
' remove an item by index into the Nodes Collection list
'
  If Index > 0 And Index <= m_MyCount Then
    With m_MyNodes(Index - 1)
      If Not .Locked Or .ClassTerminating Then    'if we can delete items
        .Deleting = True                          'tag target to delete
        .Nodes.Clear                              'clear its child nodes
      End If
      If .Deleting Then                           'if Deleting flag still set
        Set m_MyNodes(Index - 1) = Nothing        'kill off node to remove
        For Idx = m_MyCount - 1 To Index Step -1  'now collapse the array
          Set m_MyNodes(Idx - 1) = m_MyNodes(Idx)
        Next Idx
        m_MyCount = m_MyCount - 1                 'drop back counter
        If CBool(m_MyCount) Then                  'if something still there...
'''          ReDim Preserve m_MyNodes(m_MyCount - 1) 'collapse array
          Cnt = ((m_MyCount - 1) \ 10) * 10 + 10  'compute new array size
          If Cnt <> UBound(m_MyNodes) Then        'will general dim change?
            ReDim Preserve m_MyNodes(Cnt) As dynNode 'resize array if so
          End If
          Call ResetRelationships                 'reset next/prev indexes
        Else
          Erase m_MyNodes                         'else empty, so erase
        End If
      Else
        m_MyParent.Deleting = False               'ensure parent turned off from deleting
                                                  'because children still exist
      End If
    End With
  End If
End Sub

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' FRIEND ROUTINES
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'*******************************************************************************
' Get Name          : Sorted
' Purpose           : Return TRUE if the list is sorted
'*******************************************************************************
Friend Property Get Sorted() As Boolean
  Sorted = m_MySorted
End Property

'*******************************************************************************
' Let Name          : Sorted
' Purpose           : Set the sorting flag. Sort as needed
'*******************************************************************************
Friend Property Let Sorted(ByVal SetSort As Boolean)
'Force sort only if node was not set to sorting before and we are wanting it
  If Not m_MySorted And SetSort Then
    SortChildren False, m_MySortDescending  'yes, so Sort
  End If
  m_MySorted = SetSort                      'set the actual flag
End Property

'*******************************************************************************
' Get Name          : SortDescending
' Purpose           : Return TRUE if the list is to sort in descending order
'*******************************************************************************
Friend Property Get SortDescending() As Boolean
  SortDescending = m_MySortDescending
End Property

'*******************************************************************************
' Let Name          : SortDescending
' Purpose           : Set the sorting flag. Sort as needed
'*******************************************************************************
Friend Property Let SortDescending(ByVal SetDescending As Boolean)
'Force sorting if node was set for sorting, and was not set for descending
  If m_MySorted And SetDescending And Not m_MySortDescending Then
    SortChildren False, True              'yes, so sort
  End If
  m_MySortDescending = SetDescending      'set the actual flag
End Property

'*******************************************************************************
' Subroutine Name   : Init
' Purpose           : Initialize new dynNodes class with parent node reference
'*******************************************************************************
Friend Sub Init(ParentClass As dynNode)
  Set m_MyParent = ParentClass
  m_Initialized = True
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
' Function Name     : GetNewID
' Purpose           : Return a unique ID number. This uses a static variable
'                   : so that all instances of the class will share this variable
'*******************************************************************************
Friend Function GetNewID() As Long
  Dim Index As Long
  
  Index = m_MyParent.GetRootgIndex
  g_ID(Index) = g_ID(Index) + 1 'bump ID index
  GetNewID = g_ID(Index)        'return new index
End Function

'*******************************************************************************
' Subroutine Name   : ResetRelationships
' Purpose           : Reset PreviousSibling and NextSibling values
'*******************************************************************************
Friend Sub ResetRelationships()
  Dim Idx As Long, Prev As Long, Nxt As Long
  
  Prev = -1                             'start-1
  Nxt = 1                               'start+1
  For Idx = 0 To m_MyCount - 1          'do while list
    With m_MyNodes(Idx)
      .Index = Nxt                      'save index
      .sPreviousSibling Prev            'set previous index
      If Nxt = m_MyCount Then Nxt = -1  'wrap if beyond end
      .sNextSibling Nxt                 'set next
    End With
    Prev = Prev + 1                     'bump indexes
    Nxt = Nxt + 1
  Next Idx
End Sub

'*******************************************************************************
' Subroutine Name   : SortChildren
' Purpose           : Shell/Metzner Sort the child nodes (VERY FAST)
'*******************************************************************************
Friend Function SortChildren(Optional SortAllGenerations As Boolean = False, _
                             Optional ReverseOrder As Boolean = False) As Boolean
  Dim IndexLo As Long, IncIndex As Long
  Dim HalfUp As Long, IndexHi As Long
  Dim HalfDown As Long, NumberofItems As Long
  Dim CompFlag As Integer
  Dim Tmp As dynNode
'
' if not enough items to sort, then nothing to do
'
  If m_MyCount = 0 Then Exit Function  'no need if we have 0
'
' if we have only one child, it is not required to sort the current collection
' but we will try sorting the single child node
'
  If m_MyCount > 1 Then                   'if more than one child
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
    NumberofItems = m_MyCount             'get number if items to sort            (N=Number of Items)
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
          If StrComp(m_MyNodes(IndexLo).Text, _
                     m_MyNodes(IndexHi).Text, vbTextCompare) = CompFlag Then  '   IF D(I)>D(L) THEN GOTO D
            Set Tmp = m_MyNodes(IndexLo) 'swap nodes                              T=D(I)
            Set m_MyNodes(IndexLo) = m_MyNodes(IndexHi)                       '   D(I)=D(L)
            Set m_MyNodes(IndexHi) = Tmp                                      '   D(L)=T
            IndexLo = IndexLo - HalfDown  'back up index                      '   I=I-M
          Else                                                                '   IF I>=1 THEN GOTO C
            IncIndex = IncIndex + 1       'else bump counter                  D:  J=J+1
            Exit Do                                                           '   IF J>K THEN GOTO A
          End If                                                              '   GOTO B
        Loop While IndexLo >= 0            'while more things to check
      Loop
    Loop
  End If
'
' sorting all generations forces sorting on all child nodes that are not sorted
' without setting the sort flag
'
  If SortAllGenerations Then
    For IndexLo = 0 To m_MyCount - 1              'process all children
      Call m_MyNodes(IndexLo).Nodes.SortChildren(True, ReverseOrder)
    Next IndexLo
  End If
'
' finish up
'
  Call ResetRelationships               'reset next/prev indexes
  SortChildren = True                   'report success
End Function

'*******************************************************************************
' Function Name     : LocalFindID
' Purpose           : Find a specified ID. Return Nothing if not found
'                   : This method is invokes from a node by the LocalFindID()
'                   : methond there, which was in turn invoked by the
'                   : Public FindID() method. This method is very recursive
'                   : and VERY fast.
'*******************************************************************************
Friend Function LocalFindID(ByVal ID As Long) As dynNode
  Dim Index As Long
'
' scan through each node and check if they have a match. Also check their
' child nodes. These matching routines recurse for fast node drilling
'
  Index = 0
  Do While Index <> m_MyCount
    Set LocalFindID = m_MyNodes(Index).LocalFindID(ID)  'check for ID
    If Not LocalFindID Is Nothing Then Exit Function    'FOUND!
    Index = Index + 1
  Loop
End Function

'*******************************************************************************
' Function Name     : LocalFindKey
' Purpose           : Find a specified key. Return Nothing if not found
'                   : This method is invokes from a node by the LocalFindKey()
'                   : methond there, which was in turn invoked by the
'                   : Public FindKey() method. This method is very recursive
'                   : and VERY fast.
'*******************************************************************************
Friend Function LocalFindKey(lcKey As String) As dynNode
  Dim Index As Long
'
' scan through each node and check if they have a match. Also check their
' child nodes. These matching routines recurse for fast node drilling
'
  Index = 0
  Do While Index <> m_MyCount
    Set LocalFindKey = m_MyNodes(Index).LocalFindKey(lcKey) 'check for key
    If Not LocalFindKey Is Nothing Then Exit Function 'FOUND!
    Index = Index + 1
  Loop
End Function

'*******************************************************************************
' Function Name     : LocalKeyExists
' Purpose           : Find a specified key. Return False if not found
'                   : This method is invokes from a node by the LocalKeyExists()
'                   : methond there, which was in turn invoked by the
'                   : Public KeyExists() method. This method is very recursive
'                   : and VERY fast.
'*******************************************************************************
Friend Function LocalKeyExists(lcKey As String) As Boolean
  Dim Index As Long
'
' scan through each node and check if they have a match. Also check their
' child nodes. These matching routines recurse for fast node drilling
'
  Index = 0
  Do While Index <> m_MyCount
    If m_MyNodes(Index).LocalKeyExists(lcKey) Then Exit Do 'found a match!
    Index = Index + 1
  Loop
  LocalKeyExists = Index <> m_MyCount 'if we exited loop early, we found a match
End Function

'*******************************************************************************
' Subroutine Name   : ReID
' Purpose           : Reindex nodes when branch becomes separate tree
'*******************************************************************************
Friend Sub ReID(ByVal gIndex As Long)
  Dim Index As Long

  For Index = 0 To m_MyCount - 1  'scan through all children
    m_MyNodes(Index).ReID gIndex  'set new global index (Invoke ReID in each Node)
  Next Index
End Sub

