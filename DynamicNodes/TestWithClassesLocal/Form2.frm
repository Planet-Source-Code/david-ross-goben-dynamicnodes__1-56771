VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RootNode As dynNode        'define new root node
Dim FSO As FileSystemObject    'standard system I/O object

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Set aside FSO
'*******************************************************************************
Private Sub Form_Load()
  Dim Total As Long, Folders As Long
  
  If MsgBox("You could pop (and eat) some popcorn in the time it takes" & vbCrLf & _
            "to simply read an entire drive. Do you want to continue?", _
            vbYesNo Or vbQuestion, "Continue?") = vbNo Then
    Unload Me
    Exit Sub
  End If
  Set FSO = New FileSystemObject
  DrillDrive "C"                  'read the file/folder contents of the C drive
  'we can now do whatever we want with the node list branching from RootNode
  Total = RootNode.NodeCount        'get number of nodes under RootNode (plus RootNode)
  Folders = RootNode.MarkerCount   'Get number of Marked node (we marked as folders)
  
  MsgBox "Collected " & CStr(Folders) & " folders, and " & _
         CStr(Total - Folders) & " files (" & CStr(Total) & " total)"
  
  RootNode.Nodes.Clear              'remove ALL children
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Release FSO resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set RootNode = Nothing                    'release resources
  Set FSO = Nothing
End Sub

Private Sub DrillDrive(DriveLetter As String)
  Dim Drv As String
  
  Drv = DriveLetter & ":\"   'build a full filepath
  Set RootNode = CreateNewRoot(Drv, Drv)    'create new root-level node
  RootNode.Marker = True                    'mark it (we will mark all folders)
                                            'this does not 'make' it a folder, but
                                            'we are simply using this as a personal
                                            'reference. This can be used for anything.
  EnumerateNodes RootNode                   'enumerate drive folders into it
End Sub

'*******************************************************************************
' Subroutine Name   : EnumerateNodes
' Purpose           : Enumerate folders of drive into dynNode tree (recursive)
'
' Note that this routine calls itself repeatedly. This is a technique called
' recursion that is one of the coolest methods around for drilling through any
' branching-type system (unless you miss-code it and do not debug-step through it
' to ensure that it works). Then the Vulcan Nerve Pinch (Ctrl-Alt-Delete with
' one hand) may be required.
'*******************************************************************************
Private Sub EnumerateNodes(Parent As dynNode)
  Dim Fld As Folder     'Folder object from FSO classes
  Dim Fil As File       'File object from FSO classes
  Dim cNd As dynNode    'our local dynamic node
  Dim Ppath As String   'parent path (speed referencing by not using properties)
  
  Ppath = Parent.Key                              'get parent path
'
' build folders contained in Parent. Use On Error Resume Next in case we
' encounter special system-level protected folder such as those on NT/XP
' platforms that will let you see them, but doe-nah thoucha it!
'
  On Error Resume Next                            'in case of special protected folders
  For Each Fld In FSO.GetFolder(Ppath).SubFolders 'get all folders
    Set cNd = Parent.Nodes.Add(, dynNodeChild, Fld.Path & "\", Fld.Name)   'add a node
    cNd.Marker = True                             'tag it as user-marked (folder)
    EnumerateNodes cNd                            'enumerate each folder
  Next Fld
'
' build files contained in Parent (doing this into a TreeView node listwould hang
' the system on my computer, as a TreeView is limited to 32,000 entries. dynNodes
' allow up to 2,147,483,647 nodes (but who would be insame enough to try an make
' a branch list that long -- at least until personal computers are replaced by
' personal super-computers.
'
  For Each Fil In FSO.GetFolder(Ppath).Files
    Call Parent.Nodes.Add(, dynNodeChild, Fil.Path, Fil.Name)  'add a node
  Next Fil
End Sub

