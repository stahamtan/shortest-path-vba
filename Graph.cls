VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public nodesDictionary As Scripting.Dictionary
Public totalNodes As Integer

Private Sub Class_Initialize()
    Me.totalNodes = 0
    Set Me.nodesDictionary = CreateObject("Scripting.Dictionary")
End Sub

Public Sub AddNode(ByVal nodeID As String, Optional ByVal NodeName As String = vbNullString)
    Dim newNode As New Node
    
    newNode.nodeID = nodeID
    newNode.NodeName = NodeName
    
    Me.nodesDictionary.Add nodeID, newNode
    
    Me.totalNodes = Me.totalNodes + 1
End Sub

Public Sub addEdge(ByVal fromNodeID As String, ByVal toNodeID As String, Optional distance As Double = 0)
    
    Dim fromNodeExists As Boolean
    Dim toNodeExists As Boolean
    
    fromNodeExists = False
    toNodeExists = False
    
    'Check if the nodes exist
    For Each nodeID In Me.nodesDictionary.Keys
        If fromNodeID = nodeID Then
            fromNodeExists = True
        End If
        
        If toNodeID = nodeID Then
            toNodeExists = True
        End If
    Next
    
    'Add new nodes if not exist
    If Not fromNodeExists Then
        Me.AddNode fromNodeID
    End If
    
    If Not toNodeExists Then
        Me.AddNode toNodeID
    End If
    
    ' Add bidirectional edge
    Me.nodesDictionary(fromNodeID).addNeighbour Me.nodesDictionary(toNodeID), distance
    Me.nodesDictionary(toNodeID).addNeighbour Me.nodesDictionary(fromNodeID), distance
    
End Sub

Public Function getNode(ByVal nodeID As String) As Node
    Set getNode = Me.nodesDictionary(nodeID)
End Function

