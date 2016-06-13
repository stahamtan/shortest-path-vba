VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public nodeID As String
Public NodeName As String
Public neighbours As Scripting.Dictionary
Public closestNeighbour As Scripting.Dictionary           'Node ID & distance of the closest Neighbour Node
Public distance As Double
Public visited As Boolean
Private previousNode As Node

Private Sub Class_Initialize()
    Me.distance = 1E+308
    Me.visited = False
    Set Me.neighbours = CreateObject("Scripting.Dictionary")
    Set Me.closestNeighbour = CreateObject("Scripting.Dictionary")
End Sub

Property Let setPrevious(aNode As Node)
    Set previousNode = aNode
End Property

Property Get getPrevious() As Node
    Set getPrevious = previousNode
End Property

Public Sub addNeighbour(neighbourNode As Node, Optional distance As Double = 0)
    If Not Me.neighbours.Exists(neighbourNode) Then
        Me.neighbours.Add neighbourNode, distance
    End If
End Sub

Public Function getDistanceTo(neighbourNode As Node) As Double
    getDistanceTo = Me.neighbours(neighbourNode)
End Function

Public Sub getClosestNeighbour()
    Dim closestNodeID As String
    Dim closestNodeDistance As Double
    
    closestNodeID = Me.nodeID
    closestNodeDistance = 0
    
    For Each neighbourNode In Me.neighbours.Keys
        If Me.neighbours(neighbourNode) > closestNodeDistance Then
            closestNodeID = neighbourNode
            closestNodeDistance = Me.neighbours(neighbourNode)
        End If
    Next
    
    Me.closestNeighbour.Add closestNodeID, closestNodeDistance
    
End Sub
