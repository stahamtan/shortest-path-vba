Attribute VB_Name = "Main"
Const infinity = 1E+308

Function Dijkstra(theGraph As Graph, originNode As Node, destinationNode As Node) As String
    
    originNode.distance = 0
    
    n = theGraph.totalNodes
    Dim unvisitedNodes As New Collection
    
    'thegraph.nodesDictionary
    'Node.neighbours
    
    'Populate the unvisited nodes queue
    For Each nID In theGraph.nodesDictionary.Keys
        If Not theGraph.nodesDictionary(nID).visited Then
            unvisitedNodes.Add theGraph.nodesDictionary(nID), CStr(nID)
        End If
    Next
    
    Dim u As Object     'u is the node with the smallest distance from the unvisited queue
    Dim current As Object   'current node
    
    Do While unvisitedNodes.Count
    
        Set u = popShortest(unvisitedNodes)
            
            Set current = u
            current.visited = True
            
            Dim adjacentNode As New Node
            
            For Each adjacentNode In current.neighbours
                
                If Not adjacentNode.visited Then
                    newDist = current.distance + current.getDistanceTo(adjacentNode)
                    
                    If newDist < adjacentNode.distance Then
                        adjacentNode.distance = newDist
                        adjacentNode.setPrevious = current
                    End If
                
                End If
                
            Next
        
        
        Do While unvisitedNodes.Count
            unvisitedNodes.Remove unvisitedNodes.Count
        Loop
        For Each nID In theGraph.nodesDictionary.Keys
            If Not theGraph.nodesDictionary(nID).visited Then
                unvisitedNodes.Add theGraph.nodesDictionary(nID), CStr(nID)
            End If
        Next
    Loop
    
    path = ""
    path = path & shortestPath(destinationNode)
    
    Dijkstra = path
End Function

Function popShortest(unvisitedQueue As Collection) As Node
    shortestDistance = infinity
    For Each n In unvisitedQueue
        If n.distance <= shortestDistance Then
            shortestDistance = n.distance
            Set popShortest = n
        End If
    Next
End Function

Function shortestPath(targetNode As Node) As String
    Dim path As String
    Dim current As Object
    
    Dim nodeList As Object  'Array list to store node ID's
    Set nodeList = CreateObject("System.Collections.ArrayList")
    
    Set current = targetNode
    ''path = path & targetNode.nodeID
    nodeList.Add targetNode.nodeID
    Do While Not current.getPrevious Is Nothing
        ''path = path & " " & current.getPrevious.nodeID
        nodeList.Add current.getPrevious.nodeID
        Set current = current.getPrevious
    Loop
    
    Set current = Nothing
    If nodeList.Count > 1 Then
        path = ""
        For i = nodeList.Count - 1 To 1 Step -1
            path = path & NameIDConverter(False, CStr(nodeList(i))) & " --> "
        Next
        path = path & NameIDConverter(False, CStr(nodeList(0)))
    Else
        path = "No path found!!"
    End If
    
    ''path = StrReverse(path)
    shortestPath = path
End Function

Sub testDrive()
    Dim aGraph As New Graph
    
    '!!!!! Comment out the following line, if not reading the graph from a database table
    Set aGraph = getGraph()
    
    '!!!!! Uncomment the following lines if not reading the graph from a database table
'    aGraph.addNode ("A")
'    aGraph.addNode ("B")
'    aGraph.addNode ("C")
'    aGraph.addNode ("D")
'    aGraph.addNode ("E")
'    aGraph.addNode ("F")
'    aGraph.addNode ("G")
'
'    aGraph.addEdge "A", "B", 2
'    aGraph.addEdge "A", "C", 8
'    aGraph.addEdge "A", "D", 2
'    aGraph.addEdge "B", "C", 5
'    aGraph.addEdge "B", "D", 3
'    aGraph.addEdge "C", "E", 5
'    aGraph.addEdge "C", "D", 4
'    aGraph.addEdge "D", "E", 1
    
    Debug.Print Dijkstra(aGraph, aGraph.getNode("A"), aGraph.getNode("F"))
    
    
End Sub

Function getGraph() As Graph
    Dim aGraph As New Graph
    
    Dim nodeRS As Object
    Dim edgeRS As Object
    
    Dim DBPath As String
    Dim sql As String
    
    DBPath = ThisWorkbook.path + "\YourShortestPath_DB.accdb"
    sql = "SELECT * FROM Node"
    Set nodeRS = getRS(sql, DBPath)
    
    Do While Not nodeRS.EOF
        aGraph.AddNode nodeRS.Fields(0)
        nodeRS.movenext
    Loop
    
    sql = "SELECT * FROM Graph"
    Set edgeRS = getRS(sql, DBPath)
    
    Do While Not edgeRS.EOF
        ''Debug.Print edgeRS.Fields(0), edgeRS.Fields(1), edgeRS.Fields(2)
        aGraph.addEdge edgeRS.Fields(0), edgeRS.Fields(1), edgeRS.Fields(2)
        edgeRS.movenext
    Loop
    
    ''Debug.Print "Done adding edges..!!"
    
    Set nodeRS = Nothing
    Set edgeRS = Nothing
    
    Set getGraph = aGraph
    
End Function

Function getRS(sqlQuery As String, DBFullpath As String) As Object
    'This function returns a resultset by running the given sql query on the given Access database
    'This is mainly to avoid lots of duplicate code
    
    Dim cn As Object    'Database connection
    Dim rs As Object    'Result set
    Dim strConnection As String
    
    Set cn = CreateObject("ADODB.Connection")
    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & DBFullpath & ";"
    cn.Open strConnection
    Set rs = cn.Execute(sqlQuery)
    
    Set getRS = rs
End Function

Function NameIDConverter(NameToID As Boolean, inputText As String) As String

    Dim sql As String
    Dim DBFullpath As String
    Dim rs As Object
    
    DBFullpath = ThisWorkbook.path + "\YourShortestPath_DB.accdb"
    
    If NameToID Then
        sql = "SELECT NodeID FROM Node WHERE Name = '" & inputText & "'"
    Else
        sql = "SELECT Name FROM Node WHERE NodeID = '" & inputText & "'"
    End If
    
    Set rs = getRS(sql, DBFullpath)
    
    NameIDConverter = rs.Fields(0)
    
End Function
