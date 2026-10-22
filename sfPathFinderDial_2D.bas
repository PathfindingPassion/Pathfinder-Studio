Attribute VB_Name = "sfPathFinder222_2D"
' sfPathFinder222_2D v2.2.2     Royalty-free, last revised: October 14, 2026
' By Stéphane Fillon (fillon.s@wanadoo.fr)
' https://excel-pratique.com/fr/telechargements/auteur-71778


' DESCRIPTION :
' This module implements Dial's algorithm (1969), derived from Dijkstra's, optimized for 2D in VBA.
' It uses a byte-encoded representation for displacement vectors (Bytes from 0 to 8) and bypasses the
' overhead of function calls through targeted use of GoSub...Return within critical loops.

' This module contains the 2D pathfinding function: fPath,
' as well as a simple demo code using this function and its return value.

' This module is designed to be exported to other workbooks.
' Parameters that need to be adjusted for your application are followed by a comment starting with '# to make them easy to find.

' If you reuse this module in your application, note that lines ending with the '// comment sequence can be removed if you only perform short calculations.


Option Explicit

Private Type Point2D
    x As Integer
    y As Integer
End Type

Private Type lp     ' lp: List of Points
    p2d() As Point2D
    cp As Boolean    ' cp: Contains Points
End Type

' --- ADAPTATION SETTINGS FOR YOUR WORKBOOK ---
'Public WsGrid As Worksheet     '# Uncomment this line if you reuse this module in another workbook.
Private Const Ocol As Byte = 1  '# Column number of the grid's origin (top-left cell), (e.g., 5 for column "E")
Private Const Orow As Byte = 2  '# Row number of the grid's origin...

Private Const Ox As Byte = Ocol - 1
Private Const Oy As Byte = Orow - 1

Public Nodes_IC() As Integer     ' Nodes_IC(X, Y) is a 2D array containing all the points of the grid
Public ReadTime!, PathfindingTime! '//



Public Function fPath(ByVal NbX&, ByVal NbY&, ByVal StartAddress$, ByVal EndAddress$, ByVal PropagationMode As Byte) As Variant
' fPath handles pathfinding and returns an array of elementary movement vectors encoded as Bytes.
' The last element of this array is the vector from the starting point (the parent) to the first point of the path (its child).
' To use this function, start from the starting point and traverse the array in reverse order until reaching the destination point.

' Formulas to extract X and Y from these Byte vectors:
'       X = (vector mod 3) - 1
'       Y = int(vector / 3) - 1

'       vector    X   Y           vector   X   Y
'       ---------------           --------------
'         1       0  -1             0     -1  -1
'         3      -1   0             2      1  -1
'         5       1   0             6     -1   1
'         7       0   1             8      1   1

' If no solution can be found, fPath returns an Empty Variant.

 
' In addition to the point's Individual Cost (IC), two other metrics must be considered:
' - CC: standing for Cumulative Cost: the sum of costs encountered by a point from the starting point along a shortest path.
' - edv: the elementary displacement vector connecting a point to its parent.

' During each main Do Loop iteration, find the point(s) with the lowest identical CC.
' These become current nodes used to compute the CC and path for their unvisited immediate neighbors,
' setting their parent and adding them to the CCs() Array.


    Dim CCs() As lp             ' CCs is an array whose index corresponds to a CC
    Dim ZV As lp                ' ZV Zero Values
    Dim cn As Point2D           ' Current Node
    Dim n As Byte               ' n the Neighbor of the current node
    Dim Xn%, Yn%, ICn%, CCn&    ' X, Y, CI et CC of the Neighbor
    Dim CCcn&                   ' Cumulative Cost of the Current Node
    Dim CCmax&                  ' Cumulative Cost of the most expensive path found
    Dim ICmax&                  ' Maximal Individual Cost in the grid
    Dim Xsp%, Xep%, Ysp%, Yep%  ' sp Start Point; ep End Point
    Dim Neighborhood()          ' Vectors describing the immediate neighborhood
    Dim edv As Byte             ' Elementary Displacement Vector derived from Neighborhood()
    Dim ubn As Byte             ' ubn UBound Neighborhood
    Dim amoppm As Boolean       ' Adjacent Mode Or Permeable Peripheral Mode
    Dim a&, i&
    Dim cpt& '//
    Dim TimerDoEvents#, TimerTime#  ' Timers '//
    ' Two-dimensional Arrays Nodes_...(X, Y)
    Dim Nodes_edv() As Byte   ' Elementary Displacement Vector for each Node
    Dim Nodes_rtbr() As Boolean ' rtbr : Remains To Be Reached
    
    ReDim Nodes_edv(NbX + 1, NbY + 1)
    ReDim Nodes_rtbr(NbX + 1, NbY + 1)
    
    TimerTime = Timer '//
    
    
    ' If the destination point is on an obstacle cell (IC < 0), fPath returns an Empty Variant, meaning: "No solution"
    If WsGrid.Range(EndAddress) < 0 Then
        ReadTime = 0: PathfindingTime = 0 '//
        Exit Function
    End If
    
    ' Entering identical start and end addresses triggers reading the sheet without performing pathfinding
    If EndAddress = StartAddress Then GoTo WorkSheetReading
    
    Application.StatusBar = "Initialization..." '//
    
    On Error GoTo WorkSheetReading ' Catch the error if Nodes_IC is not yet initialized and fallback to reading the worksheet
    ' Reset mode
    For Xn = 1 To NbX
        For Yn = 1 To NbY
            ' If the node is not an obstacle, it Remains To Be Reached
            If Nodes_IC(Xn, Yn) >= 0 Then Nodes_rtbr(Xn, Yn) = True
    Next Yn, Xn
    ReadTime = -1 '//
    
    
PathFinding:
    DoEvents '//
    Application.StatusBar = "Calculating the shortest path... 0%" '//
    
    On Error GoTo 0
    ' Identifying the Start Point
    Xsp = WsGrid.Range(StartAddress).Column - Ocol + 1
    Ysp = WsGrid.Range(StartAddress).row - Orow + 1
    ' Identifying the End Point
    Xep = WsGrid.Range(EndAddress).Column - Ocol + 1
    Yep = WsGrid.Range(EndAddress).row - Orow + 1

    
    ' CCs initialization
    ' The starting point's CC is considered equal to zero. This CC serves as an index for the CCs() array.
    ' CCcn = 0
    ReDim CCs(CCcn)
    ' For now, CCs(0) contains only a single Point2D, the starting point -> p2d(0)
    ReDim CCs(CCcn).p2d(0)
    ' Setting the x and y coordinates of the Point2D
    CCs(CCcn).p2d(0).x = Xsp
    CCs(CCcn).p2d(0).y = Ysp
    ' CCs(0) contains a point: the starting point
    CCs(CCcn).cp = True
    ' The starting point is considered reached
    Nodes_rtbr(Xsp, Ysp) = False


    ' PropagationMode: 0=Adjacent, 1=Permeable perimeter, 2=Sprite perimeter
    ' Neighborhood() = Neighborhood of a central node consisting of elementary displacement vectors (-1, 0, 1) in X and Y, encoded as bytes
    If PropagationMode = 0 Then ' 4 nodes sharing an edge with the current node (Adjacent)
        Neighborhood = Array(1, 3, 5, 7)
    Else ' 8 nodes around the current node (Perimeter)
        Neighborhood = Array(1, 3, 5, 7, 0, 2, 6, 8)
    End If
    ubn = UBound(Neighborhood)
    
    amoppm = PropagationMode < 2
    
    TimerDoEvents = Timer '//
    
    
    ' MAIN (AND INFINITE) DIAL LOOP: at each iteration, CCcn increases or remains constant
    Do
        ' Iterating through all nodes with the same minimum CC, CCcn
        For a = 0 To UBound(CCs(CCcn).p2d)
            cpt = cpt + 1 ' Useful only for the progress indicator in the StatusBar '//
            
            ' Declaration of the current node; it will become the parent of all its remaining-to-be-reached neighbors
            cn = CCs(CCcn).p2d(a)

            ' Iterating through all nodes in the immediate neighborhood of the current node (cn)
            For n = 0 To ubn
                edv = Neighborhood(n)
                ' Adding the X and Y components of the Current Node to those of the elementary displacement vector
                Xn = cn.x + (edv Mod 3) - 1     ' = cn.x + (-1, 0 or 1)
                Yn = cn.y + Int(edv / 3) - 1    ' = cn.y + (-1, 0 or 1)
                
                ' if the neighbor remains to be reached...
                If Nodes_rtbr(Xn, Yn) Then
                    
                    If amoppm Or edv Mod 2 <> 0 Then
                        ' Adjacent or Permeable perimeter mode or Nodes sharing an edge with the current node
                        GoSub Calcul
                    Else
                        ' Sprite perimeter mode: examining surrounding obstacles for diagonal movement (edv = 0, 2, 6 or 8)
                        If Nodes_IC(cn.x, Yn) >= 0 And Nodes_IC(Xn, cn.y) >= 0 Then GoSub Calcul
                    End If
                End If
            Next n  ' we can move on to the next neighboring node
            
            ' if the arrival point has been reached...
            ' *** EXIT WITH SOLUTION ***
            If Not Nodes_rtbr(Xep, Yep) Then GoTo PathFound
        
        ' we can move on to the next current node
        Next a
        ' End of the first sub-loop: all neighbors of nodes sharing the same minimum CC (CCcn) have now been computed, and we know whether Zero Value Contains Points
        
        ' Management of Dial's algorithm cost buckets
        If ZV.cp Then
            ' If zero values were reached in the previous sub-loop, CCs(CCcn) retrieves the list of points from ZV, and CCcn is not incremented.
            ReDim CCs(CCcn).p2d(UBound(ZV.p2d))
            CCs(CCcn).p2d = ZV.p2d
            Erase ZV.p2d: ZV.cp = False   ' We reset ZV
        Else
            ' Otherwise, we clear the list of points in CCs(CCcn) which is no longer needed
            Erase CCs(CCcn).p2d
            ' and we search for the next CCs(CCcn) that is not empty of points
            For i = CCcn + 1 To CCcn + ICmax + 1
                If CCs(i).cp Then CCcn = i: Exit For    ' if CCs(i) Contains Points
            Next i
            ' If the following condition is met, all reachable points in the grid have been visited without reaching the end point
            If i > CCcn + ICmax + 1 Then  ' *** EXIT WITHOUT SOLUTION ***
                PathfindingTime = Timer - TimerTime '//
                Exit Function
            End If
        End If
        
        ' Updating progress in the status bar '//
        If Timer - TimerDoEvents > 5 Then  '//
            TimerDoEvents = Timer  '//
            DoEvents  '//
            Application.StatusBar = "Calculating the shortest path... " & Int(100 * (cpt / (NbX * NbY))) & "%"  '//
        End If  '//
        
    Loop  ' End of the main loop
    
    
Calcul:
    ' This neighbor will no longer need to be recalculated from another parent
    Nodes_rtbr(Xn, Yn) = False
    ' edv will allow tracing back from the arrival point to the starting point
    Nodes_edv(Xn, Yn) = edv
    ' Retrieving the Individual Cost of the Neighbor
    ICn = Nodes_IC(Xn, Yn)

    If ICn = 0 Then
        ' if the Individual Cost of the neighbor is equal to zero, add this point to the Zero Values points list and set its Contains Points boolean to True
        If ZV.cp Then i = UBound(ZV.p2d) + 1 Else i = 0: ZV.cp = True
        ReDim Preserve ZV.p2d(i)
        ZV.p2d(i).x = Xn
        ZV.p2d(i).y = Yn
    Else    ' If the neighbor's individual cost is greater than 0
        If ICn > ICmax Then ICmax = ICn
        ' the neighbor's Cumulative Cost is equal to the current node's CC plus the neighbor's Individual Cost
        CCn = CCcn + ICn
        
        ' if the neighbor's CC is greater than the maximum CC known so far, resize CCs
        If CCn > CCmax Then
            a = CCmax * 2 + ICmax
            If a > 268435454 Then a = 268435454 ' Max allowable value for an array type List of Points
            CCmax = a: ReDim Preserve CCs(CCmax)
        End If
        
        ' Adding the neighbor to the CCs(CCn) points list and setting its boolean flag to True
        If CCs(CCn).cp Then i = UBound(CCs(CCn).p2d) + 1 Else i = 0: CCs(CCn).cp = True
        ReDim Preserve CCs(CCn).p2d(i)
        CCs(CCn).p2d(i).x = Xn
        CCs(CCn).p2d(i).y = Yn
    End If
    Return

PathFound:
    Dim Path() As Variant
    ' Starting from the end point
    Xn = Xep: Yn = Yep
    i = -1
    ' Tracing back from the end point to the starting point
    Do Until Xn = Xsp And Yn = Ysp
        ' Retrieving the elementary displacement vector edv
        edv = Nodes_edv(Xn, Yn)
        i = i + 1
        ReDim Preserve Path(i)
        Path(i) = edv
        
        ' Extracting the X and Y components of edv and shifting Xn and Yn
        Xn = Xn - ((edv Mod 3) - 1)   ' = Xn - (-1 ,0 or 1)
        Yn = Yn - (Int(edv / 3) - 1)  ' = Yn - (-1 ,0 or 1)
    Loop
  
    fPath = Path
    PathfindingTime = Timer - TimerTime '//
    Exit Function ' ******* EXIT WITH A SOLUTION *******

WorkSheetReading:
    Application.StatusBar = "Reading the Worksheet... 0%"  '//
    
    ReDim Nodes_IC(NbX + 1, NbY + 1)
    ' Iterating through the entire grid
    TimerDoEvents = Timer '//
    For Yn = 1 To NbY        ' Row by row
        For Xn = 1 To NbX    ' Column by column
            Nodes_IC(Xn, Yn) = WsGrid.Cells(Yn + Oy, Xn + Ox)
            ' If the node is not an obstacle, it Remains To Be Reached
            If Nodes_IC(Xn, Yn) >= 0 Then Nodes_rtbr(Xn, Yn) = True
        Next Xn
        
        ' Updating progress in the status bar '//
        If Timer - TimerDoEvents > 5 Then '//
            TimerDoEvents = Timer '//
            DoEvents '//
            Application.StatusBar = "Reading the Worksheet... " & Int(Yn * (100 / NbY)) & "%" '//
        End If '//
    Next Yn
    ReadTime = Timer - TimerTime '//
    
    ' Entering identical start and end addresses triggers reading the sheet without performing pathfinding
    If EndAddress = StartAddress Then
        PathfindingTime = -1 '//
        fPath = Array(4) ' Null vector [0, 0]
        Exit Function
    End If
    
    TimerTime = Timer '//
    GoTo PathFinding
End Function



Sub Demo()
    Dim i&
    Dim NbX%, NbY%
    Dim StartAddress$, EndAddress$
    Dim PropagationMode As Byte
    Dim Path As Variant
    Dim Vector As Byte
    Dim x%, y%

    Set WsGrid = Feuil1     '# CodeName of the worksheet containing the grid
    NbX = 10                '# Number of columns in the grid
    NbY = 10                '# Number of rows in the grid
    StartAddress = "A2"     '#
    EndAddress = "J11"      '#
    PropagationMode = 0     '# 0=Adjacent, 1=Permeable perimeter, 2=Sprite perimeter

    Cells.Clear
    Range(StartAddress).Interior.Color = vbRed
    Range(EndAddress).Interior.Color = vbCyan

    ' Randomizing a grid for demo purposes
    WsGrid.Range(StartAddress & ":" & EndAddress).Formula = "=INT(RAND()*51)"

    ' Since the grid has been modified, forcing worksheet re-reading without performing pathfinding by setting identical start and arrival addresses
    Path = fPath(NbX, NbY, StartAddress, StartAddress, 0)

    ' Calling the pathfinding procedure and storing the result in Path
    Path = fPath(NbX, NbY, StartAddress, EndAddress, PropagationMode)
    ' Case of grids with no solution
    If IsEmpty(Path) Then MsgBox "Sans Solution": Exit Sub

    ' Iterating through Path from the last element (start point vector) to the first (last vector before the end point)
    For i = UBound(Path) To 0 Step -1
        Vector = Path(i)               ' Vector extraction
        x = x + (Vector Mod 3) - 1     ' Extracting and shifting the X component
        y = y + Int(Vector / 3) - 1    ' Extracting and shifting the Y component
        Range(StartAddress).Offset(y, x).Interior.Color = vbGreen
    Next i

    ' Recoloring in cyan the end cell that was previously highlighted in green as part of the path
    Range(EndAddress).Interior.Color = vbCyan
End Sub

