Attribute VB_Name = "sfPathFinder222_3D"
' sfPathFinder222_3D v2.2.2     Royalty-free, last revised: October 14, 2026
' By Stéphane Fillon (fillon.s@wanadoo.fr)
' https://excel-pratique.com/fr/telechargements/auteur-71778

' DESCRIPTION :
' This module implements Dial's algorithm (1969), derived from Dijkstra's, optimized for 3D in VBA.
' It uses a byte-encoded representation for displacement vectors (Bytes from 0 to 26) and bypasses the
' overhead of function calls through targeted use of GoSub...Return within critical loops.

' This module contains the 3D pathfinding function: fPath,
' as well as a simple demo code using this function and its return value.

' This module is designed to be exported to other workbooks.
' Parameters that need to be adjusted for your application are followed by a comment starting with '# to make them easy to find.

' If you reuse this module in your application, note that lines ending with the '// comment sequence can be removed if you only perform short calculations.

Option Explicit

Private Type point3D
    X As Integer
    Y As Integer
    Z As Integer
End Type

Private Type lp         ' lp: List of Points
    p3d() As point3D
    cp As Boolean       ' cp: Contains Points
End Type

' --- ADAPTATION SETTINGS FOR YOUR WORKBOOK ---
'Public WsGrid As Worksheet     '# Uncomment this line if you reuse this module in another workbook.
Private Const Ocol As Byte = 1  '# Column number of the grid's origin (top-left cell), (e.g., 5 for column "E")
Private Const Orow As Byte = 2  '# Row number of the grid's origin...

Private Const Ox As Byte = Ocol - 1
Private Const Oy As Byte = Orow - 1


' Nodes_IC(X, Y, Z) is a 3D array containing all the points of the grid
Public Nodes_IC() As Integer
Public ReadTime!, PathfindingTime! '//



Public Function fPath(ByVal NbX%, ByVal NbY&, ByVal NbZ&, ByVal StartAddress$, ByVal EndAddress$, ByVal PropagationMode As Byte) As Variant
' fPath handles pathfinding and returns an array of elementary movement vectors encoded as Bytes.
' The last element of this array is the vector from the starting point (the parent) to the first point of the path (its child).
' To use this function, start from the starting point and traverse the array in reverse order until reaching the destination point.


' Formulas to extract X, Y and Z from these Byte vectors:
' X = (vector mod 3) - 1
' Y = Int((vector Mod 9) / 3) - 1
' Z = Int(vector / 9) - 1
            
'                              Neighbor sharing
'          a face                   an edge                 a vertex
'    vector   X   Y   Z       vector   X   Y   Z       vector   X   Y   Z
'    ------------------       ------------------       -----------------
'     4       0   0  -1        1       0  -1  -1        0      -1  -1  -1
'     10      0  -1   0        3      -1   0  -1        2       1  -1  -1
'     12     -1   0   0        5       1   0  -1        6      -1   1  -1
'     14      1   0   0        7       0   1  -1        8       1   1  -1
'     16      0   1   0        9      -1  -1   0        18     -1  -1   1
'     22      0   0   1        11      1  -1   0        20      1  -1   1
'                              15     -1   1   0        24     -1   1   1
'                              17      1   1   0        26      1   1   1
'                              19      0  -1   1
'                              21     -1   0   1
'                              23      1   0   1
'                              25      0   1   1


' If no solution can be found, fPath returns an Empty Variant.

 
' In addition to the point's Individual Cost (IC), two other metrics must be considered:
' - CC: standing for Cumulative Cost: the sum of costs encountered by a point from the starting point along a shortest path.
' - edv: the elementary displacement vector connecting a point to its parent.

' During each main Do Loop iteration, find the point(s) with the lowest identical CC.
' These become current nodes used to compute the CC and path for their unvisited immediate neighbors,
' setting their parent and adding them to the CCs() Array.

    Dim CCs() As lp         ' CCs is an array whose index corresponds to a CC
    Dim ZV As lp            ' Zero Values
    Dim cn As point3D       ' Current Node
    Dim n As Byte           ' the Neighbor of the current node
    Dim CCcn&               ' Cumulative Cost of the Current Node
    Dim CCmax&              ' Cumulative Cost of the most expensive path found
    Dim ICmax&              ' Maximal Individual Cost in the grid
    Dim ICn%, CCn&          ' Individual Cost and Cumulative Cost of the Neighbor
    Dim Neighborhood()      ' Vectors describing the immediate neighborhood
    Dim edv As Byte         ' Elementary Displacement Vector derived from Neighborhood()
    Dim NeighborhoodP3D() As point3D  ' Same as Neighborhood encoded Point3D
    Dim edvP3D As point3D   ' Same as edv encoded Point3D
    Dim ubn As Byte         ' UBound Neighborhood
    Dim Xn%, Yn%, Zn%       ' X, Y and Z of the neighbor
    Dim Xsp%, Ysp%, Zsp%    ' X, Y and Z of the Start Point
    Dim Xep%, Yep%, Zep%    ' X, Y and Z of the End Point
    Dim a&, i&
    Dim xyzfa As point3D
    Dim amoppm As Boolean   ' Adjacent Mode Or Permeable Peripheral Mode
    ' Three-dimensional arrays Nodes_...(X, Y, Z)
    Dim Nodes_edv() As Byte   ' Elementary Displacement Vector for each Node
    Dim Nodes_rtbr() As Boolean ' rtbr : Remains To Be Reached
    
    

    Dim cpt&    '//
    Dim TimerDoEvents#, TimerTime# '//
    
    TimerTime = Timer  '//
    
    ReDim Nodes_rtbr(NbX + 1, NbY + 1, NbZ + 1)
    ReDim Nodes_edv(NbX + 1, NbY + 1, NbZ + 1)
    
    ' If the destination point is on an obstacle cell (IC < 0), fPath returns an Empty Variant, meaning: "No solution"
    If WsGrid.Range(EndAddress) < 0 Then
        ReadTime = 0: PathfindingTime = 0 '//
        Exit Function
    End If
    
    ' Entering identical start and end addresses triggers reading the sheet without performing pathfinding.
    If EndAddress = StartAddress Then GoTo WorkSheetReading
    
    Application.StatusBar = "Initialization..." '//
    
    On Error GoTo WorkSheetReading ' Capture error if Nodes_IC is not yet initialized and fall back to reading the worksheet
    ' Reset mode
    For Xn = 1 To NbX
        For Yn = 1 To NbY
            For Zn = 1 To NbZ
                ' If the node is not an obstacle, it Remains To Be Reached.
                If Nodes_IC(Xn, Yn, Zn) >= 0 Then Nodes_rtbr(Xn, Yn, Zn) = True
    Next Zn, Yn, Xn
    ReadTime = -1   '//


PathFinding:
    DoEvents    '//
    Application.StatusBar = "Calculating the shortest path... 0%"     '//
    
    On Error GoTo 0
    ' Identifying the Start Point
    xyzfa = xyzFromAddress(StartAddress, NbY)
    Xsp = xyzfa.X
    Ysp = xyzfa.Y
    Zsp = xyzfa.Z
    
    ' Identifying the End Point
    xyzfa = xyzFromAddress(EndAddress, NbY)
    Xep = xyzfa.X
    Yep = xyzfa.Y
    Zep = xyzfa.Z
    
    ' CCs initialization
    ' The starting point's CC is considered equal to zero. This CC serves as an index for the CCs() array.
    ReDim CCs(1)
    ' For now, CCs(0) contains only a single Point3D, the starting point -> p3d(0)
    ReDim CCs(0).p3d(0)
    ' Setting the x, y and z coordinates of the Point3D
    CCs(0).p3d(0).X = Xsp
    CCs(0).p3d(0).Y = Ysp
    CCs(0).p3d(0).Z = Zsp
    ' CCs(0) contains a point: the starting point
    CCs(0).cp = True
    ' The starting point is considered reached
    Nodes_rtbr(Xsp, Ysp, Zsp) = False
    
    ' PropagationMode: 0=Adjacent, 1=Permeable perimeter, 2=Sprite perimeter
    ' Neighborhood() = Neighborhood of a central node consisting of elementary displacement vectors (-1, 0, 1) in X, Y and Z, encoded as bytes
    If PropagationMode = 0 Then ' 6 nodes sharing a face with the current node (Adjacent)
        Neighborhood = Array(4, 10, 12, 14, 16, 22)
    Else ' 26 nodes around the current node (Perimeter)
        Neighborhood = Array(4, 10, 12, 14, 16, 22, _
                          1, 3, 5, 7, 9, 11, 15, 17, 19, 21, 23, 25, _
                          0, 2, 6, 8, 18, 20, 24, 26)
    End If
    ubn = UBound(Neighborhood)
    
    ' We convert the Neighborhood bytes into Point3D
    ReDim NeighborhoodP3D(ubn)
    For i = 0 To ubn
        edv = Neighborhood(i)
        ' Each of the X, Y, and Z components can be -1, 0, or 1
        NeighborhoodP3D(i).X = (edv Mod 3) - 1
        NeighborhoodP3D(i).Y = Int((edv Mod 9) / 3) - 1
        NeighborhoodP3D(i).Z = Int(edv / 9) - 1
    Next i
    
    
    amoppm = PropagationMode < 2
    TimerDoEvents = Timer '//
   
    ' MAIN (AND INFINITE) DIAL LOOP: at each iteration, CCcn increases or remains constant
    Do
        ' Iterating through all nodes with the same minimum CC, CCcn
        For a = 0 To UBound(CCs(CCcn).p3d)
            cpt = cpt + 1    ' Useful only for the progress indicator in the StatusBar '//
            ' Declaration of the current node; it will become the parent of all its remaining-to-be-reached neighbors
            cn = CCs(CCcn).p3d(a)
            
            ' Iterating through all nodes in the immediate neighborhood of the current node (cn)
            For n = 0 To ubn
                edvP3D = NeighborhoodP3D(n)
                ' if the neighbor remains to be reached...
                If Nodes_rtbr(cn.X + edvP3D.X, cn.Y + edvP3D.Y, cn.Z + edvP3D.Z) Then
                    ' Adding the X, Y AND Z components of the Current Node to those of the elementary displacement vector
                    Xn = cn.X + edvP3D.X ' =Xn + (-1, 0 or 1)
                    Yn = cn.Y + edvP3D.Y ' =Yn + (-1, 0 or 1)
                    Zn = cn.Z + edvP3D.Z ' =Zn + (-1, 0 or 1)
                    edv = Neighborhood(n)
                    
                    If amoppm Then
                        ' Adjacent and permeable perimeter modes
                        GoSub Calcul ' Prevents the overhead of function calls
                    Else
                        ' Sprite perimeter mode: examining surrounding obstacles for diagonal movement
                        Select Case edv
                            ' Sharing an edge
                            Case 9, 11, 15, 17  ' xy plane
                                If Nodes_IC(cn.X, Yn, cn.Z) >= 0 And Nodes_IC(Xn, cn.Y, cn.Z) >= 0 Then GoSub Calcul
                            Case 1, 7, 19, 25   ' yz plane
                                If Nodes_IC(cn.X, cn.Y, Zn) >= 0 And Nodes_IC(cn.X, Yn, cn.Z) >= 0 Then GoSub Calcul
                            Case 3, 5, 21, 23   ' xz plane
                                If Nodes_IC(cn.X, cn.Y, Zn) >= 0 And Nodes_IC(Xn, cn.Y, cn.Z) >= 0 Then GoSub Calcul
                            ' Sharing a vertex
                            Case 0, 2, 6, 8, 18, 20, 24, 26
                                If Nodes_IC(cn.X, Yn, Zn) >= 0 And Nodes_IC(cn.X, cn.Y, Zn) >= 0 And _
                                   Nodes_IC(Xn, cn.Y, Zn) >= 0 And Nodes_IC(Xn, cn.Y, cn.Z) >= 0 And _
                                   Nodes_IC(Xn, Yn, cn.Z) >= 0 And Nodes_IC(cn.X, Yn, cn.Z) >= 0 Then GoSub Calcul
                            ' Sharing a face
                            Case Else
                                GoSub Calcul
                        End Select
                    End If
                End If
            Next n ' we can move on to the next neighboring node
            
            ' if the arrival point has been reached...
            ' *** EXIT WITH SOLUTION ***
            If Not Nodes_rtbr(Xep, Yep, Zep) Then GoTo PathFound
            
        Next a ' we can move on to the next current node
        
        ' End of the first sub-loop: all neighbors of nodes sharing the same minimum CC (CCcn) have now been computed, and we know whether Zero Value Contains Points
        
        ' Management of Dial's algorithm cost buckets
        If ZV.cp Then
            ' If zero values were reached in the previous sub-loop, CCs(CCcn) retrieves the list of points from ZV, and CCcn is not incremented.
            ReDim CCs(CCcn).p3d(UBound(ZV.p3d))
            CCs(CCcn).p3d = ZV.p3d
            Erase ZV.p3d: ZV.cp = False    ' We reset ZV
        Else
            ' Otherwise, we clear the list of points in CCs(CCcn) which is no longer needed
            Erase CCs(CCcn).p3d
            ' and we search for the next CCs(CCcn) that is not empty of points
            For i = CCcn + 1 To UBound(CCs)
                If CCs(i).cp Then CCcn = i: Exit For ' if CCs(i) Contains Points
            Next i
            ' If the following condition is met, all reachable points in the grid have been visited without reaching the end point
            If i > UBound(CCs) Then  ' *** EXIT WITHOUT SOLUTION ***
                PathfindingTime = Timer - TimerTime '//
                Exit Function
            End If
        End If
        
        ' Updating progress in the status bar '//
        If Timer - TimerDoEvents > 5 Then  '//
            TimerDoEvents = Timer  '//
            DoEvents  '//
            Application.StatusBar = "Calculating the shortest path... " & Int(100 * (cpt / (NbX * NbY * NbZ))) & "%" '//
        End If  '//
        
        
    Loop  ' End of the main loop
    
Calcul:
    ' This neighbor will no longer need to be recalculated from another parent
    Nodes_rtbr(Xn, Yn, Zn) = False
    ' edv will allow tracing back from the arrival point to the starting point
    Nodes_edv(Xn, Yn, Zn) = edv
    ' Retrieving the Individual Cost of the Neighbor
    ICn = Nodes_IC(Xn, Yn, Zn)
    
    If ICn = 0 Then
        ' if the Individual Cost of the neighbor is equal to zero, add this point to the Zero Values points list and set its Contains Points boolean to True
        If ZV.cp Then i = UBound(ZV.p3d) + 1 Else i = 0: ZV.cp = True
        ReDim Preserve ZV.p3d(i)
        ZV.p3d(i).X = Xn
        ZV.p3d(i).Y = Yn
        ZV.p3d(i).Z = Zn
    Else    ' If the neighbor's individual cost is greater than 0
        ' the neighbor's Cumulative Cost is equal to the current node's CC plus the neighbor's Individual Cost
        CCn = CCcn + ICn
        If ICn > ICmax Then ICmax = ICn
        ' if the neighbor's CC is greater than the maximum CC known so far, resize CCs
        If CCn > CCmax Then CCmax = CCmax * 2 + ICmax: ReDim Preserve CCs(CCmax)
        ' Adding the neighbor to the CCs(CCn) points list and setting its boolean flag to True
        If CCs(CCn).cp Then i = UBound(CCs(CCn).p3d) + 1 Else i = 0: CCs(CCn).cp = True
        ReDim Preserve CCs(CCn).p3d(i)
        CCs(CCn).p3d(i).X = Xn
        CCs(CCn).p3d(i).Y = Yn
        CCs(CCn).p3d(i).Z = Zn
    End If
    Return
    
    
PathFound:
    Dim Path() As Variant
    ' Starting from the end point
    Xn = Xep: Yn = Yep: Zn = Zep
    i = -1
    ' Tracing back from the end point to the starting point
    Do Until Xn = Xsp And Yn = Ysp And Zn = Zsp
        ' Retrieving the elementary displacement vector edv
        edv = Nodes_edv(Xn, Yn, Zn)
        
        i = i + 1
        ReDim Preserve Path(i)
        Path(i) = edv
        
        ' Extracting the X, Y and Z components of edv and shifting Xn, Yn and Zn
        Xn = Xn - ((edv Mod 3) - 1)             ' = Xn - (-1, 0 or 1)
        Yn = Yn - (Int((edv Mod 9) / 3) - 1)    ' = Yn - (-1, 0 or 1)
        Zn = Zn - (Int(edv / 9) - 1)            ' = Zn - (-1, 0 or 1)
    Loop
                                                                      
    fPath = Path
    PathfindingTime = Timer - TimerTime '//
    Exit Function ' ******* EXIT WITH A SOLUTION *******
    
    
WorkSheetReading:
    Application.StatusBar = "Reading the Worksheet... 0%"
    
    ReDim Nodes_IC(NbX + 1, NbY + 1, NbZ + 1)
    TimerDoEvents = Timer
    ' Iterating through the entire grid
    For Zn = 1 To NbZ          ' Level by level
        a = Oy + (NbY + 1) * (Zn - 1)
        For Yn = 1 To NbY      ' Row by row
            For Xn = 1 To NbX  ' Column by column
                Nodes_IC(Xn, Yn, Zn) = WsGrid.Cells(Yn + a, Xn + Ox)
                ' If the node is not an obstacle, it Remains To Be Reached
                If Nodes_IC(Xn, Yn, Zn) >= 0 Then Nodes_rtbr(Xn, Yn, Zn) = True
        Next Xn, Yn
        
        ' Updating progress in the status bar '//
        If Timer - TimerDoEvents > 5 Then  '//
            TimerDoEvents = Timer  '//
            DoEvents  '//
            Application.StatusBar = "Reading the Worksheet... " & Int(100 * (Zn / NbZ)) & "%"  '//
        End If  '//
    Next Zn
    ReadTime = Timer - TimerTime '//
    
    ' Entering identical start and end addresses triggers reading the sheet without performing pathfinding
    If EndAddress = StartAddress Then
        PathfindingTime = -1 '//
        fPath = Array(13) ' Null vector [0, 0, 0]
        Exit Function
    End If
    
    TimerTime = Timer '//
    GoTo PathFinding
End Function

Public Function xyzFromAddress(ByVal Adresse$, ByVal NbY) As point3D
    Dim row&, Z&
    row = Range(Adresse).row
    Z = Int((row - Orow) / (NbY + 1)) + 1
    xyzFromAddress.X = Range(Adresse).Column - Ocol + 1
    xyzFromAddress.Y = row - ((Z - 1) * (NbY + 1)) - Orow + 1
    xyzFromAddress.Z = Z
End Function



Sub Demo()
    Dim NbX%, NbY%, NbZ%
    Dim StartAddress$, EndAddress$
    Dim PropagationMode As Byte
    Dim Path As Variant
    Dim X%, Y%, Z%
    Dim Vector As Byte
    Dim i&
    
    Set WsGrid = Feuil1     '# CodeName of the worksheet containing the grid
    NbX = 5                 '# Number of columns in the grid
    NbY = 6                 '# Number of rows in the grid
    NbZ = 3                 '# Number of levels in the grid
    StartAddress = "A2"     '#
    EndAddress = "E21"      '#
    PropagationMode = 0     '# 0=Adjacent, 1=Permeable perimeter, 2=Sprite perimeter
    
    
    Cells.Clear
    
    ' Randomizing a grid for demo purposes
    demo_GenererateGrid NbX, NbY, NbZ
    
    Range(StartAddress).Interior.Color = vbRed
    Range(EndAddress).Interior.Color = vbCyan
    
    ' Since the grid has been modified, forcing worksheet re-reading without performing pathfinding by setting identical start and arrival addresses
    Path = fPath(NbX, NbY, NbZ, StartAddress, StartAddress, 0)
    ' Calling the pathfinding procedure and storing the result in Path
    Path = fPath(NbX, NbY, NbZ, StartAddress, EndAddress, PropagationMode)
    ' Case of grids with no solution
    If IsEmpty(Path) Then MsgBox "Sans solution": Exit Sub
    ' Iterating through Path from the last element (start point vector) to the first (last vector before the end point)
    For i = UBound(Path) To 0 Step -1
        Vector = Path(i)                    ' Vector extraction
        X = X + (Vector Mod 3) - 1          ' Extracting and shifting the X component
        Y = Y + Int((Vector Mod 9) / 3) - 1 ' Extracting and shifting the Y component
        Z = Z + Int(Vector / 9) - 1         ' Extracting and shifting the Z component
        Range(StartAddress).Offset(Y + Z * (NbY + 1), X).Interior.Color = vbGreen
    Next i
    ' Recoloring in cyan the end cell that was previously highlighted in green as part of the path
    Range(EndAddress).Interior.Color = vbCyan
End Sub

Private Sub demo_GenererateGrid(ByVal NbX%, ByVal NbY&, ByVal NbZ&)
    Dim Z%
    Application.ScreenUpdating = False
    WsGrid.Cells.Clear
    For Z = 0 To NbZ - 1
        With WsGrid.Range(Cells(Z * (NbY + 1) + Orow, Ocol), Cells(Z * (NbY + 1) + Orow + NbY - 1, Ocol + NbX - 1))
            .Formula = "=INT(RAND()*51)"
            .Copy
            .PasteSpecial xlPasteValues
            
        End With
    Next Z
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    WsGrid.Cells(Orow, Ocol).Select
End Sub

