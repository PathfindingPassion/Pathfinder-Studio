Attribute VB_Name = "sfPathFinder211"
' sfPathFinder211_2D v2.1.1     libre de droits, dernière révision : 12 oct 2026
' par Stéphane Fillon (fillon.s@wanadoo.fr)
' https://excel-pratique.com/fr/telechargements/auteur-71778

' DESCRIPTION :
' This module implements Dijkstra's algorithm (1959), optimized for 2D in VBA.
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
    y As Long
End Type

' --- ADAPTATION SETTINGS FOR YOUR WORKBOOK ---
'Public WsGrid As Worksheet             '# Uncomment this line if you reuse this module in another workbook.
Private Const Ocol As Byte = 1          '# Column number of the grid's origin (top-left cell), (e.g., 5 for column "E")
Private Const Orow As Byte = 2          '# Row number of the grid's origin...


Private Const Ox As Byte = Ocol - 1
Private Const Oy As Byte = Orow - 1

Public Nodes_IC() As Double ' Nodes_IC(X, Y) is a 2D array containing all the points of the grid
Public ReadTime!, PathfindingTime! '//


Public Function fPath(ByVal NbX&, ByVal NbY&, ByVal StartAddress$, ByVal EndAddress$, ByVal PropagationMode As Byte) As Variant
' fPath handles pathfinding and returns an array of elementary movement vectors encoded as Bytes.
' The last element of this array is the vector from the starting point (the parent) to the first point of the path (its child).
' To use this function, start from the starting point and traverse the array in reverse order until reaching the destination point.

' Formulas to extract X and Y from these Byte vectors:
'       X = (vector mod 3) - 1
'       Y = int(vector / 3) - 1

'               adjacent                diagonales
'           vector   X   Y           vector  X   Y
'           ---------------           --------------
'             1       0  -1             0     -1  -1
'             3      -1   0             2      1  -1
'             5       1   0             6     -1   1
'             7       0   1             8      1   1

' If no solution can be found, fPath returns an Empty Variant.
 
' In addition to the point's Individual Cost (IC), two other metrics must be considered:
' - CC: standing for Cumulative Cost: the sum of costs encountered by a point from the starting point along a shortest path.
' - edv: the elementary displacement vector connecting a point to its parent.
    
    Dim RP() As Point2D     ' List of Reached Points whose neighbors have not yet been evaluated
    Dim smCC() As Point2D   ' Same Minimum Cumulative Cost
    Dim NR() As Point2D     ' Newly Reached; nodes reached during the first sub-loop
    Dim cn As Point2D       ' Current Node from smCC()
    Dim n As Byte           ' n the Neighbor of the current node
    Dim CCn#                ' Cumulative Cost of the neighbor
    Dim Xn%, Yn%            ' X and Y of the Neighbor
    Dim Xsp%, Ysp%          ' X and Y of the Start Point
    Dim Xep%, Yep%          ' X and Y of the End Point
    Dim flg&  ' First Loop Gain: The number of cells reached during the first sub-loop
    Dim Neighborhood()      ' Vectors describing the immediate neighborhood
    Dim edv As Byte         ' Elementary Displacement Vector
    Dim ubn As Byte         ' UBound Neighborhood
    Dim temp#, gap&, t As Point2D ' Used in sorting NR()
    Dim amoppm As Boolean   ' Adjacent Mode Or Permeable Peripheral Mode
    Dim i&, m&, k&
    Dim cpt& '//
    Dim TimerTime#, TimerDoEvents# ' Timers '//
    ' Two-dimensional Arrays Nodes_...(X, Y)
    Dim Nodes_CC() As Double ' Stores cumulative costs for all grid points
    Dim Nodes_edv() As Byte ' Elementary Displacement Vector for each Node
    
    
    TimerTime = Timer  '//
    ReDim Nodes_CC(NbX + 1, NbY + 1)
    ReDim Nodes_edv(NbX + 1, NbY + 1)
    
    ' If the destination point is on an obstacle cell (IC < 0), fPath returns an Empty Variant, meaning: "No solution"
    If WsGrid.Range(EndAddress) < 0 Then
        ReadTime = 0: PathfindingTime = 0 '//
        Exit Function
    End If
    
    ' Entering identical start and end addresses triggers reading the sheet without performing pathfinding
    If EndAddress = StartAddress Then GoTo WorkSheetReading
    
    On Error GoTo WorkSheetReading ' Catch the error if Nodes_IC is not yet initialized and fallback to reading the worksheet
    ' Reset mode
    For Xn = 1 To NbX
        For Yn = 1 To NbY
            If Nodes_IC(Xn, Yn) >= 0 Then Nodes_CC(Xn, Yn) = -1 ' CC = -1: Node is not an obstacle yet and will be added to NR, then RP
    Next Yn, Xn
    ReadTime = -1  '//
    
    
PathFinding:
    DoEvents    '//
    Application.StatusBar = "Calculating the shortest path... 0%"  '//

    On Error GoTo 0
    ' Identifying the Start Point
    Xsp = WsGrid.Range(StartAddress).Column - Ocol + 1
    Ysp = WsGrid.Range(StartAddress).row - Orow + 1
    ' Identifying the End Point
    Xep = WsGrid.Range(EndAddress).Column - Ocol + 1
    Yep = WsGrid.Range(EndAddress).row - Orow + 1

    ' Initially, there is only one Current Node: the Start Point
    ReDim smCC(0)
    smCC(0).x = Xsp
    smCC(0).y = Ysp
    ' The Starting Point's Cumulative Cost is considered to be zero
    Nodes_CC(Xsp, Ysp) = 0
    
    ReDim RP(0)     ' RP initialization
    RP(0) = smCC(0) ' The first item in RP is the start point
    
    ' PropagationMode: 0=Adjacent, 1=Permeable perimeter, 2=Sprite perimeter
    ' Neighborhood() = Neighborhood of a central node consisting of elementary displacement vectors (-1, 0, 1) in X and Y, encoded as bytes
    If PropagationMode = 0 Then ' 4 nodes sharing an edge with the current node (Adjacent)
        Neighborhood = Array(1, 3, 5, 7)
    Else    ' 8 nodes around the current node (Perimeter)
        Neighborhood = Array(0, 2, 6, 8, 1, 3, 5, 7)
    End If
    ubn = UBound(Neighborhood)
    
    amoppm = PropagationMode < 2
    TimerDoEvents = Timer '//
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MAIN (AND INFINITE) LOOP                       '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Do
        flg = 0     ' Reset flg
        Erase NR    ' Reset NR()
        ''''''''''''''''''''''''''''''''''''''''''''''
        '  FIRST SUB-LOOP: NEIGHBORHOOD CALCULATION  '              '
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' Iterating through the Current Nodes of Same Minimum Cumulative Cost
        For m = 0 To UBound(smCC)
            cpt = cpt + 1   ' Useful only for the progress indicator in the StatusBar '//
            
            cn = smCC(m) ' Isolating the Current Node whose neighborhood will be calculated
            ' Iterating through all nodes in the immediate neighborhood of the current node (cn)
            For n = 0 To ubn
                edv = Neighborhood(n)  ' Isolating the Elementary Displacement Vector
                Xn = cn.x + (edv Mod 3) - 1  ' =nc.x + (-1, 0 or 1)
                Yn = cn.y + Int(edv / 3) - 1 ' =nc.y + (-1, 0 or 1)
                
                ' CC is < 0 for all unreached Nodes that are not obstacles
                If Nodes_CC(Xn, Yn) < 0 Then
                    If amoppm Or edv Mod 2 <> 0 Then
                        ' Adjacent or Permeable perimeter mode or Nodes sharing an edge with the current node
                        GoSub Calcul
                    Else
                        ' Sprite perimeter mode: examining surrounding obstacles for diagonal movement (edv = 0, 2, 6 or 8)
                        If Nodes_IC(cn.x, Yn) >= 0 And Nodes_IC(Xn, cn.y) >= 0 Then GoSub Calcul
                    End If
                End If
            ' we can move on to the next neighboring node
            Next n
        ' we can move on to the next current node
        Next m
          
        ' if the arrival point has been reached...
        ' *** EXIT WITH SOLUTION ***
        If Nodes_CC(Xep, Yep) >= 0 Then GoTo PathFound
       
       ' *** EXIT WITHOUT SOLUTION ***
        If UBound(RP) + flg = 0 Then
            PathfindingTime = Timer - TimerTime '//
            Exit Function
        End If
        
        
        ' Now sorting NR() in descending order
        ' If at least two nodes were newly reached... (flg = Ubound(NR) + 1)
        If flg > 1 Then
            If flg = 2 Then ' If NR contains only two items, a potential swap will suffice
                If Nodes_CC(NR(0).x, NR(0).y) < Nodes_CC(NR(1).x, NR(1).y) Then
                    t = NR(0)
                    NR(0) = NR(1)
                    NR(1) = t
                End If
            Else
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' 2ND SUB-LOOP: Shell sort if NR contains more than 2 elements
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                gap = 1
                ' Calculation of the initial interval (Knuth sequence)
                Do While gap < flg \ 3
                    gap = gap * 3 + 1
                Loop
                
                ' Insertion sort with successive increments
                While gap > 0
                    For i = gap To flg - 1
                        t = NR(i)
                        temp = Nodes_CC(t.x, t.y)
                        m = i
                        
                        ' Descending order: shift if the previous element is smaller
                        Do While m >= gap
                            If Nodes_CC(NR(m - gap).x, NR(m - gap).y) < temp Then
                                NR(m) = NR(m - gap)
                                m = m - gap
                            Else
                                Exit Do
                            End If
                        Loop
                        
                        NR(m) = t
                    Next i
                    gap = gap \ 3
                Wend
            End If
        End If ' END OF NR() SORTING
        
        
        If flg > 0 Then
            m = UBound(RP): k = UBound(NR)
            ' Resize RP to include the newly reached Point2D
            ReDim Preserve RP(m + flg)
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '  ' 3RD SUB-LOOP: Gradually integrate NR into RP so that RP remains sorted in descending order                                                '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Starting from the upper bound of RP, i.e., nodes with a minimum cost
            For i = UBound(RP) To 0 Step -1
                ' If the Cumulative Cost of the Newly Reached point is lower than that of RP
                If Nodes_CC(NR(k).x, NR(k).y) <= Nodes_CC(RP(m).x, RP(m).y) Or m = 0 Then
                    ' Place the Newly Reached node into RP
                    RP(i) = NR(k)
                    If k = 0 Then Exit For
                    k = k - 1
                Else
                    ' Otherwise, keep the one from RP
                    RP(i) = RP(m)
                    m = m - 1
                End If
            Next i
        End If ' If flg > 0 Then
        
        
        m = UBound(RP): k = 0
        If m > 0 Then
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' 4TH SUB-LOOP: establishing the list of points with the Same Minimum Cumulative Cost (smCC)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Do While Nodes_CC(RP(m - 1).x, RP(m - 1).y) = Nodes_CC(RP(m).x, RP(m).y)
                ReDim Preserve smCC(k)
                smCC(k) = RP(m)
                m = m - 1: k = k + 1
                If m = 0 Then Exit Do
            Loop
            ReDim Preserve smCC(k)
            smCC(k) = RP(m)
        Else
            ReDim smCC(0)
            smCC(0) = RP(0)
        End If
        ' Remove the smCC points from RP
        m = UBound(RP): k = UBound(smCC)
        If m = k Then ReDim Preserve RP(m - k) Else ReDim Preserve RP(m - k - 1)
        
        ' Updating progress in the status bar '//
        If Timer - TimerDoEvents > 5 Then   '//
            TimerDoEvents = Timer   '//
            DoEvents   '//
            Application.StatusBar = "Calculating the shortest path... " & Int(100 * (cpt / (NbX * NbY))) & "%"   '//
        End If   '//
    
    Loop ' End of the main loop
     
Calcul:
    ' the neighbor's Cumulative Cost is equal to the current node's CC plus the neighbor's Individual Cost
    Nodes_CC(Xn, Yn) = Nodes_CC(cn.x, cn.y) + Nodes_IC(Xn, Yn)
    Nodes_edv(Xn, Yn) = edv ' Establishing the parent link, storing the elementary displacement vector
    ' The reached node is added to NR (Newly Reached)
    ReDim Preserve NR(flg)
    NR(flg).x = Xn
    NR(flg).y = Yn
    flg = flg + 1
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
        Xn = Xn - ((edv Mod 3) - 1)     ' = Xn - [-1, 0 or 1]
        Yn = Yn - (Int(edv / 3) - 1)    ' = Yn - [-1, 0 or 1]
    Loop
  
    fPath = Path
    PathfindingTime = Timer - TimerTime   '//
    Exit Function ' ******* EXIT WITH A SOLUTION *******

WorkSheetReading:
    ReDim Nodes_IC(NbX + 1, NbY + 1) ' Here is the missing line that led us here
    
    TimerDoEvents = Timer   '//
    Application.StatusBar = "Reading the Worksheet... 0%" '//
    
    ' Iterating through the entire grid
    For Xn = 1 To NbX        ' Column by column
        For Yn = 1 To NbY    ' Row by row
            Nodes_IC(Xn, Yn) = WsGrid.Cells(Yn + Oy, Xn + Ox)
            ' CC = -1: Node is not an obstacle yet and will be added to NR, then RP
            If Nodes_IC(Xn, Yn) >= 0 Then Nodes_CC(Xn, Yn) = -1
        Next Yn
        
        ' Updating progress in the status bar '//
        If Timer - TimerDoEvents > 5 Then   '//
            TimerDoEvents = Timer   '//
            DoEvents   '//
            Application.StatusBar = "Reading the Worksheet... " & Int(100 * (Xn / NbX)) & "%"   '//
        End If   '//
        
    Next Xn
    
    ReadTime = Timer - TimerTime   '//
    ' Entering identical start and end addresses triggers reading the sheet without performing pathfinding
    If EndAddress = StartAddress Then
        PathfindingTime = -1  '//
        fPath = Array(4) ' = Null vector [0, 0]
        Exit Function
    End If
    
    
    TimerTime = Timer   '//
    GoTo PathFinding

End Function



Sub Demo()
    Dim NbX%, NbY%
    Dim StartAddress$, EndAddress$
    Dim PropagationMode As Byte
    Dim Path As Variant
    Dim Vector As Byte
    Dim x%, y%
    Dim i&

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
    WsGrid.Range(StartAddress & ":" & EndAddress).Formula = "=RAND()"
    ' Since the grid has been modified, forcing worksheet re-reading without performing pathfinding by setting identical start and arrival addresses
    Path = fPath(NbX, NbY, StartAddress, StartAddress, 0)
    
    ' Calling the pathfinding procedure and storing the result in Path
    Path = fPath(NbX, NbY, StartAddress, EndAddress, PropagationMode)
    ' Case of grids with no solution
    If IsEmpty(Path) Then
        MsgBox "Sans Solution"
    Else
        ' Iterating through Path from the last element (start point vector) to the first (last vector before the end point)
        For i = UBound(Path) To 0 Step -1
            Vector = Path(i)               ' Vector extraction
            x = x + (Vector Mod 3) - 1     ' Extracting and shifting the X component
            y = y + Int(Vector / 3) - 1    ' Extracting and shifting the Y component
            Range(StartAddress).Offset(y, x).Interior.Color = vbGreen
        Next i
        ' Recoloring in cyan the end cell that was previously highlighted in green as part of the path
        Range(EndAddress).Interior.Color = vbCyan
    End If
End Sub
