Attribute VB_Name = "BeamSchedule"
Public Sub BeamSchedule()
    'Sub to prepare beam schedule
    '============================================
    'Variable declaration
    Dim iStartRow As Integer    'where the beam details start
    Dim iStartColumn As Integer 'Where the beam details start
    Dim iSearchRow As Integer
    Dim iSearchCol As Integer   'Row and column reference of where data is to be searched for
    Dim iScheduleStartRow As Integer
    Dim iScheduleStartCol As Integer
    
    Dim bGotReinforcement(1 To 3) As Boolean
    Dim bContinueSearch As Boolean 'Should we continue seraching for data?
    Dim iLocation As Integer    'variable for looping along different sections along the beam
    Dim iBeamOfInterest As Integer 'Looping variable to loop over a continuous beam
    Dim iCount As Integer   'Count of beams dealth with
    'Beam related
    '--------------------------------------------
    Dim ContinuousBeam As Collection
    Dim Beam As RCBeam
    Dim sName As String
    Dim sDwgName As String
    Dim sSectionName As String
    Dim iBeamType As Integer 'The beam type --> 1 --- 2----2----.....---2----3
    
    Dim iWidth As Integer 'Width of RC beam (mm)
    Dim iDepth As Integer 'Depth of RC beam (mm)
    Dim dAreq_top(1 To 3) As Double 'Area of reinforcement required on top (mm2)
    Dim dAreq_bot(1 To 3) As Double 'Area of reinforcement required at bottom (mm2)
    
    Dim dAreq_tor(1 To 3) As Double 'Area of torsional reinforcement required
    'Reinforcement Details
    Dim iClearCover As Integer 'Clear cover to rebar
    Dim iLinkDia As Integer 'Dia of outer reinforcement links/stirrups
    Dim iClearSpacing(1 To 2) As Integer '[Bottom Top]
    Dim iPreferredBarDia As Integer 'Detailing preference
    Dim iConcreteGrade As Integer 'MPa
    Dim iSteelGrade As Integer 'MPa
    '[Basic bar dia, Number of bars in basic layers, number of basic layers, additional bar dia, number of bars in additional layer, number of layers of additional layer]
    '[1           , 2                             , 3                     , 4                 , 5                                 , 6                                   ]
    Dim iaTop_reinforcement(1 To 6) As Integer
    Dim iaBot_reinforcement(1 To 6) As Integer
    Dim iaSide_reinforcement(1 To 2) As Integer ' [Dia Spacing]
    
    Dim sBottomReinforcement As String
    Dim sTopReinforcement As String
    Dim sSideFaceReinforcement As String
    
    Dim iaReinforcement(1 To 6, 1 To 3) As Integer
    
    Dim dAtop_left As Double
    Dim dAtop_right As Double
    Dim sLeftSideReinforcement As String
    
    'Progress control
    Dim progress As Double
    Dim StartTime As Double
    Dim TimeRemaining As String
    Dim iTotalNumberOfBeams As Integer
    Dim sTrackerFile As String
    sTrackerFile = "O:\GDC-India\INBLC\Shared PPT\DECI_Usage\RC_Beam_Schedule_Maker.txt"
    
    Dim iLastRow As Integer
    Dim iLastColumn As Integer
    
    'Related to beam orientation
    Dim iOrientation(1 To 3) As Integer 'Default [1 2 3]; if the beam is wrongly oriented, [3 2 1]
    Dim iNode1 As Integer
    Dim iNode2 As Integer
    Dim dCoordinatesNode1(1 To 3) As Double
    Dim dCoordinatesNode2(1 To 3) As Double
    Dim bGotNode1 As Boolean
    Dim bGotNode2 As Boolean
    Dim dx As Double
    Dim dY As Double
    '============================================
    'Initialisation
    iStartRow = 11
    iStartColumn = 11
    'common
    iClearCover = ActiveWorkbook.Sheets("Beams Data").Cells(2, 2) 'Cell B2
    iLinkDia = ActiveWorkbook.Sheets("Beams Data").Cells(3, 2) 'Cell B3
    iClearSpacing(1) = ActiveWorkbook.Sheets("Beams Data").Cells(4, 2) 'Cell B4
    iClearSpacing(2) = ActiveWorkbook.Sheets("Beams Data").Cells(5, 2) 'Cell B5
    If ActiveWorkbook.Sheets("Beams Data").Cells(6, 2) = "Auto" Then
        iPreferredBarDia = 0
    Else
        iPreferredBarDia = ActiveWorkbook.Sheets("Beams Data").Cells(6, 2)
    End If
    
    iConcreteGrade = ActiveWorkbook.Sheets("Beams Data").Cells(7, 2) 'Cell B7
    iSteelGrade = ActiveWorkbook.Sheets("Beams Data").Cells(8, 2) 'Cell B8
    
    iCount = 0 'Count of beams scheduled
    iScheduleStartRow = 11
    iScheduleStartCol = 1
    
    'Clear Contents------------------
    iLastRow = ActiveWorkbook.Sheets("Schedule").Cells(Rows.Count, 1).End(xlUp).Row 'Last non empty row in column A
    iLastColumn = 10 'Column J for now
    
    If iLastRow > iScheduleStartRow Then
        Range(ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow, 1), ActiveWorkbook.Sheets("Schedule").Cells(iLastRow, iLastColumn)).ClearContents
    End If
    ActiveWorkbook.Sheets("Schedule").Cells(7, 2) = 0
    ActiveWorkbook.Sheets("Schedule").Cells(7, 3) = "Done"
    ActiveWorkbook.Sheets("Schedule").Cells(1, 1) = ""
    

    
    
    ActiveWorkbook.Sheets("Schedule").Cells(7, 2).Value = "Estimating time required"
    iTotalNumberOfBeams = 0
    iRow = iStartRow
    Do While ActiveWorkbook.Sheets("Beams Data").Cells(iRow, iStartColumn) <> ""
        iCol = iStartColumn
        Do While ActiveWorkbook.Sheets("Beams Data").Cells(iRow, iCol) <> ""
            iTotalNumberOfBeams = iTotalNumberOfBeams + 1
            iCol = iCol + 1
        Loop
        iRow = iRow + 1
    Loop
    Application.ScreenUpdating = False
    '============================================
    iRow = iStartRow
    StartTime = Timer 'Timer starts
    Do While ActiveWorkbook.Sheets("Beams Data").Cells(iRow, iStartColumn) <> ""
        iCol = iStartColumn
        Do While ActiveWorkbook.Sheets("Beams Data").Cells(iRow, iCol) <> ""
            
            'Get beam data
            sName = ActiveWorkbook.Sheets("Beams Data").Cells(iRow, iCol).Value
            
            
            
            If sName = "B244" Then
                Debug.Print "Watch"
            End If
            'get name in drawing
            iSearchRow = 11
            iSearchCol = 1
            bContinueSearch = True
            Do While bContinueSearch
                If ActiveWorkbook.Sheets("DwgNames").Cells(iSearchRow, iSearchCol) = sName Then 'If you come across the same beam name in the list
                    sDwgName = ActiveWorkbook.Sheets("DwgNames").Cells(iSearchRow, iSearchCol + 1)
                    bContinueSearch = False
                ElseIf ActiveWorkbook.Sheets("DwgNames").Cells(iSearchRow, iSearchCol) = "" Then
                    sDwgName = sName & "-NA-"
                    bContinueSearch = False
                End If
                iSearchRow = iSearchRow + 1
            Loop
            'Get Beam dimensions
            'a) Get Section Name
            iSearchRow = 4
            iSearchCol = 37 'Column AK
            bContinueSearch = True
            Do While bContinueSearch
                If ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = sName Then 'If you come across the same beam name in the list
                    sSectionName = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 6) 'AK+6 = AQ column --> Design section
                    bContinueSearch = False
                ElseIf ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = "" Then
                    sSectionName = InputBox("Property missing for " & sName & "(" & sDwgName & ")" & Chr(13) & "Enter Section name:", "Missing Section Name", "Section Name")
                    bContinueSearch = False
                End If
                iSearchRow = iSearchRow + 1
            Loop
            'b) Get Width and Depth
            iSearchRow = 4
            iSearchCol = 11 'Column K
            bContinueSearch = True
            Do While bContinueSearch
                If ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = sSectionName Then 'If you come across the same beam name in the list
                    iWidth = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 4) 'K+4 = O column --> t2 = Width
                    iDepth = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 3) 'K+3 = N column --> t3 = Depth
                    bContinueSearch = False
                ElseIf ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = "" Then
                    iWidth = InputBox("Dimensions missing for " & sSectionName & Chr(13) & "Enter Width:", "Missing Section Properties", 400)
                    iDepth = InputBox("Dimensions missing for " & sSectionName & Chr(13) & "Enter Depth:", "Missing Section Properties", 400)
                    'Log it into the sheet just in case
                    ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = sSectionName
                    ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 4) = iDepth
                    ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 6) = iWidth
                    ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 7) = "Added by the sheet"
                    bContinueSearch = False
                End If
                iSearchRow = iSearchRow + 1
            Loop
            
            'Beam orientation Checking
            iSearchRow = 4
            iSearchCol = 1 'Column A
            bContinueSearch = True
            Do While bContinueSearch
                If ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = sName Then 'If you come across the same beam name in the list
                    iNode1 = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 1) 'A+1 = B column --> I-End Point
                    iNode2 = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 2) 'A+2 = C column --> J-End Point
                    bContinueSearch = False
                ElseIf ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = "" Then
                    iNode1 = InputBox("Node information missing for " & sName & Chr(13) & "Enter I-End Point:", "Missing Node connectivity", "Node number")
                    iNode2 = InputBox("Node information missing for " & sName & Chr(13) & "Enter J-End Point:", "Missing Node connectivity", "Node number")
                    bContinueSearch = False
                End If
                iSearchRow = iSearchRow + 1
            Loop 'At the end of this loop, we know the node numbers
            iSearchRow = 4
            iSearchCol = 6 'Column F
            bGotNode1 = False
            bGotNode2 = False
            bContinueSearch = True
            Do While bContinueSearch
                If ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = iNode1 Or ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = iNode2 Then 'If you come across the same beam name in the list
                    If ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = iNode1 Then
                        dCoordinatesNode1(1) = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 1) 'F+1 = G ---> X coordinate
                        dCoordinatesNode1(2) = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 2) 'F+2 = H ---> X coordinate
                        dCoordinatesNode1(3) = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 3) 'F+3 = I ---> X coordinate
                        bGotNode1 = True
                    Else
                        dCoordinatesNode2(1) = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 1) 'F+1 = G ---> X coordinate
                        dCoordinatesNode2(2) = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 2) 'F+2 = H ---> X coordinate
                        dCoordinatesNode2(3) = ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol + 3) 'F+3 = I ---> X coordinate
                        bGotNode2 = True
                    End If
                    If bGotNode1 And bGotNode2 Then
                        bContinueSearch = False
                    End If
                ElseIf ActiveWorkbook.Sheets("ETABS_Input").Cells(iSearchRow, iSearchCol) = "" Then
                    'Not addressed yet
                    Debug.Print "Nodal coordinates are missing for " & iNode1 & " and/or " & iNode2
                End If
                iSearchRow = iSearchRow + 1
            Loop 'now we have coordinates; still need to find orientation
            dY = dCoordinatesNode2(2) - dCoordinatesNode1(2)
            dx = dCoordinatesNode2(1) - dCoordinatesNode1(1)
            If dY >= 0 And dx >= 0 Then 'First quqdrant
                'No issue proceed as normal
                iOrientation(1) = 1
                iOrientation(2) = 2
                iOrientation(3) = 3
            ElseIf dY >= 0 And dx < 0 Then 'dY>0 and dX<0 ==> second quadrant
                'If slope is very steep ---> nearly vertical => do not worry
                If Abs(dY / dx) > 0.5 Then
                    iOrientation(1) = 1
                    iOrientation(2) = 2
                    iOrientation(3) = 3
                Else 'swap orientation for scheduling purpose
                    iOrientation(1) = 3
                    iOrientation(2) = 2
                    iOrientation(3) = 1
                End If
            ElseIf dY < 0 And dx < 0 Then 'dY>0 and dX<0 ==> third quadrant
                'Swap
                iOrientation(1) = 3
                iOrientation(2) = 2
                iOrientation(3) = 1
            Else 'Fourth quadrant
                'If slope is very steep ---> nearly vertical => do not worry
                If Abs(dY / dx) > 0.5 Then
                    iOrientation(1) = 1
                    iOrientation(2) = 2
                    iOrientation(3) = 3
                Else 'swap orientation for scheduling purpose
                    iOrientation(1) = 3
                    iOrientation(2) = 2
                    iOrientation(3) = 1
                End If
            End If
            
            'Reinforcement Estimate
            '=======================
            'Bottom and top reinforcement
            iSearchRow = 4
            iSearchCol = 1 'Column A
            
            bGotReinforcement(1) = False
            bGotReinforcement(2) = False
            bGotReinforcement(3) = False
            If bGotReinforcement(1) And bGotReinforcement(2) And bGotReinforcement(3) Then
                bContinueSearch = False
            Else
                bContinueSearch = True
            End If
            Do While bContinueSearch
                If sName = "B244" Then
                    Debug.Print "Row =" & iSearchRow
                End If
                If iSearchRow = 109 Then
                    Debug.Print "About to find reinforcement areas"
                End If
                If ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol) = sName Then 'If you come across the same beam name in the list
                    If ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 3) = "End-I" Then 'A+3 = D column --> Section location along beam
                        iLocation = iOrientation(1)
                        dAreq_top(iLocation) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 6) 'A+6 = G --> Top reinforcement demand
                        dAreq_bot(iLocation) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 9) 'A+9 = J --> Bottom reinforcement demand
                        bGotReinforcement(iLocation) = True
                    ElseIf ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 3) = "Middle" Then 'A+3 = D column --> Section location along beam
                        iLocation = iOrientation(2)
                        dAreq_top(iLocation) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 6) 'A+6 = G --> Top reinforcement demand
                        dAreq_bot(iLocation) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 9) 'A+9 = J --> Bottom reinforcement demand/
                        bGotReinforcement(iLocation) = True
                    Else
                        iLocation = iOrientation(3)
                        dAreq_top(iLocation) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 6) 'A+6 = G --> Top reinforcement demand
                        dAreq_bot(iLocation) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 9) 'A+9 = J --> Bottom reinforcement demand
                        bGotReinforcement(iLocation) = True
                    End If
                    If bGotReinforcement(1) And bGotReinforcement(2) And bGotReinforcement(3) Then
                        bContinueSearch = False
                    Else
                        bContinueSearch = True
                    End If
                ElseIf ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol) = "" Then
                    dAreq_top(1) = 0 'Zero as there is no record
                    dAreq_bot(1) = 0
                    dAreq_top(2) = 0 'Zero as there is no record
                    dAreq_bot(2) = 0
                    dAreq_top(3) = 0 'Zero as there is no record
                    dAreq_bot(3) = 0
                    bContinueSearch = False
                End If
                iSearchRow = iSearchRow + 1
            Loop
            'Torsional reinforcement
            iSearchRow = 4
            iSearchCol = 12 'Column L
            bGotReinforcement(1) = False
            bGotReinforcement(2) = False
            bGotReinforcement(3) = False
            If bGotReinforcement(1) And bGotReinforcement(2) And bGotReinforcement(3) Then
                bContinueSearch = False
            Else
                bContinueSearch = True
            End If
            Do While bContinueSearch
                If ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol) = sName Then 'If you come across the same beam name in the list
                    If ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 3) = "End-I" Then 'L+3 = O column --> Section location along beam
                        iLocation = iOrientation(1)
                        dAreq_tor(iLocation) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 12) 'L+12 = X --> Torsional reinforcement demand
                        bGotReinforcement(iLocation) = True
                    ElseIf ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 3) = "Middle" Then 'L+3 = O column --> Section location along beam
                        iLocation = iOrientation(2)
                        dAreq_tor(2) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 12) 'L+12 = X --> Torsional reinforcement demand
                        bGotReinforcement(iLocation) = True
                    Else
                        iLocation = iOrientation(3)
                        dAreq_tor(3) = ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol + 12) 'L+12 = X --> Torsional reinforcement demand
                        bGotReinforcement(iLocation) = True
                    End If
                    If bGotReinforcement(1) And bGotReinforcement(2) And bGotReinforcement(3) Then
                        bContinueSearch = False
                    Else
                        bContinueSearch = True
                    End If
                ElseIf ActiveWorkbook.Sheets("ETABS_Output").Cells(iSearchRow, iSearchCol) = "" Then
                    dAreq_tor(1) = 0 'Zero as there is no record
                    dAreq_tor(2) = 0
                    dAreq_tor(3) = 0
                    bContinueSearch = False
                End If
                iSearchRow = iSearchRow + 1
            Loop
            'Find beam type
            If ActiveWorkbook.Sheets("Beams Data").Cells(iRow, iCol - 1) = "" Then 'Nothing to left
                iBeamType = 1
            ElseIf ActiveWorkbook.Sheets("Beams Data").Cells(iRow, iCol + 1) = "" Then 'Nothing to follow
                iBeamType = 3
            Else
                iBeamType = 2
            End If
            'Create and solve for beams
            '===================================
            Set Beam = New RCBeam
            Beam.Name = sName
            Beam.NameInDrawing = sDwgName
            Beam.Length = VBA.Sqr(dx ^ 2 + dY ^ 2)
            Beam.Cover = iClearCover
            Beam.LinkDia = iLinkDia
            Beam.Width = iWidth
            Beam.Depth = iDepth
            Beam.BeamType = iBeamType
            Beam.MinClearSpacingBottom = iClearSpacing(1)
            Beam.MinClearSpacingTop = iClearSpacing(2)
            Beam.PreferredBarDia(1) = iPreferredBarDia
            Beam.PreferredBarDia(2) = iPreferredBarDia
            Beam.ConcreteGrade = iConcreteGrade
            Beam.SteelGrade = iSteelGrade
            For iLocation = 1 To 3 'For each location along the beam
                Beam.AreaRequiredBottom(iLocation) = dAreq_bot(iLocation)
                Beam.AreaRequiredTop(iLocation) = dAreq_top(iLocation)
                Beam.AreaRequiredTorsion(iLocation) = dAreq_tor(iLocation)
            Next iLocation
            '============================================
            '========= Preparing the schedule ===========
            '============================================
            ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol) = Beam.Name
            ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 1) = Beam.NameInDrawing
            If Beam.BeamType <> 3 Then
                ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 2) = Beam.BeamType
            Else
                ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 2) = 2 'My 3 is 2b in standard as 3 is cantilever
            End If
            ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 3) = Beam.Width
            ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 4) = Beam.Depth
            For iLocation = 1 To 3
                'Bottom
                For iRebarComponent = 1 To 6
                    iaBot_reinforcement(iRebarComponent) = WorksheetFunction.Index(Beam.GetReinforcement(iLocation), 1, iRebarComponent) '1 For bottom
                Next iRebarComponent
                sBottomReinforcement = WorksheetFunction.Index(Beam.GetReinforcementDescriptionAsString(), iLocation, 1) '1 for bottom face
                
                'Top
                For iRebarComponent = 1 To 6
                    iaTop_reinforcement(iRebarComponent) = WorksheetFunction.Index(Beam.GetReinforcement(iLocation), 2, iRebarComponent) '2 For top
                Next iRebarComponent
                sTopReinforcement = WorksheetFunction.Index(Beam.GetReinforcementDescriptionAsString(), iLocation, 2) '2 for top face
                
                'Side - not relevant now
                'Scheduling
                
                
                
                
                If iLocation = 1 And Beam.BeamType = 1 Then 'Beginning of continous beam
                    ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 5) = sTopReinforcement 'A
                ElseIf iLocation = 3 And Beam.BeamType = 3 Then 'End of the continuous beam
                    ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 5) = sTopReinforcement 'A
                ElseIf iLocation = 2 Then
                    'Maximum of LeftBot, MidBot and RightBot is to be provided as A bar --- Default: iLocation=2 is the critical for bottom steel
                    If Beam.ReinforcementAreaProvided(3, 1) = WorksheetFunction.Max(Beam.ReinforcementAreaProvided(1, 1), Beam.ReinforcementAreaProvided(2, 1), Beam.ReinforcementAreaProvided(3, 1)) Then
                        'End-J needs more area; so provide that as bottom steel
                        sBottomReinforcement = WorksheetFunction.Index(Beam.GetReinforcementDescriptionAsString(), 3, 1) '1 for bottom face and 3 for End-J
                    ElseIf Beam.ReinforcementAreaProvided(1, 1) = WorksheetFunction.Max(Beam.ReinforcementAreaProvided(1, 1), Beam.ReinforcementAreaProvided(2, 1), Beam.ReinforcementAreaProvided(3, 1)) Then
                        'End-I needs more area; so provide that as bottom steel
                        sBottomReinforcement = WorksheetFunction.Index(Beam.GetReinforcementDescriptionAsString(), 1, 1) '1 for bottom face
                    End If
                    ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 6) = sBottomReinforcement 'B
                    ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 7) = sTopReinforcement 'C
                ElseIf iLocation = 3 Then ' Left end of internal support []=====||=====
                    dAtop_left = Beam.ReinforcementAreaProvided(iLocation, 2) 'iLocation = 3 - the right end and 2 => top surface
                    sLeftSideReinforcement = sTopReinforcement
                ElseIf iLocation = 1 Then 'Right end of internal support
                    If dAtop_left > Beam.ReinforcementAreaProvided(iLocation, 2) Then 'iLocation = 1 - the left end and 2 => top surface
                        ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 8) = sLeftSideReinforcement 'D
                    Else
                        ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 8) = sTopReinforcement 'D
                    End If
                End If
                'Side face
                ActiveWorkbook.Sheets("Schedule").Cells(iScheduleStartRow + iCount, iScheduleStartCol + 9) = WorksheetFunction.Index(Beam.GetReinforcementDescriptionAsString(), 1, 3) ' 1=>At End-I (assuming uniform section); 3=> Side face
            Next iLocation
            progress = 1# * iCount / iTotalNumberOfBeams
            'Update the progress to user
            If progress > 0 Then
                TimeRemaining = Format(((Timer - StartTime) / progress) / 8400, "hh:mm:ss")
                Application.ScreenUpdating = True
                ActiveWorkbook.Sheets("Schedule").Cells(7, 2).Value = progress
                ActiveWorkbook.Sheets("Schedule").Cells(7, 3).Value = TimeRemaining & "remaining"
                Application.ScreenUpdating = False
            End If
            iCol = iCol + 1
            iCount = iCount + 1
            'End of one column loop
        Loop
        iRow = iRow + 1
    Loop
    sFileName = ThisWorkbook.FullName
    ActiveWorkbook.Sheets("Schedule").Cells(1, 1).Value = Application.UserName & " scheduled " & iCount & " beams using " & ActiveWorkbook.Sheets("Read Me").Cells(1, 1) & " on " & Format(Now(), "dd/MM/yyyy  h:mm:ss") & " from " & sFileName
    Open sTrackerFile For Append As #1
    Write #1, Application.UserName & " scheduled" & Chr(9) & iCount & Chr(9) & "beams using " & ActiveWorkbook.Sheets("Read Me").Cells(1, 1) & " on " & Format(Now(), "dd/MM/yyyy  h:mm:ss") & " from " & sFileName
    Close #1
    ActiveWorkbook.Sheets("Schedule").Cells(7, 2) = 1
    ActiveWorkbook.Sheets("Schedule").Cells(7, 3) = "Done"
End Sub
Private Function GetReinforcementAsString(ByVal Reinforcement As Variant) As String
    Dim OptimumRebar As String
    If Reinforcement(1) = Reinforcement(4) And Reinforcement(2) = Reinforcement(5) Then 'if all layers are of same number and dia
        OptimumRebar = (Reinforcement(3) + Reinforcement(6)) & "X" & Reinforcement(2) & "-T" & Reinforcement(1)
    Else
        If Reinforcement(3) > 1 Then 'Avoid saying 1 layer which is default
            OptimumRebar = Reinforcement(3) & "X" & Reinforcement(2) & "-T" & Reinforcement(1)
        Else
            OptimumRebar = Reinforcement(2) & "-T" & Reinforcement(1)
        End If
        
        If Reinforcement(6) > 0 Then 'Only if there is an additional layer
            If Reinforcement(6) > 1 Then 'Avoid saying 1 layer which is default
                OptimumRebar = OptimumRebar & " + " & Reinforcement(6) & "X" & Reinforcement(5) & "-T" & Reinforcement(4)
            Else
                OptimumRebar = OptimumRebar & " + " & Reinforcement(5) & "-T" & Reinforcement(4)
            End If
        End If
    End If
    GetReinforcementAsString = OptimumRebar
End Function

Private Function GetReinforcementArea(ByVal Reinforcement As Variant)
    
    GetReinforcementArea = Reinforcement(3) * Reinforcement(2) * WorksheetFunction.Pi() / 4 * Reinforcement(1) ^ 2
    GetReinforcementArea = Reinforcement(6) * Reinforcement(5) * WorksheetFunction.Pi() / 4 * Reinforcement(4) ^ 2
End Function
