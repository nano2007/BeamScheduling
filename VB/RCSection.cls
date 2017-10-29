VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RCSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class of Reinforced Concrete beams
Private psDwgName As String 'Name as referenced in drawings
Private psName As String 'Name as in ETABS
Private piWidth As Integer 'Width of RC beam (mm)
Private piDepth As Integer 'Depth of RC beam (mm)
Private pdAreq_top As Double 'Area of reinforcement required on top (mm2)
Private pdAreq_bot As Double 'Area of reinforcement required at bottom (mm2)
Private pdAreq_tor As Double 'Area of torsional reinforcement required
'Reinforcement Details
Private piClearCover As Integer 'Clear cover to rebar
Private piLinkDia As Integer 'Dia of outer reinforcement links/stirrups
'[Basic bar dia, Number of bars in basic layers, number of basic layers, additional bar dia, number of bars in additional layer, number of layers of additional layer]
'[1           , 2                             , 3                     , 4                 , 5                                 , 6                                   ]
Private piaTop_reinforcement(1 To 6) As Integer
Private piaBot_reinforcement(1 To 6) As Integer
Private piaSide_reinforcement(1 To 2) As Integer ' [Dia Spacing]
Private pdAsideface As Double
Private pbReinforcementProvided As Boolean 'True if reinforcement has been estimated
Private pdAreaProvided(1 To 2) As Double ' [Botom area; Top area]
Private psReinforcementDescription(1 To 3) As String  'bottom description, top description & side face description
Private piMinClearSpacingBottom As Integer
Private piMinClearSpacingTop As Integer
Private piPreferredDia(1 To 2)  As Integer 'Preferred dia of main rebar {Bottom Top}
Private piConcreteGrade As Integer 'MPa
Private piSteelGrade As Integer 'MPa Aa
'Shear
Private piaShearLinks(1 To 3, 1 To 2) As Integer  ' {{Leg_outer, Dia_outer, Spacing_outer};{Leg_internal, Dia_internal, Spacing_internal}}
Private pdLinkArea_tor As Double 'Area of shear links (outer) required due to torsion
Private pdLinkArea_shear As Double 'Area of sheare links required due to normal shear alone


'========================
Private Sub Class_Initialize()
    piClearCover = 40
    piLinkDia = 10
    piMinClearSpacingBottom = 50
    piMinClearSpacingTop = 75
    pbReinforcementProvided = False
    piPreferredDia(1) = 0
    piPreferredDia(2) = 0
End Sub

'======= Name Properties =====
Public Property Get Name() As String
    Name = psName
End Property
Public Property Let Name(ByVal ETABS_Name As String)
    psName = ETABS_Name
End Property

Public Property Get NameInDrawing() As String
    NameInDrawing = psDwgName
End Property
Public Property Let NameInDrawing(ByVal BeamNameAsInDrawing As String)
    psDwgName = BeamNameAsInDrawing
End Property
'Sectional Properties
Public Property Get Width() As Integer
    Width = piWidth
End Property
Public Property Let Width(ByVal iBeamWidth As Integer)
    piWidth = iBeamWidth
End Property
Public Property Get Depth() As Integer
    Depth = piDepth
End Property
Public Property Let Depth(ByVal iDepthOfBeam As Integer)
    piDepth = iDepthOfBeam
End Property

Public Property Get Cover() As Integer
    Cover = piClearCover
End Property
Public Property Let Cover(ByVal iClearCover As Integer)
    piClearCover = iClearCover
End Property
Public Property Get LinkDia() As Integer
    LinkDia = piLinkDia
End Property
Public Property Let LinkDia(ByVal iLinkDia As Integer)
    piLinkDia = iLinkDia
End Property

'Reinforcemetn demands
Public Property Get AreaRequiredTop() As Double
    AreaRequiredTop = pdAreq_top
End Property
Public Property Let AreaRequiredTop(ByVal dTopReinforcementAreaRequired As Double)
    pdAreq_top = dTopReinforcementAreaRequired
End Property
Public Property Get AreaRequiredBottom() As Double
    AreaRequiredBottom = pdAreq_bot
End Property
Public Property Let AreaRequiredBottom(ByVal dBottomReinforcementAreaRequired As Double)
    pdAreq_bot = dBottomReinforcementAreaRequired
End Property
Public Property Get AreaRequiredTorsion() As Double
    AreaRequiredTop = pdAreq_tor
End Property
Public Property Let AreaRequiredTorsion(ByVal dTorReinforcementAreaRequired As Double)
    pdAreq_tor = dTorReinforcementAreaRequired
End Property
Public Property Get MinClearSpacingBottom() As Integer
    MinClearSpacingBottom = piMinClearSpacingBottom
End Property
Public Property Let MinClearSpacingBottom(ByVal iMinClearSpacingBottom As Integer)
    piMinClearSpacingBottom = iMinClearSpacingBottom
End Property
Public Property Get MinClearSpacingTop() As Integer
    MinClearSpacingTop = piMinClearSpacingTop
End Property
Public Property Let MinClearSpacingTop(ByVal iMinClearSpacingTop As Integer)
    piMinClearSpacingTop = iMinClearSpacingTop
End Property
Public Property Get PreferredBarDia(ByVal iFace As Integer) As Integer 'Zero for automatic
    PreferredBarDia = piPreferredDia(iFace)
End Property
Public Property Let PreferredBarDia(ByVal iFace As Integer, ByVal iPreferredBarDia As Integer)
    piPreferredDia(iFace) = iPreferredBarDia
End Property
Public Property Get ConcreteGrade() As Integer
    ConcreteGrade = piConcreteGrade
End Property
Public Property Let ConcreteGrade(ByVal iConcreteGrade As Integer)
    piConcreteGrade = iConcreteGrade
End Property
Public Property Get SteelGrade() As Integer
    SteelGrade = piSteelGrade
End Property
Public Property Let SteelGrade(ByVal iSteelGrade As Integer)
    piSteelGrade = iSteelGrade
End Property
Public Property Get StirrupAreaForTorsion() As Double
    StirrupAreaForTorsion = pdLinkArea_tor
End Property
Public Property Let StirrupAreaForTorsion(ByVal dLinkAreaRequiredDueToTorsion As Double)
    pdLinkArea_tor = dLinkAreaRequiredDueToTorsion
End Property
Public Property Get StirrupAreaForTShear() As Double
    StirrupAreaForTorsion = pdLinkArea_shear
End Property
Public Property Let StirrupAreaForShear(ByVal dLinkAreaRequiredDueToShear As Double)
    pdLinkArea_shear = dLinkAreaRequiredDueToShear
End Property
'==================================
'====== Other methods =============
'==================================
Private Sub EstimateSideFaceReinforcement()
    Dim dApproximateRebarDepth As Double
    Dim iNumberOfBars As Integer
    dApproximateRebarDepth = 0.2 * piDepth 'Conservative?
    'Side face is mandatory only if Depth is more than 900 mm as per ACI318M08, cl. 10.6.7
    If piDepth < 750 Then 'Amended to 750 mm based on the Middle-East detailing standard
        piaSide_reinforcement(1) = 0 'Using 12 mm as the side face bar dia
        piaSide_reinforcement(2) = 0 'Spacing is fixed as 150 mm
        iNumberOfBars = 0
        'Do you want to provide side face for shall beams?
    Else
        piaSide_reinforcement(1) = 12 'Using 12 mm as the side face bar dia
        piaSide_reinforcement(2) = 150 'Spacing is fixed as 150 mm
        iNumberOfBars = WorksheetFunction.Floor_Math((Me.Depth - 2 * dApproximateRebarDepth) / piaSide_reinforcement(2), 1) - 1 'Conservative (N-1 bars due to spacing
    End If
    
    pdAsideface = 2 * iNumberOfBars * WorksheetFunction.Pi() / 4 * piaSide_reinforcement(1) ^ 2
End Sub

Private Sub EstimateBottomReinforcement(Optional ByVal iNumberOfBarsToBeForced As Integer = 0)
    'If Side face reinforcement > Torsional area required; Only bottom steel required. Otherwise add half of Torsional steel
    Dim dAreq As Double
    Dim iPreferredDia As Double
    Dim iMinSpacing As Double
    Dim bReturnAsText As Boolean
    Dim dAmin As Double 'Minimum area of reinforcement
        
    bReturnAsText = False
    iMinSpacing = piMinClearSpacingBottom
    iPreferredDia = piPreferredDia(1) 'Suitably amend if there is a preference. Leaving without bias for this phase of development
    If pdAsideface < pdAreq_tor Then
        dAreq = pdAreq_bot + (pdAreq_tor - pdAsideface) / 2
    Else
        dAreq = pdAreq_bot
    End If
    'Check for minimum steel ---> conservative approximation of clause 10.5.1 of ACI 318M-08
    dAmin = 0.25 * VBA.Sqr(piConcreteGrade) * piWidth * (piDepth - piClearCover - piLinkDia) / piSteelGrade
    dAmin = WorksheetFunction.Max(dAmin, 1.4 * piWidth * (piDepth - piClearCover - piLinkDia) / piSteelGrade)
    If dAmin > dAreq Then
        dAreq = dAmin
    End If
    Call ProvideOptimumRebar(1, dAreq, iNumberOfBarsToBeForced) ' 1=> iFace = 1 => Bottom face
    'Update area provided
    pdAreaProvided(1) = piaBot_reinforcement(3) * piaBot_reinforcement(2) * WorksheetFunction.Pi() / 4 * piaBot_reinforcement(1) ^ 2
    pdAreaProvided(1) = pdAreaProvided(1) + piaBot_reinforcement(6) * piaBot_reinforcement(5) * WorksheetFunction.Pi() / 4 * piaBot_reinforcement(4) ^ 2
End Sub
Private Sub EstimateTopReinforcement(Optional ByVal iNumberOfBarsToBeForced As Integer = 0)
    'If Side face reinforcement > Torsional area required; Only bottom steel required. Otherwise add half of Torsional steel
    Dim dAreq As Double
    Dim iPreferredDia As Double
    Dim iMinSpacing As Double
    Dim bReturnAsText As Boolean
    Dim dAmin As Double
    bReturnAsText = False

    iMinSpacing = piMinClearSpacingTop
    iPreferredDia = piPreferredDia(2) 'Suitably amend if there is a preference. Leaving without bias for this phase of development
    If pdAsideface < pdAreq_tor Then
        dAreq = pdAreq_top + (pdAreq_tor - pdAsideface) / 2
    Else
        dAreq = pdAreq_top
    End If
    'Check for minimum steel ---> conservative approximation of clause 10.5.1 of ACI 318M-08
    dAmin = 0.25 * VBA.Sqr(piConcreteGrade) * piWidth * (piDepth - piClearCover - piLinkDia) / piSteelGrade
    dAmin = WorksheetFunction.Max(dAmin, 1.4 * piWidth * (piDepth - piClearCover - piLinkDia) / piSteelGrade)
    If dAmin > dAreq Then
        dAreq = dAmin
    End If
    'If psName = "B32 Section 1" Or psName = "B26 Section 3" Then
        'Debug.Print "Top area for " & psName & " = " & dAreq
    'End If
    Call ProvideOptimumRebar(2, dAreq, iNumberOfBarsToBeForced) ' 2=> iFace = 2 => Top face
    
    pdAreaProvided(2) = piaTop_reinforcement(3) * piaTop_reinforcement(2) * WorksheetFunction.Pi() / 4 * piaTop_reinforcement(1) ^ 2
    pdAreaProvided(2) = pdAreaProvided(2) + piaTop_reinforcement(6) * piaTop_reinforcement(5) * WorksheetFunction.Pi() / 4 * piaTop_reinforcement(4) ^ 2
End Sub

Private Sub EstimateReinforcement()
    'Estimate all reinforcement
    Call EstimateSideFaceReinforcement
    Call EstimateBottomReinforcement

    Call EstimateTopReinforcement

    pbReinforcementProvided = True
End Sub

'Read only property
Public Function GetAreaOfSteelProvided() '[Abot ; Atop]
    GetAreaOfSteelProvided = Array(pdAreaProvided(1), pdAreaProvided(2))
End Function
Public Function GetReinforcementAsString() '[Bottom; Top; Side]
    Dim sBottomReinforcement As String
    Dim sTopReinforcement As String
    Dim sSideFaceReinforcement As String
    If Not pbReinforcementProvided Then
        Call EstimateReinforcement
    End If
    Call GetReinforcementDescription
    sBottomReinforcement = psReinforcementDescription(1)
    sTopReinforcement = psReinforcementDescription(2)
    sSideFaceReinforcement = psReinforcementDescription(3)
    GetReinforcementAsString = Array(sBottomReinforcement, sTopReinforcement, sSideFaceReinforcement)
End Function
Public Function GetReinforcement() As Variant
    'Function to return the reinforcement schedule
    If Not pbReinforcementProvided Then
        Call EstimateReinforcement
    End If
    Dim iaReinforcement(1 To 6, 1 To 3) As Integer
    'Bottom reinforcement
    iaReinforcement(1, 1) = piaBot_reinforcement(1) 'Bottom base dia
    iaReinforcement(2, 1) = piaBot_reinforcement(2) 'Bottom base number of rebar per layer
    iaReinforcement(3, 1) = piaBot_reinforcement(3) 'Bottom base number of layers
    iaReinforcement(4, 1) = piaBot_reinforcement(4) 'Bottom additional dia
    iaReinforcement(5, 1) = piaBot_reinforcement(5) 'Bottom additional number of rebar per layer
    iaReinforcement(6, 1) = piaBot_reinforcement(6) 'Bottom additional number of layers
    'Top Reinforcement
    iaReinforcement(1, 2) = piaTop_reinforcement(1) 'Top base dia
    iaReinforcement(2, 2) = piaTop_reinforcement(2) 'Top base number of rebar per layer
    iaReinforcement(3, 2) = piaTop_reinforcement(3) 'Top base number of layers
    iaReinforcement(4, 2) = piaTop_reinforcement(4) 'Top additional dia
    iaReinforcement(5, 2) = piaTop_reinforcement(5) 'Top additional number of rebar per layer
    iaReinforcement(6, 2) = piaTop_reinforcement(6) 'Top additional number of layers
    'Side face
    iaReinforcement(1, 3) = piaSide_reinforcement(1)
    iaReinforcement(2, 3) = piaSide_reinforcement(2)
    iaReinforcement(3, 3) = 0
    iaReinforcement(4, 3) = 0
    iaReinforcement(5, 3) = 0
    iaReinforcement(6, 3) = 0
    'Return this array
    GetReinforcement = iaReinforcement
End Function

Private Sub GetReinforcementDescription()
    Dim OptimumRebar As String
    Dim iRebarComponent As Integer
    Dim Reinforcement(1 To 6) As Integer
    
    For iFace = 1 To 2 '1 = Bottom and 2 = Top
        If iFace = 1 Then
            For iRebarComponent = 1 To 6
                Reinforcement(iRebarComponent) = piaBot_reinforcement(iRebarComponent)
            Next iRebarComponent
        Else
            For iRebarComponent = 1 To 6
                Reinforcement(iRebarComponent) = piaTop_reinforcement(iRebarComponent)
            Next iRebarComponent
        End If
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
        psReinforcementDescription(iFace) = OptimumRebar
    Next iFace
    If piaSide_reinforcement(1) > 0 Then
        psReinforcementDescription(3) = "T" & piaSide_reinforcement(1) & "-" & piaSide_reinforcement(2)
    Else
        psReinforcementDescription(3) = "-"
    End If
End Sub




'================================================
'========== REBAR OPTIONS FUNCTIONS============
'================================================


'ProvideOptimumRebar(iFace,dAreaRequired,[iForceNumberInLayer])
Private Sub ProvideOptimumRebar(ByVal iFace As Integer, ByVal dAreaRequired As Double, Optional ByVal iForceNumberInLayer As Integer = 0)
    'Underdevelopment;
    'Prepared by Harikrishnan Panikkaveettil

    Dim iNumberOfBars As Integer
    Dim iNumberOfLayers As Integer
    Dim iOption As Integer 'Looping variable for size
    Dim iDia As Integer
    Dim iAreaOfSingleBar As Double ' Area of a single rebar of size iDia
    Dim iTotalNumberOfBarsRequired As Integer
    Dim dAreaProvided As Double
    Dim iPreferredOption As Integer 'Preferred Solution
    Dim bOddNumberAcceptable As Boolean
    
    'Rebar Dia                : 10 12 16 20 25 32
    'Number of bars per layer : 00 00 00 00 00 00
    'Number of layers         : 01 01 01 01 01 01
    'Marks                    : 00 00 00 00 00 00       Highest marks is recommended
    
    Dim iaRebarOptions_base(1 To 4, 1 To 6) As Integer 'Base solution
    Dim iaRebarOptions_add(1 To 4, 1 To 6) As Integer 'Additional solution
    Dim dAreaRequired_base As Double 'IBasic area required
    Dim dAreaRequired_add As Double 'Inner most layer if required
    Dim dResult(1 To 6) As Integer 'Result to be returned if the return is not exptected as text
    
    '======== Initialising =========
    If iFace = 1 Then
        iMinSpacing = piMinClearSpacingBottom
    ElseIf iFace = 2 Then
        iMinSpacing = piMinClearSpacingTop
    Else
        Debug.Print "OptimumRebar is asked to design for something other than Bottom/Top bars iFace = " & iFace
    End If
    
    'Initialising the Array
    dAreaRequired_base = dAreaRequired
    iaRebarOptions_base(1, 1) = 10
    iaRebarOptions_base(1, 2) = 12
    iaRebarOptions_base(1, 3) = 16
    iaRebarOptions_base(1, 4) = 20
    iaRebarOptions_base(1, 5) = 25
    iaRebarOptions_base(1, 6) = 32
    iMinNoOfBars = 2
    '======== Prioity list =========
    'Rebar Dia                : 10 12 16 20 25 32
    'Number of bars per layer : 00 00 00 00 00 00
    'Number of layers         : 01 01 01 01 01 01
    '<<<<<<<<<<< Yet to workout the strategy >>>>>>`
    
    
    '======== Preparing Options =========
    For iOption = 1 To 6
        iDia = iaRebarOptions_base(1, iOption) 'Size of rebar
        'Spacing between the centres of extreme bars are given by (piWidth - 2 * piClearCover - 2 * piLinkDia - iDia)
        'If the clear spacing is S and dia of bar is d, there will be N number of bar dias for N number of spacing
        'Thus, Number of bars = 1 + N
        
        If iForceNumberInLayer > 0 Then
            iNumberOfBars = iForceNumberInLayer
        Else
            iNumberOfBars = (piWidth - 2 * piClearCover - 2 * piLinkDia - iDia) / (iMinSpacing + iDia) + 1 'number of bars per layer govered by spacing
        End If
        dAreaOfSingleBar = WorksheetFunction.Pi() / 4 * iDia * iDia
        iTotalNumberOfBarsRequired = WorksheetFunction.Ceiling_Math(dAreaRequired / dAreaOfSingleBar, 1)
        iTotalNumberOfBarsRequired = WorksheetFunction.Max(iTotalNumberOfBarsRequired, iMinNoOfBars)
        iNumberOfLayers = WorksheetFunction.Ceiling_Math(iTotalNumberOfBarsRequired / iNumberOfBars, 1) ' Final number of layers of rebar
        'If fractional number of layers are coming, optimize the final layer
        If iNumberOfLayers > 1 And dAreaRequired < iTotalNumberOfBarsRequired * dAreaOfSingleBar Then
            'Get solution for the last layer
            dAreaRequired_base = (iNumberOfLayers - 1) * iNumberOfBars * dAreaOfSingleBar
            dAreaRequired_add = dAreaRequired - (iNumberOfLayers - 1) * iNumberOfBars * dAreaOfSingleBar
            'Base solution
            iaRebarOptions_base(2, iOption) = iNumberOfBars
            iaRebarOptions_base(3, iOption) = iNumberOfLayers - 1
            
            'Solution for inner most layer
            'If iDia = 32 Then
                'MsgBox (iDia & " dia bar")
            'End If
            
            If iNumberOfBars Mod 2 = 0 Then 'If the number of bars in base layer are even, odd number of bars should not be used in the last layer
                bOddNumberAcceptable = False
            Else
                bOddNumberAcceptable = True
            End If
            'ProvideOptimumRebar_internal(iFace,dAreaRequired_add,iMinNoOfBars,iMaxNoOfBars,bOddNumberAcceptable,iLargestDia)
            iDia = WorksheetFunction.Index(ProvideOptimumRebar_internal(iFace, dAreaRequired_add, 2, iNumberOfBars, bOddNumberAcceptable, iDia), 1) 'Returned value = [Dia, Number of bars, Number of layers] for the given dia
            iNumberOfBars = WorksheetFunction.Index(ProvideOptimumRebar_internal(iFace, dAreaRequired_add, 2, iNumberOfBars, bOddNumberAcceptable, iDia), 2) 'Returned value = [Dia, Number of bars, Number of layers] for the given dia
            iNumberOfLayers = WorksheetFunction.Index(ProvideOptimumRebar_internal(iFace, dAreaRequired_add, 2, iNumberOfBars, bOddNumberAcceptable, iDia), 3) 'Returned value = [Dia, Number of bars, Number of layers] for the given dia
            iaRebarOptions_add(1, iOption) = iDia
            iaRebarOptions_add(2, iOption) = iNumberOfBars
            iaRebarOptions_add(3, iOption) = iNumberOfLayers
        Else
            'Get the basic solution
            iNumberOfBars = WorksheetFunction.Ceiling_Math(iTotalNumberOfBarsRequired / iNumberOfLayers, 1) ' Final number of bars per layer
            iaRebarOptions_base(2, iOption) = iNumberOfBars
            iaRebarOptions_base(3, iOption) = iNumberOfLayers
        End If
    Next iOption
    
    
    
    '================================
    '======== OPTION RATING =========
    '================================
    
    
    
    
    'Deciding which option is best

    'Preferred bar dia   -> 3 points to star with
    'Number of layers    -> 2 points reduced for each layer after first layer
    'Bar dia             -> smaller the better; +1 if less than Width/20 and -1 if more than width/15 (2 mm tolerance)
    'Area provided       -> Lowest area above Areq is preferred. Mark = Closest Integer 10 x (Aprovided/Arequired -1)*5 ---- Note: to be deducted as lowest is better
    Dim iPreferenceMarks As Integer
    Dim iLayerMarks As Integer
    Dim iEconomyMarks As Integer
    Dim iSpacingMarks As Integer
    Dim dActualSpacing As Double 'Actul spacing between bars
    iPreferenceMarks = 3
    iLayerMarks = -2        'Please note that it is negative
    iEconomyMarks = 2
    iSpacingMarks = 2 'Added if less than iMinSpacing +25 and deducted if above iMinSpacing+50

    For iOption = 1 To 6
        iDia = iaRebarOptions_base(1, iOption) 'size of rebar

        'Initialising
        iaRebarOptions_base(4, iOption) = 0
        'Preference Marks
        If iDia = piPreferredDia(iFace) Then 'If the current size is the preferred dia
            iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) + iPreferenceMarks
        End If
        
        'Number of layers marks
        iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) + (iaRebarOptions_base(3, iOption) + iaRebarOptions_add(3, iOption) - 1) * iLayerMarks
        'Bar dia marks
        If iDia <= (piWidth / 20) + 2 Then
            iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) + 1
        ElseIf iDia > (piWidth / 15) + 2 Then
            iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) - 1
        End If
        'Marks based on bar spacing
        
        dActualSpacing = (piWidth - 2 * piClearCover - 2 * piLinkDia - iDia * iaRebarOptions_base(2, iOption)) / (iaRebarOptions_base(2, iOption) - 1)
        If dActualSpacing < iMinSpacing + 25 Then
            iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) + iSpacingMarks
        ElseIf dActualSpacing > iMinSpacing + 25 Then
            iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) - iSpacingMarks
        End If
        
        'Number of bars marks --- Just that people are crazy about number of bars.
        If iaRebarOptions_base(2, iOption) = WorksheetFunction.Floor_Math(piWidth / 100, 1) Then
            iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) + 1
        End If
        
        
        
        'Marks based on the area provided
        dAreaProvided1 = iaRebarOptions_base(3, iOption) * iaRebarOptions_base(2, iOption) * (WorksheetFunction.Pi() / 4 * iaRebarOptions_base(1, iOption) ^ 2)
        dAreaProvided2 = iaRebarOptions_add(3, iOption) * iaRebarOptions_add(2, iOption) * (WorksheetFunction.Pi() / 4 * iaRebarOptions_add(1, iOption) ^ 2)
        
        
        dAreaProvided = dAreaProvided1 + dAreaProvided2
        'Safety check
        If dAreaProvided < dAreaRequired Then
            MsgBox "Error: Something unexpected has happened and the area I calculated is less than what is required. Please check the code" & Chr(13) & "Contact: Harikrishnan Panikkaveettil", vbCritical, "Error"
        ElseIf dAreaRequired > 0 Then
            iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iOption) - Round((dAreaProvided / dAreaRequired - 1) * 10 * iEconomyMarks, 0)
        End If
        
    Next iOption

    'Final result
    iPreferredOption = 1 'smallest dia for now
    'Checking if something else is better
    For iOption = 2 To 6
        If iaRebarOptions_base(4, iOption) > iaRebarOptions_base(4, iPreferredOption) Then
            iPreferredOption = iOption
        ElseIf iaRebarOptions_base(4, iOption) = iaRebarOptions_base(4, iPreferredOption) Then
            iBonus_Option = 0
            iBonus_current = 0
            iDia = iaRebarOptions_base(1, iOption)
            'Spacing marks
            dActualSpacing = (piWidth - 2 * piClearCover - 2 * piLinkDia - iDia * iaRebarOptions_base(2, iOption)) / (iaRebarOptions_base(2, iOption) - 1)
            If dActualSpacing < iMinSpacing + 25 Then
                iBonus_Option = iBonus_Option + iSpacingMarks
            ElseIf dActualSpacing > iMinSpacing + 25 Then
                iBonus_Option = iBonus_Option - iSpacingMarks
            End If
            'Number of bars marks --- Just that people are crazy about number of bars.
            If iaRebarOptions_base(2, iOption) = WorksheetFunction.Floor_Math(piWidth / 100, 1) Then
                iBonus_Option = iBonus_Option + 1
            End If
            
            If iDia <= (piWidth / 20) + 2 And iDia >= (piWidth / 25) + 2 Then
                iBonus_Option = iBonus_Option + 1
            ElseIf iDia > (piWidth / 15) + 2 Then
                iBonus_Option = iBonus_Option - 1
            End If
            
            'for preferred
            iDia = iaRebarOptions_base(1, iPreferredOption)
            dActualSpacing = (piWidth - 2 * piClearCover - 2 * piLinkDia - iDia * iaRebarOptions_base(2, iPreferredOption)) / (iaRebarOptions_base(2, iPreferredOption) - 1)
            If dActualSpacing < iMinSpacing + 25 Then
                iBonus_current = iBonus_current + iSpacingMarks
            ElseIf dActualSpacing > iMinSpacing + 25 Then
                iBonus_current = iBonus_current - iSpacingMarks
            End If
            
            'for current preferred
            If iaRebarOptions_base(2, iPreferredOption) = WorksheetFunction.Floor_Math(piWidth / 100, 1) Then
                iBonus_current = iBonus_current + 1
            End If
            
            If iDia <= (piWidth / 20) + 2 And iDia >= (piWidth / 25) + 2 Then
                iBonus_current = iBonus_current + 1
            ElseIf iDia > (piWidth / 15) + 2 Then
                iBonus_current = iBonus_current - 1
            End If
            If iBonus_Option > iBonus_current Then
                iPreferredOption = iOption
            End If
        End If
    Next iOption
    
    '============ Ouput Display format ==============
    If iFace = 1 Then
        piaBot_reinforcement(1) = iaRebarOptions_base(1, iPreferredOption)
        piaBot_reinforcement(2) = iaRebarOptions_base(2, iPreferredOption)
        piaBot_reinforcement(3) = iaRebarOptions_base(3, iPreferredOption)
        piaBot_reinforcement(4) = iaRebarOptions_add(1, iPreferredOption)
        piaBot_reinforcement(5) = iaRebarOptions_add(2, iPreferredOption)
        piaBot_reinforcement(6) = iaRebarOptions_add(3, iPreferredOption)
    ElseIf iFace = 2 Then
    'piaTop_reinforcement
        piaTop_reinforcement(1) = iaRebarOptions_base(1, iPreferredOption)
        piaTop_reinforcement(2) = iaRebarOptions_base(2, iPreferredOption)
        piaTop_reinforcement(3) = iaRebarOptions_base(3, iPreferredOption)
        piaTop_reinforcement(4) = iaRebarOptions_add(1, iPreferredOption)
        piaTop_reinforcement(5) = iaRebarOptions_add(2, iPreferredOption)
        piaTop_reinforcement(6) = iaRebarOptions_add(3, iPreferredOption)
    Else
        Debug.Print "Incorrect iFace information passed to ProvideOptimumRebar"
    End If

End Sub

                
                'OptimumRebar_internal(iFace,dAreaRequired_add,iMinNoOfBars,iMaxNoOfBars,bOddNumberAcceptable,iLargestDia)
                Private Function ProvideOptimumRebar_internal(ByVal iFace As Integer, ByVal dAreaRequired As Double, Optional ByVal iMinNoOfBars = 2, Optional ByVal iMaxNoOfBars = 100, Optional ByVal bOddNumberAcceptable As Boolean = False, Optional ByVal iLargestDia As Integer = 32) As Variant
                    'This is an internal function for OptimumRebar Function . The difference is that it returns the values as an array instead of a text unlike the other function
                    Dim iNumberOfBars As Integer
                    Dim iNumberOfLayers As Integer
                    Dim iOption As Integer 'Looping variable for size
                    Dim iDia As Integer
                    Dim iAreaOfSingleBar As Double ' Area of a single rebar of size iDia
                    Dim iTotalNumberOfBarsRequired As Integer
                    Dim dAreaProvided As Double
                    Dim iPreferredOption As Integer 'Preferred Solution
                    'Rebar Dia                : 10 12 16 20 25 32
                    'Number of bars per layer : 00 00 00 00 00 00
                    'Number of layers         : 01 01 01 01 01 01
                    'Marks                    : 00 00 00 00 00 00       Highest marks is recommended
                
                
                    '======== Initialising =========
                    If iFace = 1 Then
                        iMinSpacing = piMinClearSpacingBottom
                    ElseIf iFace = 2 Then
                        iMinSpacing = piMinClearSpacingTop
                    Else
                        Debug.Print "OptimumRebar is asked to design for something other than Bottom/Top bars iFace = " & iFace
                    End If
                    Dim iaRebarOptions(1 To 4, 1 To 6) As Long
                    iaRebarOptions(1, 1) = 10
                    iaRebarOptions(1, 2) = 12
                    iaRebarOptions(1, 3) = 16
                    iaRebarOptions(1, 4) = 20
                    iaRebarOptions(1, 5) = 25
                    iaRebarOptions(1, 6) = 32
                
                    '======== Prioity list =========
                    'Rebar Dia                : 10 12 16 20 25 32
                    'Number of bars per layer : 00 00 00 00 00 00
                    'Number of layers         : 01 01 01 01 01 01
                    '<<<<<<<<<<< Yet to workout the strategy >>>>>>`
                
                
                
                    'Looping
                    For iOption = 1 To 6
                        iDia = iaRebarOptions(1, iOption) 'Size of rebar
                        'Spacing between the centres of extreme bars are given by (piWidth - 2 * piClearCover - 2 * piLinkDia - iDia)
                        'If the clear spacing is S and dia of bar is d, there will be N number of bar dias for N number of spacing
                        'Thus, Number of bars = 1 + N
                        
                        iNumberOfBars = (piWidth - 2 * piClearCover - 2 * piLinkDia - iDia) / (iMinSpacing + iDia) + 1 'number of bars per layer govered by spacing
                        'Acceptanc criteria
                        If iNumberOfBars < iMinNoOfBars Then
                            iNumberOfBars = iMinNoOfBars
                            If iNumberOfBars Mod 2 <> 0 And Not (bOddNumberAcceptable) Then
                                iNumberOfBars = iNumberOfBars + 1
                            End If
                        End If
                        If iNumberOfBars > iMaxNoOfBars Then
                            iNumberOfBars = iMaxNoOfBars
                            If iNumberOfBars Mod 2 <> 0 And Not (bOddNumberAcceptable) Then
                                iNumberOfBars = iNumberOfBars - 1
                            End If
                        End If
                        
                        dAreaOfSingleBar = 3.14 / 4 * iDia * iDia
                        iTotalNumberOfBarsRequired = WorksheetFunction.Ceiling_Math(dAreaRequired / dAreaOfSingleBar, 1)
                        iNumberOfLayers = WorksheetFunction.Ceiling_Math(iTotalNumberOfBarsRequired / iNumberOfBars, 1) ' Final number of layers of rebar
                        iNumberOfBars = WorksheetFunction.Ceiling_Math(iTotalNumberOfBarsRequired / iNumberOfLayers, 1) ' Final number of bars per layer
                        'Acceptanc criteria
                        If iNumberOfBars < iMinNoOfBars Then
                            iNumberOfBars = iMinNoOfBars
                        End If
                        If iNumberOfBars Mod 2 <> 0 And Not (bOddNumberAcceptable) Then
                            iNumberOfBars = iNumberOfBars + 1
                        End If
                        
                        
                        
                        iaRebarOptions(2, iOption) = iNumberOfBars
                        iaRebarOptions(3, iOption) = iNumberOfLayers
                    Next iOption
                
                    'Deciding which option is best
                
                    'Preferred bar dia   -> 3 points to star with
                    'Number of layers    -> 2 points reduced for each layer after first layer
                    'Bar dia             -> smaller the better; +1 if less than Width/20 and -1 if more than width/15 (2 mm tolerance)
                    'Area provided       -> Lowest area above Areq is preferred. Mark = Closest Integer 10 x (Aprovided/Arequired -1)*5 ---- Note: to be deducted as lowest is better
                    Dim iPreferenceMarks As Integer
                    Dim iLayerMarks As Integer
                    Dim iEconomyMarks As Integer
                    iPreferenceMarks = 3
                    iLayerMarks = -3        'Please note that it is negative
                    iEconomyMarks = 1
                
                
                    For iOption = 1 To 6
                        iDia = iaRebarOptions(1, iOption) 'size of rebar
                        'If iDia = 12 Or iDia = 20 Or iDia = 25 Then
                            'MsgBox (iDia & " dia bar")
                        'End If
                        'Initialising
                        iaRebarOptions(4, iOption) = 0
                        'Preference Marks
                        If iDia = piPreferredDia(iFace) Then 'If the current size is the preferred dia
                            iaRebarOptions(4, iOption) = iaRebarOptions(4, iOption) + iPreferenceMarks
                        End If
                        'Number of layers marks
                        iaRebarOptions(4, iOption) = iaRebarOptions(4, iOption) + (iaRebarOptions(3, iOption) - 1) * iLayerMarks
                        'Bar dia marks
                        If iDia <= (piWidth / 20) + 2 Then
                            iaRebarOptions(4, iOption) = iaRebarOptions(4, iOption) + 1
                        ElseIf iDia > (piWidth / 15) + 2 Then
                            iaRebarOptions(4, iOption) = iaRebarOptions(4, iOption) - 1
                        End If
                        
                        
                        If iDia > iLargestDia Then
                            iaRebarOptions(4, iOption) = iaRebarOptions(4, iOption) - 100
                        End If
                        'Points based on the area provided
                        dAreaProvided = iaRebarOptions(3, iOption) * iaRebarOptions(2, iOption) * (3.14 / 4 * iaRebarOptions(1, iOption) ^ 2) ' Area
                        'Safety check
                        If dAreaProvided < dAreaRequired Then
                            MsgBox "Error: Something unexpected has happened and the area I calculated is less than what is required. Please check the code" & Chr(13) & "Contact: Harikrishnan Panikkaveettil", vbCritical, "Error"
                        ElseIf dAreaRequired > 0 Then
                            iaRebarOptions(4, iOption) = iaRebarOptions(4, iOption) - Round((dAreaProvided / dAreaRequired - 1) * 10 * iEconomyMarks, 0)
                        End If
                        
                    Next iOption
                
                    'Final result
                    iPreferredOption = 1 'smallest dia for now
                    'Checking if something else is better
                    For iOption = 2 To 6
                        If iaRebarOptions(4, iOption) > iaRebarOptions(4, iPreferredOption) Then
                            iPreferredOption = iOption
                        End If
                    Next iOption
                
                    ProvideOptimumRebar_internal = Array(iaRebarOptions(1, iPreferredOption), iaRebarOptions(2, iPreferredOption), iaRebarOptions(3, iPreferredOption))
                End Function

Private Sub ProvdeShearLinks()
    'Design for Shear
    '10-100 '10-150
End Sub




'================================================
'============ OPTIMISATION OF OPTIONS ===========
'================================================

Public Sub RefineTheReinforcement(ByVal iFace As Integer, Optional ByVal iNumberOfBarsToBeForced As Integer = 0, Optional ByVal iNewPreferredDia As Integer)
    'Refines the reinforcement detail for iFace based on the prefrences passed
    If Not IsMissing(iNewPreferredDia) Then 'If nothing is provided, never mind; otherwise update the preference
        piPreferredDia(iFace) = iNewPreferredDia
    End If
    If iFace = 1 Then 'If Bottom reinforcement is to be refined
        Call EstimateBotttomReinforcement(iNumberOfBarsToBeForced)
    ElseIf iFace = 2 Then 'If top reinforcement is to be refined
        Cal EstimateTopReinforcement(iNumberOfBarsToBeForced)
    Else
        Debug.Print "ERROR: " & Me.Name & " is asked to refine the reinforcement of face " & iFace
    End If
End Sub

