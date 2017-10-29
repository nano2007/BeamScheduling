VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RCBeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private psDwgName As String 'Name as referenced in drawings
Private psName As String 'Name as in ETABS
Private piWidth As Integer 'Width of RC beam (mm)
Private piDepth As Integer 'Depth of RC beam (mm)
Private piClearCover As Integer 'Clear cover to rebar
Private piLinkDia As Integer 'Dia of outer reinforcement links/stirrups
Private piBeamType As Integer
Private piMinClearSpacingBottom As Integer
Private piMinClearSpacingTop As Integer
'Sections
Private cCrossSections(1 To 3) As RCSection  'End-I, Middle and End-J
'Structure of pdAreaProvided
'i-top  m-top   j-top
'i-bot  m-bot   j-bot
Private pdAreaProvided(1 To 2, 1 To 3) As Double   'Area provided at each section;
Private psReinforcementProvided(1 To 2, 1 To 3) As String   'Same structure as pdAreaProvided but as description
Private piPreferredDia(1 To 2)  As Integer '{Bottom Top}
Private piConcreteGrade As Integer
Private piSteelGrade As Integer
Private pdLength As Double 'Lengtth of beam --> ignoring z offset

Private Sub Class_Initialize()
    Dim iLocation As Integer
    For iLocation = 1 To 3
        Set cCrossSections(iLocation) = New RCSection
         cCrossSections(iLocation).Cover = 40
         cCrossSections(iLocation).LinkDia = 10
         cCrossSections(iLocation).PreferredBarDia(1) = 0
         cCrossSections(iLocation).PreferredBarDia(2) = 0
         cCrossSections(iLocation).MinClearSpacingBottom = 50
         cCrossSections(iLocation).MinClearSpacingTop = 75
    Next iLocation
End Sub


'======= Name Properties =====
Public Property Get Name() As String
    Name = psName
End Property
Public Property Let Name(ByVal ETABS_Name As String)
    Dim iLocation As Integer
    psName = ETABS_Name
    For iLocation = 1 To 3
         cCrossSections(iLocation).Name = ETABS_Name & " Section " & iLocation
    Next iLocation
End Property

Public Property Get NameInDrawing() As String
    NameInDrawing = psDwgName
End Property
Public Property Let NameInDrawing(ByVal BeamNameAsInDrawing As String)
    Dim iLocation As Integer
    psDwgName = BeamNameAsInDrawing
    For iLocation = 1 To 3
         cCrossSections(iLocation).NameInDrawing = BeamNameAsInDrawing & " Section " & iLocation
    Next iLocation
End Property
'Sectional Properties
Public Property Get Width() As Integer
    Width = piWidth
End Property
Public Property Let Width(ByVal iBeamWidth As Integer)
    Dim iLocation As Integer
    piWidth = iBeamWidth
    For iLocation = 1 To 3
         cCrossSections(iLocation).Width = iBeamWidth
    Next iLocation
End Property
Public Property Get Depth() As Integer
    Depth = piDepth
End Property
Public Property Let Depth(ByVal iDepthOfBeam As Integer)
    Dim iLocation As Integer
    piDepth = iDepthOfBeam
    For iLocation = 1 To 3
         cCrossSections(iLocation).Depth = iDepthOfBeam
    Next iLocation
End Property

Public Property Get Cover() As Integer
    Cover = piClearCover
End Property
Public Property Let Cover(ByVal iClearCover As Integer)
    Dim iLocation As Integer
    piClearCover = iClearCover
    For iLocation = 1 To 3
         cCrossSections(iLocation).Cover = iClearCover
    Next iLocation
End Property
Public Property Get LinkDia() As Integer
    LinkDia = piLinkDia
End Property
Public Property Let LinkDia(ByVal iLinkDia As Integer)
    Dim iLocation As Integer
    piLinkDia = iLinkDia
    For iLocation = 1 To 3
         cCrossSections(iLocation).LinkDia = iLinkDia
    Next iLocation
End Property
'Reinforcemetn demands
Public Property Get AreaRequiredTop(ByVal iLocation As Integer) As Double
    AreaRequiredTop = cCrossSections(iLocation).AreaRequiredTop
End Property
Public Property Let AreaRequiredTop(ByVal iLocation As Integer, dTopReinforcementAreaRequired As Double)
    cCrossSections(iLocation).AreaRequiredTop = dTopReinforcementAreaRequired
End Property
Public Property Get AreaRequiredBottom(ByVal iLocation As Integer) As Double
    AreaRequiredBottom = cCrossSections(iLocation).AreaRequiredBottom
End Property
Public Property Let AreaRequiredBottom(ByVal iLocation As Integer, ByVal dBottomReinforcementAreaRequired As Double)
    cCrossSections(iLocation).AreaRequiredBottom = dBottomReinforcementAreaRequired
End Property
Public Property Get AreaRequiredTorsion(ByVal iLocation As Integer) As Double
    AreaRequiredTorsion = cCrossSections(iLocation).AreaRequiredTorsion
End Property
Public Property Let AreaRequiredTorsion(ByVal iLocation As Integer, ByVal dTorReinforcementAreaRequired As Double)
    cCrossSections(iLocation).AreaRequiredTorsion = dTorReinforcementAreaRequired
End Property
Public Property Get BeamType() As Integer
    BeamType = piBeamType
End Property
Public Property Let BeamType(ByVal iBeamType As Integer)
    piBeamType = iBeamType
End Property
Public Property Get MinClearSpacingBottom() As Integer
    MinClearSpacingBottom = piMinClearSpacingBottom
End Property
Public Property Let MinClearSpacingBottom(ByVal iMinClearSpacingBottom As Integer)
    piMinClearSpacingBottom = iMinClearSpacingBottom
    For iLocation = 1 To 3
         cCrossSections(iLocation).MinClearSpacingBottom = iMinClearSpacingBottom
    Next iLocation
End Property
Public Property Get MinClearSpacingTop() As Integer
    MinClearSpacingTop = piMinClearSpacingTop
End Property
Public Property Let MinClearSpacingTop(ByVal iMinClearSpacingTop As Integer)
    piMinClearSpacingTop = iMinClearSpacingTop
    For iLocation = 1 To 3
         cCrossSections(iLocation).MinClearSpacingTop = iMinClearSpacingTop
    Next iLocation
End Property
Public Property Get PreferredBarDia(ByVal iFace As Integer) As Integer 'Zero for automatic
    PreferredBarDia = piPreferredDia(iFace)
End Property
Public Property Let PreferredBarDia(ByVal iFace As Integer, ByVal iPreferredBarDia As Integer)
    piPreferredDia(iFace) = iPreferredBarDia
    For iLocation = 1 To 3
        cCrossSections(iLocation).PreferredBarDia(iFace) = iPreferredBarDia
    Next iLocation
End Property
Public Property Get ConcreteGrade() As Integer
    ConcreteGrade = piConcreteGrade
End Property
Public Property Let ConcreteGrade(ByVal iConcreteGrade As Integer)
    piConcreteGrade = iConcreteGrade
    For iLocation = 1 To 3
        cCrossSections(iLocation).ConcreteGrade = iConcreteGrade
    Next iLocation
End Property
Public Property Get SteelGrade() As Integer
    SteelGrade = piSteelGrade
End Property
Public Property Let SteelGrade(ByVal iSteelGrade As Integer)
    piSteelGrade = iSteelGrade
    For iLocation = 1 To 3
        cCrossSections(iLocation).SteelGrade = iSteelGrade
    Next iLocation
End Property

Public Property Get Length() As Double
    Length = pdLength
End Property
Public Property Let Length(ByVal dLength As Double)
    pdLength = dLength
End Property


'================================================
'================ Other methods =================
'================================================

'Function that returns the area provided at a specific location or in all locations
Public Function ReinforcementAreaProvided(Optional ByVal iLocation As Integer, Optional ByVal iFace As Integer)
    'iLocation = 1 to 3 for End-I, Middle and End-J respectively
    'iFace = 1 to 2 for Bottom face and top face respectively
    Dim dAreaProvided(3, 2) As Double
    If IsMissing(iLocation) Then
        For iLocation = 1 To 3
            If IsMissing(iFace) Then
                For iFace = 1 To 2
                    dAreaProvided(iLocation, iFace) = WorksheetFunction.Index(cCrossSections(iLocation).GetAreaOfSteelProvided(), iFace)
                Next iFace
                ReinforcementAreaProvided = dAreaProvided
            Else
                'Not addressed properly yet
                Debug.Print "This case has not been addressed properly yet in RCBeam.ReinforcementAreaProvided()"
            End If
        Next iLocation
    Else
        If IsMissing(iFace) Then
            For iFace = 1 To 2
                dAreaProvided(iLocation, iFace) = WorksheetFunction.Index(cCrossSections(iLocation).GetAreaOfSteelProvided(), iFace)
            Next iFace
            ReinforcementAreaProvided = dAreaProvided
        Else
            dAreaProvided(iLocation, iFace) = WorksheetFunction.Index(cCrossSections(iLocation).GetAreaOfSteelProvided(), iFace)
            ReinforcementAreaProvided = dAreaProvided(iLocation, iFace)
        End If
    End If
End Function
'Returns the reinforcement As array
Public Function GetReinforcement(ByVal iLocation As Integer)
    Dim iReinforcementArray(1 To 3, 1 To 6) As Integer  'Each column one face reinforcement; each layer one section
    ' 1 to 3 - iFace : [Bottom, Top, Side]
    ' 1 to 6 - Reinforcement array
    Dim iFace As Integer
    Dim iRebarComponent As Integer
    For iFace = 1 To 2 'For each face (bottom/top)
        For iRebarComponent = 1 To 6 ' each of the array items
            iReinforcementArray(iFace, iRebarComponent) = WorksheetFunction.Index(cCrossSections(iLocation).GetReinforcement(), iRebarComponent, iFace)
        Next iRebarComponent
    Next iFace
    iReinforcementArray(3, 1) = WorksheetFunction.Index(cCrossSections(1).GetReinforcement(), 1, 3) '1 => Row (here dia of side face); 3=> 3rd component (col - here side face bars)
    iReinforcementArray(3, 1) = WorksheetFunction.Index(cCrossSections(1).GetReinforcement(), 2, 3) '2 => Row (here spacing of side face); 3=> 3rd component (col - here side face bars)
    GetReinforcement = iReinforcementArray
End Function
'Returns a  3 x 3 array with each column representng the locations and each row representing bottom ,top and side faces
Public Function GetReinforcementDescriptionAsString()
    Dim iLocation As Integer
    Dim sDescription(1 To 3, 1 To 3) As String 'Location (End-I, Middle & End-J) , face (Bottom, Top, Side)
    For iLocation = 1 To 3
        sDescription(iLocation, 1) = WorksheetFunction.Index(cCrossSections(iLocation).GetReinforcementAsString(), 1) 'Bottom face
        sDescription(iLocation, 2) = WorksheetFunction.Index(cCrossSections(iLocation).GetReinforcementAsString(), 2) 'Top face
        sDescription(iLocation, 3) = WorksheetFunction.Index(cCrossSections(iLocation).GetReinforcementAsString(), 3) 'Side face
    Next iLocation
    GetReinforcementDescriptionAsString = sDescription ' 3 x 3 array with (iLocation,iFace); iFace => 1=Bottom, 2=Top, 3=side
End Function
Private Function GetDevelopmentLength(ByVal iLocation As Integer, ByVal iFace As Integer) As Double
    Dim iDia As Integer
    Dim iFc As Integer 'Grade of concrete
    Dim iFy As Integer 'Grade of steel
    Dim dPsi_t As Double
    Dim dPsi_e As Double
    Dim dLambda As Double
    Dim dFactor As Double
    'Initialising
    iDia = WorksheetFunction.Index(cCrossSections(iLocation).GetReinforcement(), 1, iFace)
    iFc = Me.ConcreteGrade
    iFy = Me.SteelGrade
    'Section 12.201(a) of ACI 318M-08
    If iFace = 1 Then
        dPsi_t = 1#
    ElseIf (piDepth - piClearCover - piLinkDia) > 300 Then
        dPsi_t = 1.3
    Else
        dPsi_t = 1#
    End If
    dPsi_e = 1#  'section 12.2.4 (b) of ACI 318M-08
    dLambda = 1#  'Normal weight concrete; section 12.2.4(d) of ACI 318M-08
    
    If iDia <= 19 Then
        dFactor = 2.1
    ElseIf iDia >= 22 Then
        dFactor = 1.7
    Else
        dFactor = 2.1 + (1.7 - 2.1) / (22 - 19) * (iDia - 19)
    End If
    GetDevelopmentLength = (iFy * dPsi_t * dPsi_e) / (dFactor * dLambda * VBA.Sqr(iFc)) * iDia
End Function




'================================================
'============ OPTIMISATION METHODS ==============
'================================================
Public Sub OptimiseMidTopReinforcement()
    'Optimises the top bar at mid-section (iLocation = 2) of the beam.
    '   If Both ends have a given bar dia/numbers we would prefer same dia/numbers
    '   If Type 2 beam, both dia and
    
    '-----------------------------------------
    '      What to enforce (dia/Number)      |
    '-----------------------------------------
    'Beam Length    | Low Shear | High Shear |
    '-----------------------------------------
    ' Short         |Dia & N    | Dia & N    |
    ' Long          |Dia        | N          |
    ' Very Long     |(none)     | N          |
    '-----------------------------------------
    
    
    ' Economby estimation
    ' Due to laps (wastage), 3.5 LTA = Same area as larger dia
    ' Due to labout charges, 1.5 LTA = Same cost
    '---------------------- say total 5 LTA
    ' Assuming splicing to happen at L/3
    '---------- Net => use same dia for middle also if 1/3 Span < 5 LTS or Span < 15 LTS
    '                   This typically means that we should not curtail bars
    
    Dim iDia(1 To 3) As Integer '{End I Dia, Middle Dia , End J Dia}
    Dim iNumbers(1 To 3) As Integer  '{End I N, Middle N, End N}
    
    
End Sub


