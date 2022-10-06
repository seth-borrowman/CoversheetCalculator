Attribute VB_Name = "Module2"
Option Explicit
    'Transfer all cell values from the input file to variables
    Public FormulationFile As String
    'FormulationFile = Range("C7").Value
    Public FormulationType As String
    'FormulationType = Range("C8").Value
    Public Refs As Integer
    'Refs = Range("C9").Value
    Public numMix As Double
    'numMix = Range("E9").Value
    Public Aliquots As Integer
    'Aliquots = Range("G9").Value
    Public FormulationType1 As String
    'FormulationType1 = Range("C10").Value
    Public Refs1 As Integer
    'Refs1 = Range("C11").Value
    Public SumRefs As Integer
    'SumRefs = Refs + Refs1
    Public numMix1 As Double
    'numMix1 = Range("E11").Value
    Public Aliquots1 As Integer
    'Aliquots1 = Range("G11").Value
    Public SumAliquots As Integer
    'SumAliquots = Aliquots + Aliquots1
    Public Tubes As Integer
    'Tubes = SumRefs * (numMix + numMix1)
    Public NonSynth As Integer
    'NonSynth = Range("C12").Value
    Public ShipCondit As String
    'ShipCondit = Range("C13").Value
    Public HandOD As String
    'HandOD = Range("C14").Value
    Public BECC As String
    'BECC = Range("C15").Value
    Public Labels As String
    'Labels = Range("C16").Value
    Public SpecSheets As String
    'SpecSheets = Range("C17").Value
    Public ManReview As String
    'Review = Range("C18").Value
    Public Traces As String
    'Traces = Range("C19").Value
    Public Packaging As String
    'Packaging = Range("C20").Value
    Public Partial As String
    'Partial = Range("C21").Value
    Public Adapter As String
    'Adapter = Range("C22").Value

    Public SingleSetup As Double
    Public DuplexSetup As Double
    Public MixSetup As Double
    
    'Create Hours variable to be exported
    Public HoursVar As Double
    Public HoursVarRound As Integer
    Public HoursVarArray(11) As Double
    
    
    
    Function File()
    'Change formulation hours based on the formulation file selected
    FormulationFile = Range("C7").Value
    If FormulationFile = "CPF Master Formulation Template" Then
        HoursVar = HoursVar
        'Add different formulation files once timing is complete
    End If
    
    File = HoursVar
End Function
Function SingleOligo(R As Integer, A As Integer)
    'Evaluate added time if the formulation type is a single oligo R is ref#s A is aliquots
    
    'Max oligos on 1 file for single oligos with BECC is 7
    BECC = Range("C15").Value
    If BECC = "Yes" Then
        Dim SingleFiles As Integer
        SingleFiles = Application.WorksheetFunction.RoundUp((R / 7), 0)
    Else
        SingleFiles = Application.WorksheetFunction.RoundUp((R / 96), 0)
    End If
    
    'Add time for organization
    HoursVar = HoursVar + ((Range("M4").Value / 96) * R)
    'Add time for form file
    HoursVar = HoursVar + (Range("M5").Value * SingleFiles)
    'Add time for post-form and norm/form
    HoursVar = HoursVar + (Range("M9").Value * SingleFiles)
    SingleSetup = (Range("M4").Value) + (Range("M5").Value) + (Range("M9").Value)
    'Add time for hydrating oligos
    HoursVar = HoursVar + (Range("M6").Value * R)
    'Add time for vortexing oligos
    HoursVar = HoursVar + (Range("M7").Value * R)
    'Add time for transferring oligos
    HoursVar = HoursVar + (Range("M8").Value * R)
    
    'Add time for aliquots
    'Label
    HoursVar = HoursVar + (Range("M14").Value / 96 * A)
    'Create aliquots
    HoursVar = HoursVar + (Range("M15").Value / 96 * A)
    'Grav test
    If A > 10 Then
        HoursVar = HoursVar + (Range("M16").Value * R)
    End If
    'Cap
    HoursVar = HoursVar + (Range("M17").Value / 96 * A)
    
    SingleOligo = HoursVar
End Function
Function Duplex(R As Integer, N As Double, A As Integer)
    'Evaluate added time if the formulation type is a duplex
    
    'Number of form files needed - 14 duplexes per form file
    Dim DuplexFiles As Integer
    DuplexFiles = Application.WorksheetFunction.RoundUp((R / 15), 0)

    'Add time for organization
    HoursVar = HoursVar + (Range("P4").Value / 96 * (R * 2))
    'Add time for form file
    HoursVar = HoursVar + (Range("P5").Value * DuplexFiles)
    'Add time for post-form and norm/form
    HoursVar = HoursVar + (Range("P10").Value * DuplexFiles)
    DuplexSetup = (Range("P4").Value) + (Range("P5").Value) + (Range("P10").Value)
    'Add time for hydrating oligos
    HoursVar = HoursVar + (Range("P6").Value * (R * 2))
    'Add time for vortexing oligos
    HoursVar = HoursVar + (Range("P7").Value * (R * 2))
    'Add time for creating top level ref
    HoursVar = HoursVar + (Range("P8").Value * R)
    'Add time for transferring oligos
    HoursVar = HoursVar + (Range("P9").Value * (R * 2))
        
    'Add time for aliquots
    'Label
    HoursVar = HoursVar + (Range("M14").Value / 96 * A)
    'Create aliquots
    HoursVar = HoursVar + (Range("M15").Value / 96 * A)
    'Grav test
    If A > 10 Then
        HoursVar = HoursVar + (Range("M16").Value * R)
    End If
    'Bulk Hand OD
    If A > 10 Then
        HoursVar = HoursVar + (Range("M22").Value * R)
    Else
        'HoursVar = HoursVar not necessary, inlcuded to aid clarity
        HoursVar = HoursVar
    End If
    'Cap
    HoursVar = HoursVar + (Range("M17").Value / 96 * A)
    
    Duplex = HoursVar
End Function
Function Mix(R As Integer, N As Double, A As Integer)
    'Evaluate added time if the formulation type is a mix
    
    'Number of form files needed - 22 2-oligo mixes per form file, 13 3-oligo mixes, 12 4-oligo mixes, 1 multi-oligo mix
    Dim MixFiles As Integer
    If N = 2 Then
        MixFiles = Application.WorksheetFunction.RoundUp((R / 22), 0)
    ElseIf N = 3 Then
        MixFiles = Application.WorksheetFunction.RoundUp((R / 13), 0)
    ElseIf N = 4 Then
        MixFiles = Application.WorksheetFunction.RoundUp((R / 12), 0)
    ElseIf N > 4 Then
        MixFiles = R
    End If
    
    'Add time for organization
    HoursVar = HoursVar + (Range("S4").Value / 96 * (R * N))
    'Add time for form file
    HoursVar = HoursVar + (Range("S5").Value * MixFiles)
    'Add time for post-form and norm/form
    HoursVar = HoursVar + (Range("S10").Value * MixFiles)
    MixSetup = (Range("S4").Value) + (Range("S5").Value) + (Range("S10").Value)
    
    'Add time for hydrating oligos
    HoursVar = HoursVar + (Range("S6").Value * (R * N))
    'Add time for vortexing oligos
    HoursVar = HoursVar + (Range("S7").Value * (R * N))
    'Add time for creating top level ref
    HoursVar = HoursVar + (Range("S8").Value * MixFiles)
    'Add time for transferring oligos
    HoursVar = HoursVar + (Range("S9").Value * (R * N))
        
    'Add time for aliquots
    'Label
    HoursVar = HoursVar + (Range("M14").Value / 96 * A)
    'Create aliquots
    HoursVar = HoursVar + (Range("M15").Value / 96 * A)
    'Grav test
    If A > 10 Then
        HoursVar = HoursVar + (Range("M16").Value * R)
    End If
    'Bulk Hand OD
    If A > 10 Then
        HoursVar = HoursVar + (Range("M22").Value * R)
    Else
        HoursVar = HoursVar
    End If
    'Cap
    HoursVar = HoursVar + (Range("M17").Value / 96 * A)

    Mix = HoursVar
End Function
Function FormType()
    'Add time to formulation based on formulation type
    FormulationType = Range("C8").Value
    Refs = Range("C9").Value
    numMix = Range("E9").Value
    Aliquots = Range("G9").Value
    Partial = Range("C21").Value
    If Partial = "Yes" Then
        Refs = Refs / Range("M35").Value
        Aliquots = Aliquots / Range("M35").Value
    End If
    If Aliquots = 0 Then
        Aliquots = 1
    End If
    If FormulationType = "Single" Then
        Call SingleOligo(Refs, Aliquots)
    ElseIf FormulationType = "Duplex" Then
        Call Duplex(Refs, numMix, Aliquots)
    ElseIf FormulationType = "Mix" Then
        Call Mix(Refs, numMix, Aliquots)
    End If
    
    FormType = HoursVar
End Function
Function FormType1()
    'Add time to formulation based on the second formulation type
    FormulationType1 = Range("C10").Value
    Refs1 = Range("C11").Value
    numMix1 = Range("E11").Value
    Aliquots1 = Range("G11").Value
    If FormulationType1 = "Single" Then
        Call SingleOligo(Refs1, Aliquots1)
    ElseIf FormulationType1 = "Duplex" Then
        Call Duplex(Refs1, numMix1, Aliquots1)
    ElseIf FormulationType1 = "Mix" Then
        Call Mix(Refs1, numMix1, Aliquots1)
    End If
    FormType1 = HoursVar
    
    'Create a variable that is a total number of refs being processed
    SumRefs = Refs + Refs1
    Tubes = SumRefs * (numMix + numMix1)
    SumAliquots = Aliquots + Aliquots1
End Function
Function NonSynthTime()
    NonSynth = Range("C12").Value
    'Add time to formulation for using non-synth oligos
    If NonSynth > 0 Then
        If HandOD = "No" Then
            HoursVar = HoursVar + (Range("P14").Value * NonSynth) + (Range("P15").Value * NonSynth) + HandODTime(NonSynth)
        Else
            HoursVar = HoursVar + (Range("P14").Value * NonSynth) + (Range("P15").Value * NonSynth)
        End If
    Else
        HoursVar = HoursVar
    End If
    
    'Setting the ending of each function to the value of HoursVar allows us to see if each function is getting pulled correctly for hours()
    NonSynthTime = HoursVar
End Function
Function HandODTime(R As Integer)
    'Add the time it takes to Hand OD oligos
    HandOD = Range("C14").Value
    'Find how many dropsense chips will be needed
    Dim Chips As Integer
    Chips = Application.WorksheetFunction.RoundUp((R / 15), 0)
    
    If HandOD = "Yes" Then
            HoursVar = HoursVar + (Range("M20").Value) + (Range("M21").Value * Chips)
    ElseIf HandOD = "Yes - Cary" Then
        HoursVar = HoursVar + (Range("M20").Value) + (Range("M21").Value * R) + (Range("M23").Value) * R
    Else
        HoursVar = HoursVar
    End If
    
    'Combining HandOD and BECC since they're related
    BECC = Range("C15").Value
    If BECC = "Yes" Then
        If HandOD = "Yes" Then
            HoursVar = HoursVar + (Range("M20").Value) + (Range("M21").Value * Chips)
        ElseIf HandOD = "Yes - Cary" Then
            HoursVar = HoursVar + ((Range("M20").Value) + (Range("M23").Value * R))
        End If
    'Add time for BECC labels and spec sheets
    HoursVar = HoursVar + ((Range("P20").Value + Range("P21").Value) * R)
    ElseIf BECC = "No" Then
        HoursVar = HoursVar
    End If
    
    HandODTime = HoursVar
End Function
Function LabelTime()
    Labels = Range("C16").Value
    If Labels = "Custom" Then
        HoursVar = HoursVar + (Range("S21").Value * SumRefs)
    ElseIf Labels = "Standard" Then
        HoursVar = HoursVar + (Range("S20").Value * SumRefs)
    End If
    
    LabelTime = HoursVar
End Function
Function SpecSheetTime()
    SpecSheets = Range("C17").Value
    If SpecSheets = "Custom" Then
        HoursVar = HoursVar + (Range("M28").Value * SumRefs)
    ElseIf SpecSheets = "Standard" Then
        HoursVar = HoursVar + (Range("M27").Value * SumRefs)
    ElseIf SpecSheets = "None" Then
        HoursVar = HoursVar
    End If
    
    SpecSheetTime = HoursVar
End Function
Function TracesTime()
    Traces = Range("C19").Value
    If (Traces = "ESI" Or Traces = "CE" Or Traces = "RP-HPLC") Then
        HoursVar = HoursVar + (Range("P28").Value)
    ElseIf Traces = "RNase/DNase" Then
        HoursVar = HoursVar + (Range("P27").Value)
    ElseIf Traces = "Multiple or Other" Then
        HoursVar = HoursVar + (Range("P29").Value)
    ElseIf Traces = "Multiple incl. RNase/DNase" Then
        HoursVar = HoursVar + (Range("P28").Value) + (Range("P27").Value) + ((Range("P29").Value) * SumRefs)
    ElseIf Traces = "None" Then
        HoursVar = HoursVar
    Else
        'Was having issues if this was left out
        HoursVar = HoursVar
    End If
    
    TracesTime = HoursVar
End Function
Function SpecialPack()
    Dim PackagingCount As Double
    Packaging = Range("C20").Value
    'Should aliquots or ref#s be used to count tubes
    If SumAliquots > SumRefs Then
        PackagingCount = SumAliquots
    Else
        PackagingCount = SumRefs
    End If
    
    If Packaging = ">100 Aliquots" Then
        HoursVar = HoursVar + ((Range("S27").Value) * (PackagingCount / 96))
    ElseIf Packaging = "Bullet Box" Then
        HoursVar = HoursVar + ((Range("S28").Value) * (PackagingCount / 100))
    ElseIf Packaging = "Individually Bagged" Then
        HoursVar = HoursVar + ((Range("S29").Value) * PackagingCount)
    ElseIf Packaging = "Individually Bagged w/ SAP Label" Then
        HoursVar = HoursVar + ((Range("S30").Value) * PackagingCount)
    ElseIf Packaging = "Foil Bags" Then
        HoursVar = HoursVar + ((Range("S31").Value) * PackagingCount)
    ElseIf Packaging = "None" Then
        HoursVar = HoursVar
    End If
    'Add outgoing packaging
    HoursVar = HoursVar + (Range("S32").Value)
    
    'Add shipping condition
    ShipCondit = Range("C13").Value
    If ShipCondit = "Dry" Then
        'Add time for loading speedvac/IVAP if shipping dry
        HoursVar = HoursVar + (Range("S14").Value)
    ElseIf ShipCondit = "Wet" Then
        'Add time for longer preform if shipping wet
        HoursVar = HoursVar + Range("S15").Value
    Else
        HoursVar = HoursVar
    End If
    
    SpecialPack = HoursVar
End Function
Function ShipPartial()
    Partial = Range("C21").Value
    If Partial = "Yes" Then
        HoursVar = HoursVar * Range("M35").Value
    ElseIf Partial = "No" Then
        HoursVar = HoursVar
    End If
    
    ShipPartial = HoursVar
End Function
Function AdapterReview()
    Adapter = Range("C22").Value
    If Adapter = "Yes" Then
        HoursVar = HoursVar + (Range("P35").Value) + (Range("P36").Value * SumRefs)
    ElseIf Adapter = "No" Then
        HoursVar = HoursVar
    End If
    
    AdapterReview = HoursVar
End Function
Function ManagerReview()
    ManReview = Range("C18").Value
    If ManReview = "CCM" Then
        HoursVar = HoursVar + Range("S35").Value + (Range("S36").Value * SumRefs)
    ElseIf ManReview = "Qiagen" Then
        HoursVar = HoursVar + Range("S35").Value + (Range("S36").Value * SumRefs) + Range("S36").Value
    ElseIf ManReview = "Bio-Rad" Then
        HoursVar = HoursVar + Range("S35").Value + (Range("S36").Value * SumRefs) + Range("S37").Value
    ElseIf ManReview = "No" Then
        HoursVar = HoursVar
    End If
    
    ManagerReview = HoursVar
End Function

Function FormulatePartial()
    Dim formPartCount As Integer
    Partial = Range("C21").Value
    If ((HoursVar > 360) And (Partial = "No")) Then
        formPartCount = Round(HoursVar / 180, 0)
        If FormulationType = "Single" Then
            HoursVar = HoursVar + (SingleSetup * formPartCount)
        ElseIf FormulationType = "Duplex" Then
            HoursVar = HoursVar + (DuplexSetup * formPartCount)
        ElseIf FormulationType = "Mix" Then
            HoursVar = HoursVar + (MixSetup * formPartCount)
        End If
    End If
    
    FormulatePartial = HoursVar

End Function
Function Hours()
    HoursVar = 0#
    HoursVarRound = 0
    
    File
    FormType
    HoursVarArray(0) = HoursVar
    FormType1
    HoursVarArray(1) = HoursVar
    NonSynthTime
    HoursVarArray(2) = HoursVar
    If FormulationType = "Single" Then
        HandODTime (SumRefs)
    Else
        HandODTime (Tubes)
    End If
    HoursVarArray(3) = HoursVar
    LabelTime
    HoursVarArray(4) = HoursVar
    SpecSheetTime
    HoursVarArray(5) = HoursVar
    TracesTime
    HoursVarArray(6) = HoursVar
    SpecialPack
    HoursVarArray(7) = HoursVar
    ShipPartial
    HoursVarArray(8) = HoursVar
    AdapterReview
    HoursVarArray(9) = HoursVar
    ManagerReview
    HoursVarArray(10) = HoursVar
    'FormulatePartial
    HoursVarArray(11) = HoursVar

    'Round up hours to an integer
    HoursVarRound = Application.WorksheetFunction.RoundUp((HoursVar / 60), 0)
    'Output the final rounded time
    Hours = HoursVarRound
End Function

Sub Intervals()

    'Activate Input sheet
    Worksheets("Input").Activate
    
    Dim forcount, cells, cellstring
        For cells = 35 To 46
            forcount = cells - 35
            cellstring = "C" + CStr(cells)
            ActiveSheet.Range(cellstring).Value = HoursVarArray(forcount)
        Next cells

End Sub

