Option Explicit

Public OptimisationCollection As New Collection
Public MediumCount As Integer
Public MonthCount As Integer
Public BudgetStep As Integer
Public Budget As Long
Public AdStock As Long
Public BudgetStepCount As Integer
Public DemandSeasonalityArray(1 To 24) As Double
Public DecayArray() As Double
Public MediaCostSeasonalityArray() As Double
Public OptimisedInputValueArray() As Double
Public MediaArray() As String
Public MinsValueArray() ' [Month] / [Medium]
Public MaxValueArray() ' [Month] / [Medium]
Public TresholdValueArray() ' [Month] / [Medium]
Public ArrayValue() ' [Month]/[budget step] / [medium]
Public ArrayMediaOptimization() '1 to medium, 0 to budget,  ], 0 to 12 [0: total revenue from total budget for given medium, from 1 to 12: budget allocated in given month] ,
Public MaxRange As Integer
Public AdstockMultiplier As Double
Public CandidateMediaArray() ' 1 - YES | 0 - NO
Public OptimizationFunctionType()
Public BudgetsBeyond12Months()
Public AdstockFromBefore12Months()
Public AdstockStepValue As Long
Public BudgetStepValue As Long
Public SumOfMins As Double

Public Sub Main()
'

Dim StartTime As Double
Dim SecondsElapsed As Double
Dim MediumCounter As Integer
    
    StartTime = Timer

Debug.Print Now

'set ranges for data reading
MediumCount = OptimisationSheet.Range("OptimisationMediaName").Count
MonthCount = 24

'read min/max/threshold as well as 'sum of mins' variable important from the perspective of setting total budget to be optimized
FillMinMaxTresholdValueArray

'fill remainging public variables (including the total budget adjusted for the sum of mins)
FillPublicVariable

FillMediaArray
FillMediaCostSeasonalityArray
FillDemandSeasonalityArray
FillOptimisedInputValueArray
FillDecayArray

FillCandidateMediaArray
FillBudgetsBeyond12Months
FillAdstockFromBefore12Months
FillOptimizationFunctionType

'prepare cross-media optimization template
ReDim ArrayMediaOptimization(1 To MediumCount, 0 To MaxRange, 0 To 12)

'loop over media and fill ArrayMediaOptimization table
For MediumCounter = 1 To MediumCount
    
    'check if given medium is a candidate medium
    If CandidateMediaArray(MediumCounter) = 1 Then
        FillMediaOptimizationArray (MediumCounter)
    Else
        FillMediaOptimizationArray_ForNoCandidateMedium (MediumCounter)
    End If

Next MediumCounter

'optimize across media and print results
FillCrossMediaOptimizationArray

Debug.Print Now

  SecondsElapsed = Round(Timer - StartTime, 2)
  MsgBox "Optimisation complete in " & SecondsElapsed & " seconds", vbInformation

End Sub

Public Function FillOptimizationFunctionType()

Dim MediumCounter As Integer

'prepare array
ReDim OptimizationFunctionType(1 To MediumCount)

'read data
For MediumCounter = 1 To MediumCount

    OptimizationFunctionType(MediumCounter) = Hidden_Settings.Cells(1 + MediumCounter, 5).Value
    
Next MediumCounter

End Function

Public Function FillBudgetsBeyond12Months()

Dim MediumCounter As Integer
Dim MonthCounter As Integer

'prepare array of the right dimensions
ReDim BudgetsBeyond12Months(1 To MediumCount, 1 To 12)

'read data
For MediumCounter = 1 To MediumCount
    For MonthCounter = 1 To 12
    
        BudgetsBeyond12Months(MediumCounter, MonthCounter) = Hidden_Settings.Cells(33 + MediumCounter, 1 + MonthCounter).Value
            
    Next MonthCounter
Next MediumCounter

End Function

Public Function FillAdstockFromBefore12Months()

Dim MediumCounter As Integer

'prepare array of the right dimensions
ReDim AdstockFromBefore12Months(1 To MediumCount)

For MediumCounter = 1 To MediumCount
    
    AdstockFromBefore12Months(MediumCounter) = Hidden_Settings.Cells(58 + MediumCounter, 2).Value
    
Next MediumCounter

End Function

Public Function FillMediaOptimizationArray_ForNoCandidateMedium(MediumIndex As Integer)

Dim BudgetStep As Long

'save 0 for 0 allocation
ArrayMediaOptimization(MediumIndex, 0, 0) = 0

'loop over budgets and save penalty payoff value
For BudgetStep = 1 To MaxRange

   ArrayMediaOptimization(MediumIndex, BudgetStep, 0) = -999999999

Next BudgetStep

End Function

Public Function FillPublicVariable()
'
    'read total budget to be optimized
    Budget = OptimisationSheet.Range("Budget").Cells(1, 1).Value
            
    'set the adstock range - intentionally based on total budget and not the one adjusted for the sum of mins
    AdstockMultiplier = 0.25
    AdStock = Budget * AdstockMultiplier
    
    'adjust the budget for the sum of mins <- this will be the budget which we'll optimize
    Budget = Budget - SumOfMins
    
    'set the range and thus decide on the dollar value of the optimization step
    MaxRange = 200

    'calcualte value of steps - separately for adstock and budget (spend)
    AdstockStepValue = AdStock / MaxRange
    BudgetStepValue = Budget / MaxRange
'
End Function

Public Function FillCrossMediaOptimizationArray()

Dim MediumIndex As Integer, BudgetStep As Long, CurrentBudget As Long, CurrentBestBudget As Long, CurrentBestPayoff As Long, CurrentAnalyzedPayoff_FromCurrentMedium As Long, CurrentAnalyzedPayoff_FromPreviousMedia As Long, CurrentAnalyzedPayoff_Total As Long
Dim CrossMediaOptimization_Temp()
Dim BudgetLeft As Long
Dim BudgetLeftMonthly As Long
Dim MonthIndex As Integer
Dim CurrentMonthMin As Double

'initialize optimization table
ReDim CrossMediaOptimization_Temp(0 To MaxRange, 1 To MediumCount, 1 To 2) 'last dimension: 1 for budget, 2 for payoff

'initialize first medium
For CurrentBudget = 0 To MaxRange

    CrossMediaOptimization_Temp(CurrentBudget, 1, 1) = CurrentBudget
    CrossMediaOptimization_Temp(CurrentBudget, 1, 2) = ArrayMediaOptimization(1, CurrentBudget, 0)
    
Next CurrentBudget
        
'loop over the rest of media
For MediumIndex = 2 To MediumCount
    For BudgetStep = 0 To MaxRange
    
    'clear current best
    CurrentBestBudget = 0
    CurrentBestPayoff = 0
 
    'loop over possible budget splits
    For CurrentBudget = 0 To BudgetStep
    
        'calculate current option
        CurrentAnalyzedPayoff_FromCurrentMedium = ArrayMediaOptimization(MediumIndex, CurrentBudget, 0)
        CurrentAnalyzedPayoff_FromPreviousMedia = CrossMediaOptimization_Temp(BudgetStep - CurrentBudget, MediumIndex - 1, 2)
        
        CurrentAnalyzedPayoff_Total = CurrentAnalyzedPayoff_FromCurrentMedium + CurrentAnalyzedPayoff_FromPreviousMedia
        
        'compare with current best and update if needed
        If CurrentAnalyzedPayoff_Total > CurrentBestPayoff Then
                
            CurrentBestPayoff = CurrentAnalyzedPayoff_Total
            CurrentBestBudget = CurrentBudget
                
        End If
        
    Next CurrentBudget
    
    'save current best
    CrossMediaOptimization_Temp(BudgetStep, MediumIndex, 1) = CurrentBestBudget
    CrossMediaOptimization_Temp(BudgetStep, MediumIndex, 2) = CurrentBestPayoff
       
    Next BudgetStep
Next MediumIndex

'print results
Worksheets("output").Cells(29, 1).Value = MaxRange * BudgetStepValue + SumOfMins '<- total budget for our optimization
Worksheets("output").Cells(29, 2).Value = CrossMediaOptimization_Temp(MaxRange, MediumCount, 2) '<- total revenue from our optimization

BudgetLeft = MaxRange

For MediumIndex = MediumCount To 1 Step -1
    
    'print spend per given medium (in total)
    Worksheets("output").Cells(MediumIndex + 30, 1).Value = CrossMediaOptimization_Temp(BudgetLeft, MediumIndex, 1) * BudgetStepValue
    
    'print spends per months of given medium (and total revenue from this medium in first row)
    Worksheets("output").Cells(MediumIndex + 30, 2).Value = ArrayMediaOptimization(MediumIndex, CrossMediaOptimization_Temp(BudgetLeft, MediumIndex, 1), 0)
    
    For MonthIndex = 1 To 12
        
        'calculate current month min
        CurrentMonthMin = GetMinsValueArray(MediumIndex, MonthIndex)
        
        'print current month min plus the value added on top by the optimization
        Worksheets("output").Cells(MediumIndex + 30, MonthIndex + 2).Value = (CurrentMonthMin + ArrayMediaOptimization(MediumIndex, CrossMediaOptimization_Temp(BudgetLeft, MediumIndex, 1), MonthIndex)) * BudgetStepValue
        
    Next MonthIndex
    
    'update budget passed over
    BudgetLeft = BudgetLeft - CrossMediaOptimization_Temp(BudgetLeft, MediumIndex, 1)

Next MediumIndex

End Function

Public Function FillMediaOptimizationArray(MediumIndex As Integer)
'
Dim MonthIndex As Integer, BudgetStep As Long, CurrentBudget As Long, MaxValue As Double, AdstockStep As Long, CurrentBestBudget As Long, CurrentBestPayoff As Long, CurrentAnalyzedPayoff_FromCurrentMonth As Long, CurrentAnalyzedPayoff_FromFutureMonths As Long, CurrentAnalyzedPayoff_Total As Long
Dim ArrayMediaOptimization_Temp()
Dim AdstockFromPreviousMonths As Long, BudgetLeft As Long, LastMonthMin As Double
Dim BudgetStepAdjustedForMax As Long
Dim CurrentMonthMax As Double, CurrentMonthMin As Double, CurrentMonthThreshold As Double
Dim DecayFactor As Double, MonthlyCostSeasonalityFactor As Double
Dim AdstockTotalValueFromPreviousMonth As Double
Dim BudgetStepWithMinAdded As Double, CurrentBudgetWithMinAdded As Double
Dim PayoffCoefficient As Double, PayoffAlpha As Double, PayoffZeta As Double, PayoffEta As Double
Dim CurrentBudgetMax As Long
Dim OptimizationFunctionChoice As Integer

'initialize optimization table
ReDim ArrayMediaOptimization_Temp(0 To MaxRange, 0 To MaxRange, 0 To 12, 1 To 2) 'budget, adstocked budget, month, payoff & budget

'calculate decay
DecayFactor = DecayArray(MediumIndex)

'calculate payoff function coefficient, alpha, zeta, eta

PayoffCoefficient = OptimisedInputValueArray(MediumIndex, 1)
PayoffAlpha = OptimisedInputValueArray(MediumIndex, 2)
PayoffZeta = OptimisedInputValueArray(MediumIndex, 3)
PayoffEta = OptimisedInputValueArray(MediumIndex, 4)

'read the type of optimization function for current medium
OptimizationFunctionChoice = OptimizationFunctionType(MediumIndex)

'initialize last month (the 12th month)

MonthlyCostSeasonalityFactor = MediaCostSeasonalityArray(12, MediumIndex)

CurrentMonthMax = GetMaxsValueArray(MediumIndex, 12)
CurrentMonthMin = GetMinsValueArray(MediumIndex, 12)
CurrentMonthThreshold = GetTresholdsValueArray(MediumIndex, 12)

For BudgetStep = 0 To MaxRange
    
    'calculate budget step adjusted for min
    BudgetStepWithMinAdded = BudgetStep + CurrentMonthMin
    
    For AdstockStep = 0 To MaxRange '[!] 'pytanie czy nie potrzeba osobnego maksymalnego zasiegu dla adstocku
    
        'calculate payoff given budget level and adstock level from previous month (comment: adstock is the level of budget passed over from previous month (already multiplied by decay)
        If BudgetStepWithMinAdded > CurrentMonthMax _
                Or (BudgetStepWithMinAdded <> 0 And BudgetStepWithMinAdded < CurrentMonthThreshold) Then
                    
                    ArrayMediaOptimization_Temp(BudgetStep, AdstockStep, 12, 1) = BudgetStep
                    ArrayMediaOptimization_Temp(BudgetStep, AdstockStep, 12, 2) = -999999999
                
                Else
                    
                    'account for payoff from current month
                    CurrentAnalyzedPayoff_FromCurrentMonth = OptimizationValueWithAdStock(BudgetStepWithMinAdded * BudgetStepValue, AdstockStepValue * AdstockStep, 12, MediumIndex, DecayFactor, MonthlyCostSeasonalityFactor, PayoffCoefficient, PayoffAlpha, PayoffZeta, PayoffEta, OptimizationFunctionChoice)
                    
                    'account for months beyond our 12 optimized months
                    CurrentAnalyzedPayoff_FromFutureMonths = 0
                    AdstockTotalValueFromPreviousMonth = (BudgetStepWithMinAdded * BudgetStepValue + AdstockStep * AdstockStepValue) * DecayFactor
                                        
                    For MonthIndex = 1 To (MonthCount - 12)
                        
                        'calculate monthly cost seasonality
                        MonthlyCostSeasonalityFactor = MediaCostSeasonalityArray(12 + MonthIndex, MediumIndex)
                        
                        'calculate contribution of currently considered month to future months total sum
                        CurrentAnalyzedPayoff_FromFutureMonths = CurrentAnalyzedPayoff_FromFutureMonths + _
                                    OptimizationValueWithAdStock(BudgetsBeyond12Months(MediumIndex, MonthIndex), AdstockTotalValueFromPreviousMonth, MonthIndex, MediumIndex, DecayFactor, MonthlyCostSeasonalityFactor, PayoffCoefficient, PayoffAlpha, PayoffZeta, PayoffEta, OptimizationFunctionChoice)
                        
                        'calculate adstock step from current month for the purpose of next month calculation
                        AdstockTotalValueFromPreviousMonth = (BudgetsBeyond12Months(MediumIndex, MonthIndex) + AdstockTotalValueFromPreviousMonth) * DecayFactor
                        
                    Next MonthIndex
                    
                    CurrentAnalyzedPayoff_Total = CurrentAnalyzedPayoff_FromCurrentMonth + CurrentAnalyzedPayoff_FromFutureMonths
                    
                    ArrayMediaOptimization_Temp(BudgetStep, AdstockStep, 12, 1) = BudgetStep
                    ArrayMediaOptimization_Temp(BudgetStep, AdstockStep, 12, 2) = CurrentAnalyzedPayoff_Total
                End If
               
    Next AdstockStep
Next BudgetStep

'loop over previous months starting from the last but one (which as been initialized above)

For MonthIndex = 11 To 1 Step -1
    
    'calculate min/max/threshold for current month
    CurrentMonthMax = GetMaxsValueArray(MediumIndex, MonthIndex)
    CurrentMonthMin = GetMinsValueArray(MediumIndex, MonthIndex)
    CurrentMonthThreshold = GetTresholdsValueArray(MediumIndex, MonthIndex)
    
    'calculate media cost seasonality
    MonthlyCostSeasonalityFactor = MediaCostSeasonalityArray(MonthIndex, MediumIndex)
    
    For BudgetStep = 0 To MaxRange
        For AdstockStep = 0 To MaxRange
            
            'find optimal value for given set of initial parameters by analyzing all possible allocations into current month
            
            'clear current best option
            CurrentBestBudget = 0
            CurrentBestPayoff = -999999999 * 2
            
            'limit options taking into account monthly max (not to consider options beyond monthly max)
            If (BudgetStep + CurrentMonthMin) > CurrentMonthMax Then
            
                CurrentBudgetMax = Int(CurrentMonthMax - CurrentMonthMin) 'the int function is used to round down
            
            Else
            
                CurrentBudgetMax = BudgetStep
            
            End If
                        
            'loop over possible options
            For CurrentBudget = 0 To CurrentBudgetMax
                
                'adjust current budget for the min value
                CurrentBudgetWithMinAdded = CurrentBudget + CurrentMonthMin
                
                'check min/max/threshold compliance
                If CurrentBudgetWithMinAdded > CurrentMonthMax _
                Or (CurrentBudgetWithMinAdded <> 0 And CurrentBudgetWithMinAdded < CurrentMonthThreshold) Then
                    CurrentAnalyzedPayoff_Total = -999999999
                Else
                
                    'calculate payoff
                    CurrentAnalyzedPayoff_FromCurrentMonth = OptimizationValueWithAdStock(BudgetStepValue * CurrentBudgetWithMinAdded, AdstockStepValue * AdstockStep, MonthIndex, MediumIndex, DecayFactor, MonthlyCostSeasonalityFactor, PayoffCoefficient, PayoffAlpha, PayoffZeta, PayoffEta, OptimizationFunctionChoice)
                    CurrentAnalyzedPayoff_FromFutureMonths = ArrayMediaOptimization_Temp(BudgetStep - CurrentBudget, GetAdStockIndex(CurrentBudgetWithMinAdded, AdstockStep, MonthIndex, MediumIndex, DecayFactor, MonthlyCostSeasonalityFactor), MonthIndex + 1, 2)
                                        
                    CurrentAnalyzedPayoff_Total = CurrentAnalyzedPayoff_FromCurrentMonth + CurrentAnalyzedPayoff_FromFutureMonths
                End If
                
                'compare with current best
                If CurrentAnalyzedPayoff_Total > CurrentBestPayoff Then
                
                    'assign new best allocation
                    CurrentBestBudget = CurrentBudget
                    CurrentBestPayoff = CurrentAnalyzedPayoff_Total
                
                End If
            Next CurrentBudget
            
            'save best allocation to results table
            ArrayMediaOptimization_Temp(BudgetStep, AdstockStep, MonthIndex, 1) = CurrentBestBudget
            ArrayMediaOptimization_Temp(BudgetStep, AdstockStep, MonthIndex, 2) = CurrentBestPayoff
            
        Next AdstockStep
    Next BudgetStep
Next MonthIndex

'--- save this medium results to all media results table ---

'loop over possible budget allocations for currently considered medium
For CurrentBudget = 0 To MaxRange

    'memorize adstock to read the right results [!]
    AdstockFromPreviousMonths = GetAdStockIndex(0, AdstockFromBefore12Months(MediumIndex) / AdstockStepValue, 12, MediumIndex, DecayFactor, MediaCostSeasonalityArray(12, MediumIndex))
        
    'save total revenue
    ArrayMediaOptimization(MediumIndex, CurrentBudget, 0) = ArrayMediaOptimization_Temp(CurrentBudget, AdstockFromPreviousMonths, 1, 2)
    
    'save monthly allocations for the purpose of cross-media optimization
    BudgetLeft = CurrentBudget
    
    ArrayMediaOptimization(MediumIndex, CurrentBudget, 1) = ArrayMediaOptimization_Temp(CurrentBudget, AdstockFromPreviousMonths, 1, 1)
    BudgetLeft = BudgetLeft - ArrayMediaOptimization(MediumIndex, CurrentBudget, 1)
    
    For MonthIndex = 2 To 12
                        
        'calculate monthly cost seasonality
        MonthlyCostSeasonalityFactor = MediaCostSeasonalityArray(MonthIndex, MediumIndex)
        
        'calculate last month's min
        LastMonthMin = GetMinsValueArray(MediumIndex, MonthIndex - 1)
        
        'calculate value of adstock from previous month to be used in current month
        AdstockFromPreviousMonths = GetAdStockIndex(ArrayMediaOptimization(MediumIndex, CurrentBudget, MonthIndex - 1) + LastMonthMin, AdstockFromPreviousMonths, MonthIndex, MediumIndex, DecayFactor, MonthlyCostSeasonalityFactor)
        ArrayMediaOptimization(MediumIndex, CurrentBudget, MonthIndex) = ArrayMediaOptimization_Temp(BudgetLeft, AdstockFromPreviousMonths, MonthIndex, 1)
        
        'update budget left
        BudgetLeft = BudgetLeft - ArrayMediaOptimization(MediumIndex, CurrentBudget, MonthIndex)
        
    Next MonthIndex
        
Next CurrentBudget

End Function

Public Function GetMinsValueArray(MediumIndex As Integer, MonthIndex As Integer) As Double

If MinsValueArray(MediumIndex, MonthIndex) = "" Then
    GetMinsValueArray = 0
Else
    GetMinsValueArray = MinsValueArray(MediumIndex, MonthIndex) / BudgetStepValue
End If
End Function

Public Function GetMaxsValueArray(MediumIndex As Integer, MonthIndex As Integer) As Double

If MaxValueArray(MediumIndex, MonthIndex) = "" Then
    GetMaxsValueArray = 999999999
Else
    GetMaxsValueArray = MaxValueArray(MediumIndex, MonthIndex) / BudgetStepValue
End If
End Function

Public Function GetTresholdsValueArray(MediumIndex As Integer, MonthIndex As Integer) As Double

If TresholdValueArray(MediumIndex, MonthIndex) = "" Then
    GetTresholdsValueArray = 0
Else
    GetTresholdsValueArray = TresholdValueArray(MediumIndex, MonthIndex) / BudgetStepValue
End If
End Function

Public Function GetAdStockIndex(ByVal BudgetStep As Long, ByVal AdstockStep As Long, MonthIndex As Integer, MediumIndex As Integer, Decay As Double, ByVal MediaCostSeasonality As Double) As Long
'
    Dim BudgetTmp As Double
    
    BudgetTmp = BudgetStep * BudgetStepValue
        
    GetAdStockIndex = Round(((BudgetTmp / MediaCostSeasonality) * Decay) / AdstockStepValue + Decay * AdstockStep, 0)

    'limit to large adstocks
    If GetAdStockIndex > MaxRange Then
        GetAdStockIndex = MaxRange
    End If
    
End Function

Public Function OptimizationValueWithAdStock(ByVal Budget As Long, AdStock As Double, MonthIndex As Integer, MediumIndex As Integer, Decay As Double, MediaCostSeasonality As Double, Coeff As Double, Alpha As Double, Zeta As Double, Eta As Double, FunctionChoice As Integer) As Double
'

    Dim DemandSeasonality As Double
        
    'read demand seasonality
    DemandSeasonality = DemandSeasonalityArray(MonthIndex)
        
    'apply the right payoff function
    If FunctionChoice = 0 Then
        
        'power function
        OptimizationValueWithAdStock = ((AdStock + Budget / MediaCostSeasonality) ^ Alpha) * Coeff * DemandSeasonality
      
    Else
        
        'S-curve function
        OptimizationValueWithAdStock = (1 - Exp(Zeta * (AdStock + Budget / MediaCostSeasonality) ^ Eta)) * Coeff * DemandSeasonality
        
    End If
              
'
End Function

Public Function FillDecayArray()
'
    Dim MediumIndex As Integer
    ReDim DecayArray(1 To MediumCount)
    
    For MediumIndex = 1 To MediumCount
        DecayArray(MediumIndex) = OptimisationSheet.Range("Decay").Cells(MediumIndex, 1)
    Next MediumIndex
'
End Function

Public Function FillOptimisedInputValueArray()
'
    Dim MediumIndex As Integer, ParameterIndex As Integer
    
    ReDim OptimisedInputValueArray(1 To MediumCount, 1 To 4)
    
    For MediumIndex = 1 To MediumCount
        For ParameterIndex = 1 To 4 '[Coeff]/[Alpha]/[Zeta]/[Eta]
            OptimisedInputValueArray(MediumIndex, ParameterIndex) = OptimisationSheet.Range("OptimisedInputValue").Cells(MediumIndex, ParameterIndex)
        Next ParameterIndex
    Next MediumIndex
'
End Function

Public Function FillDemandSeasonalityArray()
'
  Dim MonthIndex As Integer
  For MonthIndex = 1 To MonthCount
    DemandSeasonalityArray(MonthIndex) = CalibrationSheet.Range("DemandSeasonality").Cells(1, MonthIndex)
  Next MonthIndex
'
End Function

Public Function FillMinMaxTresholdValueArray()
'
  ReDim MinsValueArray(0 To MediumCount, 0 To 12)
  ReDim MaxValueArray(0 To MediumCount, 0 To 12)
  ReDim TresholdValueArray(0 To MediumCount, 0 To 12)
  
  Dim MediumIndex As Integer, MonthIndex As Integer
  
  'reset sum of mins variable
  SumOfMins = 0
    
  For MediumIndex = 1 To MediumCount
        For MonthIndex = 1 To 12
            
            'read min
            MinsValueArray(MediumIndex, MonthIndex) = OptimisationSheet.Range("Mins").Cells(MediumIndex, MonthIndex)
            
            'update sum of mins variable
            SumOfMins = SumOfMins + MinsValueArray(MediumIndex, MonthIndex)
            
            'read max
            MaxValueArray(MediumIndex, MonthIndex) = OptimisationSheet.Range("Max").Cells(MediumIndex, MonthIndex)
            
            'read threshold
            TresholdValueArray(MediumIndex, MonthIndex) = OptimisationSheet.Range("Treshold").Cells(MediumIndex, MonthIndex)
        
        Next MonthIndex
  Next MediumIndex
'
End Function

Public Function FillMediaArray()
'
    'OptimisationMediaName
    ReDim MediaArray(0 To MediumCount)
    
    Dim MediumIndex As Integer
    For MediumIndex = 1 To MediumCount
         MediaArray(MediumIndex) = OptimisationSheet.Range("OptimisationMediaName").Cells(MediumIndex, 1)
    Next MediumIndex
    
'
End Function

Public Function FillCandidateMediaArray()

ReDim CandidateMediaArray(1 To MediumCount)

Dim MediumCounter As Integer

For MediumCounter = 1 To MediumCount

    CandidateMediaArray(MediumCounter) = Hidden_Settings.Cells(1 + MediumCounter, 2).Value

Next MediumCounter

End Function

Public Function FillMediaCostSeasonalityArray()
'
    ReDim MediaCostSeasonalityArray(0 To MonthCount, 0 To MediumCount)
    
    Dim MediumIndex As Integer, MonthIndex As Integer
  
    For MediumIndex = 1 To MediumCount
        For MonthIndex = 1 To MonthCount
            MediaCostSeasonalityArray(MonthIndex, MediumIndex) = CalibrationSheet.Range("MediaCostSeasonality").Cells(MediumIndex, MonthIndex)
        Next MonthIndex
    Next MediumIndex
'
End Function

Public Function GetMediumIndex(MediumName As String) As Integer
'
    Dim MediumIndex As Integer
    '
    For MediumIndex = 1 To MediumCount
        If MediaArray(MediumIndex) = MediumName Then
            GetMediumIndex = MediumIndex
            Exit Function
        End If
    Next MediumIndex
'
End Function


