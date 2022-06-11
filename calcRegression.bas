Attribute VB_Name = "calcRegression"
Option Explicit

'''' Incomplete


Sub Multiple_Linear_Regression()

    Dim size As Integer
    Dim X As Variant
    Dim B As Variant
    Dim Y As Variant
    
    Dim no_iterations As Double
    Dim learn_rate As Double
    Dim cost As Double
    
    Dim predictors As Byte
    Dim lastrow As Byte
    
    
    Application.ScreenUpdating = False
    
    
    predictors = Range("No_Predictors").Value
    lastrow = Range("lastrow").Value
    
    Main.Range("C" & 2, "C" & 6).ClearContents
    
    
    
    ActiveWorkbook.Sheets("HiddenData").Activate
    
    Call Populate_HiddenData
    
    HiddenData.Range("J" & 2, Cells(lastrow, predictors + 9)).Value = HiddenData.Range("Q2:T65").Value
       

    X = HiddenData.Range("I" & 2, Cells(lastrow, predictors + 9))
    B = Main.Range("B" & 2, "B" & (predictors + 2))
    Y = HiddenData.Range("N" & 2, "N" & lastrow)
    ActiveWorkbook.Sheets("Main").Activate
    
    no_iterations = Range("No_Iterations").Value
    learn_rate = 1 / no_iterations
    

    Call Gradient_Descent(X, Y, B, learn_rate, no_iterations)
    
    MsgBox "Manual Regression Done:" & vbCrLf & "completing this linear regression makes you wonder " & vbCrLf & "how much time you've wasted coding this"
    Application.ScreenUpdating = True

End Sub

Function Gradient_Descent(X, Y, B, alpha, no_iterations) As Double

    Dim iteration As Long
    Dim y_hat As Variant
    Dim gradient As Variant
    Dim X_T As Variant
    Dim deriative_error_range As Variant
    
    Dim cost As Double
    
    For iteration = 1 To no_iterations
        ' compute the derivate with current parameters
        y_hat = WorksheetFunction.MMult(X, B)
        deriative_error_range = ConvertToRange(Subtract_Matrix(y_hat, Y))
        X_T = WorksheetFunction.Transpose(X)
        
    
        ' update the parameters
        
        gradient = WorksheetFunction.MMult(X_T, deriative_error_range)
        
        B = Update_Coefficient(B, gradient, alpha)
        

        
        
        ' compute cost at every iteration of beta
        cost = CostFunction(X, Y, B)
        HiddenData.Range("AD" & iteration + 1) = cost
    Next
    
    Gradient_Descent = cost
     
     
End Function
Function Update_Coefficient(B, gradient, alpha) As Variant

    Dim predictors As Integer
        predictors = Range("No_Predictors").Value

    Dim idx As Long
    Dim start_idx As Long
    Dim end_idx As Long
    
    start_idx = LBound(B)
    end_idx = UBound(B)
    
        
    
    For idx = start_idx To end_idx
        Main.Range("C" & idx + 1).Value = B(idx, 1) - gradient(idx, 1) * alpha
    Next
    
    Update_Coefficient = Main.Range("C" & 2, "C" & (predictors + 2)).Value

End Function


Function CostFunction(X As Variant, Y As Variant, B As Variant) As Double

    Dim error As Variant
    Dim cost As Double
    Dim size As Integer
    
    size = UBound(X)
    
    error = WorksheetFunction.MMult(X, B) ' "- y" is done below
    
    ' computing matrix of error_square
    ' cost = ErrorSquared(error, y) / (2 * size)
    

    cost = MeanSquaredError(error, Y) / (2 * size)
    CostFunction = cost
    
    
End Function

Function MeanSquaredError(target, pred) As Double
    Dim idx As Integer
    Dim sum As Double
    Dim size As Double
    
    size = UBound(target)
    
    sum = 0
    For idx = LBound(target) To UBound(target) - 1
        sum = sum + (pred(idx, 1) - target(idx, 1)) ^ 2
    Next
    MeanSquaredError = sum / size

End Function

Function Subtract_Matrix(y_hat, Y) As Variant
    Dim idx As Integer
    Dim new_m As Variant
    Dim size As Long
    
    size = UBound(y_hat)
    
    ReDim new_m(1 To size)
    For idx = LBound(y_hat) To UBound(y_hat)
        new_m(idx) = y_hat(idx, 1) - Y(idx, 1)
    Next
    
    Subtract_Matrix = new_m
    
End Function

Function ConvertToRange(arr As Variant) As Variant
    Dim idx As Long
    Dim start_idx As Long
    Dim end_idx As Long
    
    start_idx = LBound(arr)
    end_idx = UBound(arr)
    
    For idx = start_idx To end_idx
        HiddenData.Range("AE" & idx + 1).Value = arr(idx)
    Next
    
    ConvertToRange = HiddenData.Range("AE" & start_idx + 1, "AE" & end_idx + 1).Value

End Function

