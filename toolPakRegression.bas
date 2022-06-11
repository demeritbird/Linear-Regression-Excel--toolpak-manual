Attribute VB_Name = "toolPakRegression"
Option Explicit




Sub toolpak_alldata()
    DoEvents
    
    If Main.Range("D11").Value = "" Then
        GoTo EmptyError
    End If

    On Error GoTo ToolPakError

    '' Initialising
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Loading..."
    End With
          
    Dim predictors As Byte
    Dim lastrow As Integer
    predictors = HiddenData.Range("D24").Value
    lastrow = HiddenData.Range("D23").Value
           

    '' Carries Pedictor Values to Formulas! Sheet
    HiddenData.Activate
    HiddenData.Range("D27:D30").Value = Main.Range("D11:D14").Value
    DoEvents
    
    
    '' Runs Analysis ToolPak through Normalised Values of ALLData!
    With Application
        .SendKeys "{Enter}"
        .Run "ATPVBAEN.XLAM!Regress", HiddenData.Range("O" & 2, "O" & lastrow), _
        HiddenData.Range("K" & 2, Cells(lastrow, predictors + 10)), _
        False, True, , HiddenData.Range("$T$3"), False, , False, False, _
        HiddenData.Range("$A$1"), , False
    End With
    DoEvents
    
    
    '' Returns End-User to page / confirmation message
    Main.Activate
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = ""
    End With
    
    MsgBox "Results:" & vbCrLf & "Adj. r^2 value = " & Main.Range("P10").Value, , "ToolPak Regression for ALLData finished!"
    
    Exit Sub
    
    
EmptyError:
    MsgBox "Please input at least 1 predictor!", , "Error!"
Exit Sub

ToolPakError:
DoEvents
    MsgBox "An error ocurred!" & vbCrLf & "Please check if:" & vbCrLf & "  - Macros are enabled" & vbCrLf & "  - Analysis Toolpak - VBA is installed.", , "Error?"
    Main.Activate
End Sub



Sub splitdata_w_toolpak_trainingdata()
    DoEvents

    On Error GoTo ToolPakError
    
    '' Initialising
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .StatusBar = "Splitting Data..."
    End With
    
    '' Randomises and Splits Data into TrainingSet & TestSet
    SplitData.Range("C7").Formula = "training"
    
    DataSort.AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("S1:S6651"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With DataSort.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    DoEvents
    DataSort.AutoFilter.Sort.SortFields.Clear
    
    SplitData.Range("C7").Formula = ""
    DoEvents
    
    
    '' Regression Initialising
    Application.StatusBar = "Loading..."
    
    Dim predictors As Byte
    Dim lastrow As Integer
    
    predictors = SplitData.Range("D24").Value
    lastrow = SplitData.Range("D23").Value
    
    SplitData.Activate
    

    '' Run Analysis ToolPak through Normalised Values of TrainingSet!
    With Application
        .SendKeys "{Enter}"
        .Run "ATPVBAEN.XLAM!Regress", SplitData.Range("O" & 2, "O" & lastrow), _
        SplitData.Range("K" & 2, Cells(lastrow, predictors + 10)), _
        False, True, , SplitData.Range("$T$3"), False, , False, False, _
        SplitData.Range("$A$1"), , False
    End With
    DoEvents
    
    Main.Activate
    
    
    '' Returns End-User to page / confirmation message
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = ""
    End With

    MsgBox "You may find:" & vbCrLf & "- Training Data in Sheet ""TrainingSet""" & vbCrLf & _
                                        "- Test Data in Sheet ""TestSet""" & vbCrLf & _
                                        "- The relevant formulas and regression in Sheet ""Formulas (SplitData)""" _
            , , "Splitting of Data Done!"
    
    Exit Sub
            
ToolPakError:
DoEvents
    MsgBox "An error ocurred!" & vbCrLf & "Please check if:" & vbCrLf & "  - Macros are enabled" & vbCrLf & "  - Analysis Toolpak - VBA is installed.", , "Error?"
    Main.Activate
End Sub
