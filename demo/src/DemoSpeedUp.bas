Attribute VB_Name = "DemoSpeedUp"

Option Explicit

Private Const sheetName As String = "DemoSpeedUp"

Private Timer As StopWatch
Private SpeedUp As XLSpeedUp


'Illustration of SpeedUp class for Excel.
'(Performance increase by factor of ~10 over default application settings.)
Public Sub Normal_vs_SpeedUp_Example()

    Const NoOfIterations As Long = 500
    
    ClassLoader loadTimer:=True
    
    Dim wkb As Workbook
    Set wkb = ThisWorkbook
    
    'Create output worksheet if it does not exist.
    If Not SheetExists(wkb, sheetName) Then
        Dim wks As Worksheet
        Set wks = wkb.Sheets.Add( _
                After:=wkb.Sheets(wkb.Sheets.Count))
        wks.Name = sheetName
    Else
        Set wks = wkb.Worksheets(sheetName)
    End If
    
    Dim ElapsedSecondsNormal As Double
    ElapsedSecondsNormal = NormalExample(wks, NoOfIterations)
    
    Dim ElapsedSecondsSpeedUp As Double
    ElapsedSecondsSpeedUp = SpeedUpExample(wks, NoOfIterations)
    
    'Performance benefit
    Dim Msg As String
    Msg = "SpeedUp increases performance by a factor of ~" & _
            Round(ElapsedSecondsNormal / ElapsedSecondsSpeedUp, 0) & "."
    Debug.Print vbCrLf & "** " & Msg
    MsgBox Msg & vbCrLf & vbCrLf & "(See Immediate window for additional results.)", _
            vbInformation + vbOKOnly, "SpeedUp Results"
    
End Sub


'Without performance settings.
Private Function NormalExample( _
    ByVal wks As Worksheet, _
    ByVal NoOfIterations As Long _
) As Double
    
    With Timer
        .StartTimer
        DummyProcedure wks, NoOfIterations
        .StopTimer
        .PrintPerformance "Without performance settings."
        NormalExample = .GetSecondsElapsed
    End With

End Function


'With performance settings
Private Function SpeedUpExample( _
    ByVal wks As Worksheet, _
    ByVal NoOfIterations As Long _
) As Double

    With Timer
        .Reset
        SpeedUp.TurnOn
        DummyProcedure wks, NoOfIterations
        SpeedUp.TurnOff
        .StopTimer
        .PrintPerformance "With performance settings (SpeedUp ON)."
        SpeedUpExample = .GetSecondsElapsed
    End With

End Function

'You may need to 'Reset' settings if you cancel executing code in the middle of SpeedUp.
'Because of this, you may like to use 'Reset' (instead of TurnOff) at the end of a major routine.
'This will ensure that Excel functions normally (using default application settings).
Public Sub SpeedUp_Reset_Example()

    ClassLoader
    Debug.Print "Before Speed-Up Reset:"
    Debug.Print "|> " & SpeedUp.ToString
    
    SpeedUp.Reset
    Debug.Print "After Speed-Up Reset:"
    Debug.Print "|> " & SpeedUp.ToString
    
End Sub


'Instantiate objects if they do not exist.
Private Sub ClassLoader(Optional ByVal loadTimer As Boolean = False)

    If Timer Is Nothing And loadTimer Then
        Set Timer = New StopWatch
    End If
    If SpeedUp Is Nothing Then
        Set SpeedUp = New XLSpeedUp
    End If
    
End Sub


'Returns TRUE if the specified sheet name exists in the active workbook,
'or FALSE if it does not exist.
Private Function SheetExists( _
    ByVal wkb As Workbook, _
    ByVal sheetName As String _
        ) As Boolean

    'NOTE: Sheets collection includes both Charts and Worksheets.
    On Error Resume Next
    Dim wks As Worksheet
    Set wks = wkb.Sheets(sheetName)
    SheetExists = (Not wks Is Nothing)
    
End Function


'this is a procedure which is intentionally made slow.
'Please folks, don't do that at home!!
Private Sub DummyProcedure( _
    ByVal wks As Worksheet, _
    Optional ByVal rowSize As Long = 1000, _
    Optional ByVal columnSize As Long = 3 _
)

    Dim rng As Range
    Set rng = wks.Range("A1").Resize(rowSize, columnSize)

    Application.WindowState = xlMaximized
    wks.Activate
    
    ' Start with clean slate.
    wks.UsedRange.ClearContents

    Dim SingleCell As Range
    For Each SingleCell In rng
        SingleCell.Select ' for illustration only i.e. avoid Select for faster code.
        SingleCell.Formula = "=Row()*Column()"
        DoEvents
    Next

End Sub
