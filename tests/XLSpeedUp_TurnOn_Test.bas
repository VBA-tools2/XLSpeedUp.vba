Attribute VB_Name = "XLSpeedUp_TurnOn_Test"
Attribute VB_Description = "Tests covering the XLSpeedUp class."

Option Explicit
Option Private Module

'@TestModule
'@Folder("XLSpeedUp.Tests")
'@ModuleDescription("Tests covering the XLSpeedUp class.")

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


Private sut As XLSpeedUp


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
#If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
#Else
        Set Assert = New Rubberduck.PermissiveAssertClass
#End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    Set sut = New XLSpeedUp
    sut.Reset
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    sut.TurnOff
    Set sut = Nothing
End Sub


Private Sub InvertApplicationSettings()
    With Application
        If .Calculation = xlCalculationAutomatic Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = xlCalculationAutomatic
        End If
        If .Cursor = xlDefault Then
            .Cursor = xlWait
        Else
            .Cursor = xlDefault
        End If
        .DisplayAlerts = Not .DisplayAlerts
        .EnableAnimations = Not .EnableAnimations
        If .EnableCancelKey = xlInterrupt Then
            .EnableCancelKey = xlErrorHandler
        Else
            .EnableCancelKey = xlInterrupt
        End If
        .ScreenUpdating = Not .ScreenUpdating
    End With
End Sub


'==============================================================================
'@TestMethod("TurnOn")
Private Sub TurnOn_WithDefaultSettings_ReturnsVarious()
    On Error GoTo TestFail
    
    sut.TurnOn
    
    With Application
        Assert.AreEqual XlCalculation.xlCalculationManual, .Calculation, "Calculation"
        Assert.AreEqual 1, sut.Count, "Count"
        Assert.AreEqual XlMousePointer.xlWait, .Cursor, "Cursor"
        Assert.IsFalse .DisplayAlerts, "DisplayAlerts"
        Assert.IsFalse .EnableAnimations, "EnableAnimations"
        Assert.AreEqual XlEnableCancelKey.xlErrorHandler, .EnableCancelKey, "EnableCancelKey"
        Assert.IsFalse .ScreenUpdating, "ScreenUpdating"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_WithoutDefaultSettings_ReturnsVarious()
    On Error GoTo TestFail
    
    InvertApplicationSettings
    With Application
        .Calculation = xlCalculationSemiautomatic
        .Cursor = xlIBeam
    End With
    
    sut.TurnOn
    
    With Application
        Assert.AreEqual XlCalculation.xlCalculationManual, .Calculation, "Calculation"
        Assert.AreEqual 1, sut.Count, "Count"
        Assert.AreEqual XlMousePointer.xlWait, .Cursor, "Cursor"
        Assert.IsFalse .DisplayAlerts, "DisplayAlerts"
        Assert.IsFalse .EnableAnimations, "EnableAnimations"
        Assert.AreEqual XlEnableCancelKey.xlErrorHandler, .EnableCancelKey, "EnableCancelKey"
        Assert.IsFalse .ScreenUpdating, "ScreenUpdating"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultHideDisplayPageBreaksNotGiven_ReturnsFalse()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.DisplayPageBreaks = True
    Next
    
    sut.TurnOn
    
    For Each wks In ActiveWorkbook.Worksheets
        Assert.IsFalse wks.DisplayPageBreaks
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultHideDisplayPageBreaksTrue_ReturnsFalse()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.DisplayPageBreaks = True
    Next
    
    sut.TurnOn hideDisplayPageBreaks:=True
    
    For Each wks In ActiveWorkbook.Worksheets
        Assert.IsFalse wks.DisplayPageBreaks
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultAllowEventsNotGiven_ReturnsFalse()
    On Error GoTo TestFail
    
    sut.TurnOn
    
    Assert.IsFalse Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultAllowEventsFalse_ReturnsFalse()
    On Error GoTo TestFail
    
    sut.TurnOn allowEvents:=False
    
    Assert.IsFalse Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultAllowEventsTrue_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.TurnOn allowEvents:=True
    
    Assert.IsTrue Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultStatusbarNotGiven_ReturnsDefaultString()
    On Error GoTo TestFail
    
    sut.TurnOn
    
    Assert.AreEqual "SpeedUp is on.", Application.StatusBar

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultStatusbarEmptyString_ReturnsDefaultString()
    On Error GoTo TestFail
    
    sut.TurnOn statusBarMessage:=vbNullString
    
    Assert.AreEqual "SpeedUp is on.", Application.StatusBar

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOn")
Private Sub TurnOn_DefaultStatusbarGivenString_ReturnsDefaultString()
    On Error GoTo TestFail
    
    sut.TurnOn statusBarMessage:="Just a test string"
    
    Assert.AreEqual "Just a test string", Application.StatusBar

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("MultipleTurnOnOff")
Private Sub TurnOn_MultipleTimesWithChangingSettings_ReturnsVarious()
    On Error GoTo TestFail
    
    'iteration 0
    InvertApplicationSettings
    With Application
        .Calculation = xlCalculationSemiautomatic
        .Cursor = xlIBeam
    End With
    
    'iteration 1
    sut.TurnOn
    
    With Application
        If XlCalculation.xlCalculationManual <> .Calculation Then Assert.Inconclusive
        If 1 <> sut.Count Then Assert.Inconclusive
        If XlMousePointer.xlWait <> .Cursor Then Assert.Inconclusive
        If .DisplayAlerts Then Assert.Inconclusive
        If .EnableAnimations Then Assert.Inconclusive
        If XlEnableCancelKey.xlErrorHandler <> .EnableCancelKey Then Assert.Inconclusive
        If .ScreenUpdating Then Assert.Inconclusive
    End With
    
    InvertApplicationSettings
    With Application
        .Calculation = xlCalculationSemiautomatic
        .Cursor = xlNorthwestArrow
        .EnableCancelKey = xlDisabled
    End With
    
    'iteration 2
    sut.TurnOn
    
    With Application
        Assert.AreEqual XlCalculation.xlCalculationSemiautomatic, .Calculation, "2 Calculation"
        Assert.AreEqual 2, sut.Count, "2 Count"
        Assert.AreEqual XlMousePointer.xlNorthwestArrow, .Cursor, "2 Cursor"
        Assert.IsTrue .DisplayAlerts, "2 DisplayAlerts"
        Assert.IsTrue .EnableAnimations, "2 EnableAnimations"
        Assert.AreEqual XlEnableCancelKey.xlDisabled, .EnableCancelKey, "2 EnableCancelKey"
        Assert.IsTrue .ScreenUpdating, "2 ScreenUpdating"
    End With
    
    'iteration 1
    sut.TurnOff
    
    With Application
        Assert.AreEqual XlCalculation.xlCalculationSemiautomatic, .Calculation, "1 Calculation"
        Assert.AreEqual 1, sut.Count, "1 Count"
        Assert.AreEqual XlMousePointer.xlNorthwestArrow, .Cursor, "1 Cursor"
        Assert.IsTrue .DisplayAlerts, "1 DisplayAlerts"
        Assert.IsTrue .EnableAnimations, "1 EnableAnimations"
        Assert.AreEqual XlEnableCancelKey.xlDisabled, .EnableCancelKey, "1 EnableCancelKey"
        Assert.IsTrue .ScreenUpdating, "1 ScreenUpdating"
    End With
    
    'iteration 0
    sut.TurnOff
    
    With Application
        Assert.AreEqual XlCalculation.xlCalculationSemiautomatic, .Calculation, "0 Calculation"
        Assert.AreEqual 0, sut.Count, "0 Count"
        Assert.AreEqual XlMousePointer.xlDefault, .Cursor, "0 Cursor"
        Assert.IsFalse .DisplayAlerts, "0 DisplayAlerts"
        Assert.IsFalse .EnableAnimations, "0 EnableAnimations"
        Assert.AreEqual XlEnableCancelKey.xlInterrupt, .EnableCancelKey, "0 EnableCancelKey"
        Assert.IsFalse .ScreenUpdating, "0 ScreenUpdating"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'TODO: finish test, then remove the double comment sign
'@TestMethod("MultipleTurnOnOff")
Private Sub TurnOn_MultipleTimesWithChangingOptionalSettings_ReturnsVarious()
    On Error GoTo TestFail
    
    'iteration 0
    With Application
        .EnableEvents = False
        .StatusBar = False
    End With
    
    With Application
        If .EnableEvents Then Assert.Inconclusive
        If .StatusBar <> False Then Assert.Inconclusive
    End With
    
    'iteration 1
    sut.TurnOn _
            hideDisplayPageBreaks:=False, _
            allowEvents:=True, _
            statusBarMessage:="First call"
    
    With Application
        Assert.IsTrue .EnableEvents, "1on EnableEvents"
        Assert.AreEqual "First call", .StatusBar, "1on StatusBar"
    End With
    
    'iteration 2
    sut.TurnOn _
            hideDisplayPageBreaks:=True, _
            allowEvents:=False, _
            statusBarMessage:="Second call"
    
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        Assert.IsFalse wks.DisplayPageBreaks, "2 DisplayPageBreaks"
    Next
    With Application
        Assert.IsFalse .EnableEvents, "2 EnableEvents"
        Assert.AreEqual "First call", .StatusBar, "2 StatusBar"
    End With
    
    'iteration 1
    sut.TurnOff
    
    For Each wks In ActiveWorkbook.Worksheets
        Assert.IsFalse wks.DisplayPageBreaks, "1off DisplayPageBreaks"
    Next
    With Application
        Assert.IsFalse .EnableEvents, "1off EnableEvents"
        Assert.AreEqual "First call", .StatusBar, "1off StatusBar"
    End With
    
    'iteration 0
    sut.TurnOff
    
    For Each wks In ActiveWorkbook.Worksheets
        Assert.IsFalse wks.DisplayPageBreaks, "0 DisplayPageBreaks"
    Next
    With Application
        Assert.IsFalse .EnableEvents, "0 EnableEvents"
        Assert.IsTrue (.StatusBar = False), "0 StatusBar"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
