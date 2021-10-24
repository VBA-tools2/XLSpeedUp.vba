Attribute VB_Name = "XLSpeedUp_TurnOff_Test"
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


'==============================================================================
'@TestMethod("TurnOff")
Private Sub TurnOff_WithDefaultSettings_ReturnsVarious()
    On Error GoTo TestFail
    
    With Application
        .Calculation = xlCalculationAutomatic
        .Cursor = xlDefault
        .DisplayAlerts = True
        .EnableAnimations = True
        .EnableCancelKey = xlInterrupt
        .ScreenUpdating = True
    End With
    sut.TurnOn
    
    sut.TurnOff
    
    With Application
        Assert.AreEqual XlCalculation.xlCalculationAutomatic, .Calculation, "Calculation"
        Assert.AreEqual 0, sut.Count, "Count"
        Assert.AreEqual XlMousePointer.xlDefault, .Cursor, "Cursor"
        Assert.IsTrue .DisplayAlerts, "DisplayAlerts"
        Assert.IsTrue .EnableAnimations, "EnableAnimations"
        Assert.AreEqual XlEnableCancelKey.xlInterrupt, .EnableCancelKey, "EnableCancelKey"
        Assert.IsTrue .ScreenUpdating, "ScreenUpdating"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_WithoutDefaultSettings_ReturnsVarious()
    On Error GoTo TestFail
    
    With Application
        .Calculation = xlCalculationManual
        .Cursor = xlWait
        .DisplayAlerts = False
        .EnableAnimations = False
        .EnableCancelKey = xlErrorHandler
        .ScreenUpdating = False
    End With
    sut.TurnOn
    
    sut.TurnOff
    
    With Application
        Assert.AreEqual XlCalculation.xlCalculationManual, .Calculation, "Calculation"
        Assert.AreEqual 0, sut.Count, "Count"
        Assert.AreEqual XlMousePointer.xlDefault, .Cursor, "Cursor"
        Assert.IsFalse .DisplayAlerts, "DisplayAlerts"
        Assert.IsFalse .EnableAnimations, "EnableAnimations"
        Assert.AreEqual XlEnableCancelKey.xlInterrupt, .EnableCancelKey, "EnableCancelKey"
        Assert.IsFalse .ScreenUpdating, "ScreenUpdating"
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultHideDisplayPageBreaksNotGiven_ReturnsFalse()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.DisplayPageBreaks = True
    Next
    sut.TurnOn
    
    sut.TurnOff
    
    For Each wks In ActiveWorkbook.Worksheets
        Assert.IsFalse wks.DisplayPageBreaks
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultHideDisplayPageBreaksFalse_ReturnsTrue()
    On Error GoTo TestFail
    
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.DisplayPageBreaks = True
    Next
    sut.TurnOn hideDisplayPageBreaks:=False
    
    sut.TurnOff
    
    For Each wks In ActiveWorkbook.Worksheets
        Assert.IsTrue wks.DisplayPageBreaks
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultAllowEventsNotGivenEventsEnabled_ReturnsTrue()
    On Error GoTo TestFail
    
    Application.EnableEvents = True
    sut.TurnOn
    
    sut.TurnOff
    
    Assert.IsTrue Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultAllowEventsNotGivenEventsDisabled_ReturnsFalse()
    On Error GoTo TestFail
    
    Application.EnableEvents = False
    sut.TurnOn
    
    sut.TurnOff
    
    Assert.IsFalse Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultAllowEventsFalseEventsEnabled_ReturnsTrue()
    On Error GoTo TestFail
    
    Application.EnableEvents = True
    sut.TurnOn allowEvents:=False
    
    sut.TurnOff
    
    Assert.IsTrue Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultAllowEventsFalseEventsDisabled_ReturnsFalse()
    On Error GoTo TestFail
    
    Application.EnableEvents = False
    sut.TurnOn allowEvents:=False
    
    sut.TurnOff
    
    Assert.IsFalse Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultAllowEventsTrueEventsEnabled_ReturnsTrue()
    On Error GoTo TestFail
    
    Application.EnableEvents = True
    sut.TurnOn allowEvents:=True
    
    sut.TurnOff
    
    Assert.IsTrue Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultAllowEventsTrueEventsDisabled_ReturnsFalse()
    On Error GoTo TestFail
    
    Application.EnableEvents = False
    sut.TurnOn allowEvents:=True
    
    sut.TurnOff
    
    Assert.IsFalse Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultStatusbarNotGivenStatusBarFalse_ReturnsFalse()
    On Error GoTo TestFail
    
    Application.StatusBar = False
    sut.TurnOn
    
    sut.TurnOff
    
    Assert.IsTrue (Application.StatusBar = False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultStatusbarGivenNullStringStatusBarFalse_ReturnsFalse()
    On Error GoTo TestFail
    
    Application.StatusBar = False
    sut.TurnOn statusBarMessage:=vbNullString
    
    sut.TurnOff
    
    Assert.IsTrue (Application.StatusBar = False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultStatusbarGivenDummyStringStatusBarFalse_ReturnsFalse()
    On Error GoTo TestFail
    
    Application.StatusBar = False
    sut.TurnOn statusBarMessage:="Just some text"
    
    sut.TurnOff
    
    Assert.IsTrue (Application.StatusBar = False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TurnOff")
Private Sub TurnOff_DefaultStatusbarGivenDummyStringStatusBarNotFalse_ReturnsString()
    On Error GoTo TestFail
    
    Application.StatusBar = "I am not False"
    sut.TurnOn statusBarMessage:="Just some text"
    
    sut.TurnOff
    
    Assert.AreEqual "I am not False", Application.StatusBar

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
