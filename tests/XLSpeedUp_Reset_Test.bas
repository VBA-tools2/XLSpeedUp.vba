Attribute VB_Name = "XLSpeedUp_Reset_Test"
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
'@TestMethod("Reset")
Private Sub Reset_Calculation_ReturnsXlCalculationAutomatic()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.AreEqual XlCalculation.xlCalculationAutomatic, Application.Calculation

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_CalculateBeforeSave_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.IsTrue Application.CalculateBeforeSave
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_Count_ReturnsZero()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.AreEqual 0, sut.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_Cursor_ReturnsXlDefault()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.AreEqual XlMousePointer.xlDefault, Application.Cursor

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_DisplayAlerts_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.IsTrue Application.DisplayAlerts

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_EnableAnimations_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.IsTrue Application.EnableAnimations

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_EnableCancelKey_ReturnsXlInterrupt()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.AreEqual XlEnableCancelKey.xlInterrupt, Application.EnableCancelKey

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_EnableEvents_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.IsTrue Application.EnableEvents

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_IsRunning_ReturnsFalse()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.IsFalse sut.IsRunning

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_ScreenUpdating_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.IsTrue Application.ScreenUpdating

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Reset")
Private Sub Reset_StatusBar_ReturnsFalse()
    On Error GoTo TestFail
    
    sut.Reset
    
    Assert.IsTrue (Application.StatusBar = False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
