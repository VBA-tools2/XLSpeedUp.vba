Attribute VB_Name = "XLSpeedUp_IsRunning_Test"
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
'@TestMethod("IsRunning")
Private Sub IsRunning_NotTurnOn_ReturnsFalse()
    On Error GoTo TestFail
    
    Assert.IsFalse sut.IsRunning

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsRunning")
Private Sub IsRunning_TurnOn_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.TurnOn
    
    Assert.IsTrue sut.IsRunning

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsRunning")
Private Sub IsRunning_TurnOnOn_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOn
    
    Assert.IsTrue sut.IsRunning

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsRunning")
Private Sub IsRunning_TurnOnOnOff_ReturnsTrue()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOn
    sut.TurnOff
    
    Assert.IsTrue sut.IsRunning

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsRunning")
Private Sub IsRunning_TurnOnOnOffOff_ReturnsFalse()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOn
    sut.TurnOff
    sut.TurnOff
    
    Assert.IsFalse sut.IsRunning

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsRunning")
Private Sub IsRunning_TurnOnOff_ReturnsFalse()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOff
    
    Assert.IsFalse sut.IsRunning

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
