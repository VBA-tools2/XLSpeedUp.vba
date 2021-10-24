Attribute VB_Name = "XLSpeedUp_Count_Test"
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
'@TestMethod("Count")
Private Sub Count_NotTurnOn_ReturnsZero()
    On Error GoTo TestFail
    
    Assert.AreEqual 0, sut.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Count")
Private Sub Count_TurnOn_ReturnsOne()
    On Error GoTo TestFail
    
    sut.TurnOn
    
    Assert.AreEqual 1, sut.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Count")
Private Sub Count_TurnOnOn_ReturnsTwo()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOn
    
    Assert.AreEqual 2, sut.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Count")
Private Sub Count_TurnOnOnOff_ReturnsOne()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOn
    sut.TurnOff
    
    Assert.AreEqual 1, sut.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Count")
Private Sub Count_TurnOnOnOffOff_ReturnsZero()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOn
    sut.TurnOff
    sut.TurnOff
    
    Assert.AreEqual 0, sut.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Count")
Private Sub Count_TurnOnOff_ReturnsZero()
    On Error GoTo TestFail
    
    sut.TurnOn
    sut.TurnOff
    
    Assert.AreEqual 0, sut.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
