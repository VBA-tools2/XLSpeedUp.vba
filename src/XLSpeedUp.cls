VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XLSpeedUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' SpeedUp Class for Excel
' |* Brian J. Satola
' |* https://chejunkie.com
'
'==========================================================================================================================
' Public Subs, Functions and Properties - oh my!
'==========================================================================================================================
' Accessible in this class: (For additional details see https://chejunkie.com/knowledge-base/speed-up-class-excel-vba/)
'==========================================================================================================================
'   |> Get | --- About                      :: String that describes the current class.
'   |> Get | --- Count                      :: Indicates the number of active run modes.
'   |> --------- DisplayApplicationSettings :: Displays current Excel/Application settings.
'   |> --------- DisplaySettings            :: Displays current SpeedUp/Excel settings.
'   |> Get | --- IsRunning                  :: Indicates run mode i.e. speed on/off.
'   |> Get | --- Name                       :: Returns the name of the class defined by constant 'C_NAME'.
'   |> --------- Reset                      :: Resets application settings to their default values.
'   |> --------- ToString                   :: Returns a string continaing a list of current settings.
'   |> --------- TurnOff                    :: Resets the 'SpeedUp' switches to their original state.
'   |> --------- TurnOn                     :: Turns on/off common Excel application settings to speed-up code execution.
'   |> Get | --- Version                    :: Returns version string for the class [e.g. #.# (year)].

'==========================================================================================================================
' Application Constants
'==========================================================================================================================
Private Const C_NAME        As String = "XLSpeedUp.cls"

'==========================================================================================================================
' Private Variables
'==========================================================================================================================
Private m_CalculationMode       As Excel.XlCalculation
Private m_ScreenUpdate          As Boolean
Private m_EnableEvents()        As Boolean
Private m_DisplayAlerts         As Boolean
Private m_IsRunning             As Boolean
Private m_Count                 As Integer


'''
''' Returns a string describing the class.
'''
Public Property Get About() As String
    
    About = "ChE Junkie VBA SpeedUp class module for Excel, " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details see  https://chejunkie.com/knowledge-base/speed-up-class-excel-vba/"
    
End Property


'''
''' Displays current Excel/Application settings.
'''
Public Sub DisplayApplicationSettings()
    
    Dim msg As String
    
    Select Case Application.Calculation
        Case XlCalculation.xlCalculationAutomatic:      msg = " (Automatic)"
        Case XlCalculation.xlCalculationManual:         msg = " (Manual)"
        Case XlCalculation.xlCalculationSemiautomatic:  msg = " (Semi-Automatic)"
    End Select
    
    msg = "CalculationMode:= " & Application.Calculation & msg & VBA.vbCrLf & _
            "ScreenUpdate:= " & Application.ScreenUpdating & VBA.vbCrLf & _
            "EnableEvents:= " & Application.EnableEvents & VBA.vbCrLf & _
            "DisplayAlerts:= " & Application.DisplayAlerts & VBA.vbCrLf
    
    MsgBox msg, vbInformation, "SpeedUp | Current Application Settings"
    
End Sub


'''
''' Displays current SpeedUp/Excel settings.
'''
Public Sub DisplayCurrentSettings()
    
    VBA.MsgBox Me.ToString(returnAsSingleLine:=False), vbInformation + vbOKOnly, "SpeedUp | Current Settings"
    
End Sub


'''
''' Returns the name of the class.
'''
Public Property Get Name() As String
    
    Name = C_NAME
    
End Property


'''
''' Returns version information for the class.
'''
Public Property Get Version() As String
    
    ' Contains a list of (historical) changes to the class within the comments of the procedure.
    '--------------------------------------------------------------------------------------------------------------------------
    Version = "Version 1.0 (2018)"                ' Initial (official) release.
    
End Property


'''
''' Indicates run mode i.e. speed on/off.
'''
Property Get IsRunning() As Boolean
    
    IsRunning = m_IsRunning
    
End Property


'''
''' Turns on/off common Excel application settings to speed-up code execution,
'''  with optional [statusBarMessage] and switch to [allowEvents] if necessary.
'''
Sub TurnOn(Optional statusBarMessage As String, Optional allowEvents As Boolean = False)
    
    If (statusBarMessage = vbNullString) Then
        statusBarMessage = "SpeedUp is on."
    End If
    
    If m_IsRunning Then
        ' Update counter so that 'TurnOff' knows when to reset switches back to normal.
        m_Count = m_Count + 1
        ReDim Preserve m_EnableEvents(m_Count)
        m_EnableEvents(m_Count) = allowEvents
    Else
        ' First call.
        ' Store current settings in class holders before making changes.
        m_Count = 1
        m_CalculationMode = Application.Calculation
        m_ScreenUpdate = Application.ScreenUpdating
        ReDim m_EnableEvents(m_Count)
        m_EnableEvents(m_Count) = Application.EnableEvents
        m_DisplayAlerts = Application.DisplayAlerts
        With Application
            .Calculation = XlCalculation.xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = allowEvents
            .DisplayAlerts = False
            .Cursor = XlMousePointer.xlWait
            .StatusBar = statusBarMessage
            .EnableCancelKey = XlEnableCancelKey.xlErrorHandler
        End With
        m_IsRunning = True
    End If
    
End Sub


'''
''' Resets the 'SpeedUp' switches to their original state.
'''
Sub TurnOff()
    
    Select Case m_Count
        Case 1                                     ' restore original values/settings.
            m_Count = 0
            m_IsRunning = False
            With Application
                .Calculation = m_CalculationMode
                .ScreenUpdating = m_ScreenUpdate
                .EnableEvents = m_EnableEvents(1)
                .DisplayAlerts = m_DisplayAlerts
                .Cursor = xlDefault
                .StatusBar = False
                .EnableCancelKey = xlInterrupt
            End With
        Case Is > 1                                ' peal-off one layer.
            m_Count = m_Count - 1
            ReDim Preserve m_EnableEvents(m_Count)
            Application.EnableEvents = m_EnableEvents(m_Count)
    End Select
    
End Sub


'''
''' Indicates the number of active run modes.
'''
Property Get Count() As Integer
    
    Count = m_Count
    
End Property


'''
''' Resets application settings to their default values.
''' Required if user terminates VBA code exection before TurnOff method is called.
'''
Sub Reset()
    
    With Application
        m_Count = 0
        m_IsRunning = False
        .Calculation = Excel.XlCalculation.xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .CalculateBeforeSave = True
        .Cursor = Excel.XlMousePointer.xlDefault
        .StatusBar = False
        .EnableCancelKey = Excel.XlEnableCancelKey.xlInterrupt
    End With
    
End Sub


'''
''' Returns a string continaing a list of current settings.
'''
Property Get ToString(Optional returnAsSingleLine As Boolean = True) As String
    
    Select Case Application.Calculation
        Case XlCalculation.xlCalculationAutomatic:      ToString = " (Automatic)"
        Case XlCalculation.xlCalculationManual:         ToString = " (Manual)"
        Case XlCalculation.xlCalculationSemiautomatic:  ToString = " (Semi-Automatic)"
    End Select
    
    ToString = "IsRunning:= " & m_IsRunning & VBA.vbCrLf & _
                "Count:= " & m_Count & VBA.vbCrLf & _
                "CalculationMode:= " & Application.Calculation & ToString & VBA.vbCrLf & _
                "ScreenUpdate:= " & Application.ScreenUpdating & VBA.vbCrLf & _
                "EnableEvents(" & Me.Count & "):= " & Application.EnableEvents & VBA.vbCrLf & _
                "DisplayAlerts:= " & Application.DisplayAlerts
    If returnAsSingleLine Then
        ToString = VBA.Replace(ToString, VBA.vbCrLf, ", ")
    End If
    
End Property


Private Sub Class_Initialize()
    '//// Class constructor.
    '--------------------------------------------------------------------------------------------------------------------------
Debug.Print "|> Initializing:= " & Me.Name
    
End Sub
