VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XLSpeedUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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
'   |> --------- ToString                   :: Returns a string containing a list of current settings.
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
Private m_EnableAnimations      As Boolean
Private m_StatusBar             As Boolean
Private m_IsRunning             As Boolean
Private m_Count                 As Integer
Private m_Silent                As Boolean


'Returns a string describing the class.
Public Property Get About() As String
    
    About = "ChE Junkie VBA SpeedUp class module for Excel, " & Me.Version & _
            "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details see  " & _
            "<https://chejunkie.com/knowledge-base/speed-up-class-excel-vba/>"
    
End Property


'Displays current Excel/Application settings.
Public Sub DisplayApplicationSettings()
    
    Dim Msg As String
    
    Select Case Application.Calculation
        Case XlCalculation.xlCalculationAutomatic
            Msg = " (Automatic)"
        Case XlCalculation.xlCalculationManual
            Msg = " (Manual)"
        Case XlCalculation.xlCalculationSemiautomatic
            Msg = " (Semi-Automatic)"
    End Select
    
    With Application
        Msg = "CalculationMode:= " & .Calculation & Msg & VBA.vbCrLf & _
                "ScreenUpdate:= " & .ScreenUpdating & VBA.vbCrLf & _
                "EnableEvents:= " & .EnableEvents & VBA.vbCrLf & _
                "DisplayAlerts:= " & .DisplayAlerts & VBA.vbCrLf & _
                "EnableAnimations:= " & .EnableAnimations
    End With
    
    MsgBox Msg, vbInformation, "SpeedUp | Current Application Settings"
    
End Sub


'Displays current SpeedUp/Excel settings.
Public Sub DisplayCurrentSettings()
    VBA.MsgBox _
            Me.ToString(returnAsSingleLine:=False), _
            vbInformation + vbOKOnly, _
            "SpeedUp | Current Settings"
End Sub


'Returns the name of the class.
Public Property Get Name() As String
    Name = C_NAME
End Property


Public Property Get Silent() As Boolean
    Silent = m_Silent
End Property
Public Property Let Silent(Value As Boolean)
    m_Silent = Value
End Property


'Returns version information for the class.
Public Property Get Version() As String
    
    'Contains a list of (historical) changes to the class within the comments of the procedure.
    '--------------------------------------------------------------------------------------------------------------------------
    Version = "Version 1.0 (2018)"                'Initial (official) release.
    
End Property


'Indicates run mode i.e. speed on/off.
Property Get IsRunning() As Boolean
    IsRunning = m_IsRunning
End Property


'Turns on/off common Excel application settings to speed-up code execution,
'with optional switches to [hideDisplayPageBreaks] and [allowEvents], and
'to set a [statusBarMessage] if necessary.
Sub TurnOn( _
    Optional ByVal hideDisplayPageBreaks As Boolean = True, _
    Optional ByVal allowEvents As Boolean = False, _
    Optional ByVal statusBarMessage As String = vbNullString _
)
    
    If Len(statusBarMessage) = 0 Then
        Dim Message As String
        Message = "SpeedUp is on."
    ElseIf TypeName(statusBarMessage) = "String" Then
        Message = statusBarMessage
    End If
    
    If m_IsRunning Then
        'Update counter so that 'TurnOff' knows when to reset switches back to normal.
        m_Count = m_Count + 1
        ReDim Preserve m_EnableEvents(m_Count)
        m_EnableEvents(m_Count) = allowEvents
    Else
        'First call.
        'Store current settings in class holders before making changes.
        With Application
            m_Count = 1
            m_CalculationMode = .Calculation
            m_ScreenUpdate = .ScreenUpdating
            ReDim m_EnableEvents(m_Count)
            m_EnableEvents(m_Count) = .EnableEvents
            m_DisplayAlerts = .DisplayAlerts
            m_EnableAnimations = .EnableAnimations
            
            HideDisplayPageBreaksOnAllSheets hideDisplayPageBreaks
            
            .Calculation = XlCalculation.xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = allowEvents
            .DisplayAlerts = False
            .EnableAnimations = False
            .Cursor = XlMousePointer.xlWait
            SetStatusBar Message
            .EnableCancelKey = XlEnableCancelKey.xlErrorHandler
        End With
        m_IsRunning = True
    End If
    
End Sub


'Resets the 'SpeedUp' switches to their original state.
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
                .EnableAnimations = m_EnableAnimations
                .Cursor = xlDefault
                If m_StatusBar = False Then
                    .StatusBar = False
                End If
                .EnableCancelKey = xlInterrupt
            End With
        Case Is > 1                                ' peal-off one layer.
            m_Count = m_Count - 1
            ReDim Preserve m_EnableEvents(m_Count)
            Application.EnableEvents = m_EnableEvents(m_Count)
    End Select
    
End Sub


'Indicates the number of active run modes.
Property Get Count() As Integer
    Count = m_Count
End Property


'Resets application settings to their default values.
'Required if user terminates VBA code execution before TurnOff method is called.
Sub Reset()
    With Application
        m_Count = 0
        m_IsRunning = False
        .Calculation = Excel.XlCalculation.xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .EnableAnimations = True
        .CalculateBeforeSave = True
        .Cursor = Excel.XlMousePointer.xlDefault
        .StatusBar = False
        .EnableCancelKey = Excel.XlEnableCancelKey.xlInterrupt
    End With
End Sub


'Returns a string containing a list of current settings.
Property Get ToString( _
    Optional returnAsSingleLine As Boolean = True _
) As String
    
    Select Case Application.Calculation
        Case XlCalculation.xlCalculationAutomatic
            ToString = " (Automatic)"
        Case XlCalculation.xlCalculationManual
            ToString = " (Manual)"
        Case XlCalculation.xlCalculationSemiautomatic
            ToString = " (Semi-Automatic)"
    End Select
    
    With Application
        ToString = _
                "IsRunning:= " & m_IsRunning & VBA.vbCrLf & _
                "Count:= " & m_Count & VBA.vbCrLf & _
                "CalculationMode:= " & .Calculation & ToString & VBA.vbCrLf & _
                "ScreenUpdate:= " & .ScreenUpdating & VBA.vbCrLf & _
                "EnableEvents(" & Me.Count & "):= " & .EnableEvents & VBA.vbCrLf & _
                "DisplayAlerts:= " & .DisplayAlerts & VBA.vbCrLf & _
                "EnableAnimations:= " & .EnableAnimations
    End With
    
    If returnAsSingleLine Then
        ToString = VBA.Replace(ToString, VBA.vbCrLf, ", ")
    End If
    
End Property


Private Sub Class_Initialize()
    '//// Class constructor.
    '--------------------------------------------------------------------------------------------------------------------------
    m_Silent = True
    
    ShowMessage "|> Initializing:= " & Me.Name
End Sub


Private Sub SetStatusBar(ByVal Message As String)
    With Application
        'only change StatusBar text when there is none
        If .StatusBar = False Then
            m_StatusBar = False
            .StatusBar = Message
        Else
            m_StatusBar = True
        End If
    End With
End Sub


Private Sub HideDisplayPageBreaksOnAllSheets( _
    ByVal hideDisplayPageBreaks As Boolean _
)
    If hideDisplayPageBreaks Then
        Dim wks As Worksheet
        For Each wks In ActiveWorkbook.Worksheets
            wks.DisplayPageBreaks = False
        Next
    End If
End Sub


Private Sub ShowMessage(ByVal Msg As String)
    If m_Silent = False Then
        Debug.Print Msg
    End If
End Sub
