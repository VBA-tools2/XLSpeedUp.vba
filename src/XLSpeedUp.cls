VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XLSpeedUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'@Exposed

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
'   |> Get | --- About                          :: String that describes the current class.
'   |> Get | --- Count                          :: Indicates the number of active run modes.
'   |> --------- DisplayApplicationSettings     :: Displays current Excel/Application settings.
'   |> --------- DisplaySettings                :: Displays current SpeedUp/Excel settings.
'   |> Get | --- IsRunning                      :: Indicates run mode i.e. speed on/off.
'   |> Get | --- Name                           :: Returns the name of the class defined by constant 'C_NAME'.
'   |> --------- Reset                          :: Resets application settings to their default values.
'   |> Get | Let Silent                         :: Suppresses messages if set to TRUE
'   |> --------- ToString                       :: Returns a string containing a list of current settings.
'   |> --------- TurnOff                        :: Resets the 'SpeedUp' switches to their original state.
'   |> --------- TurnOn                         :: Turns on/off common Excel application settings to speed-up code execution.
'   |> Get | --- Version                        :: Returns version string for the class [e.g. #.# (year)].

'==========================================================================================================================
' Application Constants
'==========================================================================================================================
Private Const C_NAME    As String = "XLSpeedUp.cls"

'==========================================================================================================================
' Private Variables
'==========================================================================================================================
Private Type TXLSpeedUp
    CalculationMode     As Excel.XlCalculation
    Count               As Byte
    DisplayAlerts       As Boolean
    EnableAnimations    As Boolean
    EnableEvents()      As Boolean
    IsRunning           As Boolean
    ScreenUpdate        As Boolean
    Silent              As Boolean
    StatusBar           As Boolean
End Type
Private This As TXLSpeedUp


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
    Silent = This.Silent
End Property

Public Property Let Silent(Value As Boolean)
    This.Silent = Value
End Property


'Returns version information for the class.
Public Property Get Version() As String
    
    'Contains a list of (historical) changes to the class within the comments of the procedure.
    '----------------------------------------------------------------------------------------------------
    'Version = "Version 1.0 (2018)"  ' Initial (official) release.
    Version = "Version 2.0 (2020)"  ' Thank you Stefan Pinnow for sharing!
                                    ' |> Remove some typos
                                    ' |> Remove some unnecessary code
                                    ' |> Add some more (small) features like e.g. disabling animations.
    
End Property


'Indicates run mode i.e. speed on/off.
Property Get IsRunning() As Boolean
    IsRunning = This.IsRunning
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
    
    If This.IsRunning Then
        With This
            'Update counter so that 'TurnOff' knows when to reset switches back to normal.
            .Count = .Count + 1
            ReDim Preserve .EnableEvents(.Count)
            .EnableEvents(.Count) = allowEvents
        End With
    Else
        'First call.
        'Store current settings in class holders before making changes.
        With Application
            This.Count = 1
            This.CalculationMode = .Calculation
            This.ScreenUpdate = .ScreenUpdating
            ReDim This.EnableEvents(This.Count)
            This.EnableEvents(This.Count) = .EnableEvents
            This.DisplayAlerts = .DisplayAlerts
            This.EnableAnimations = .EnableAnimations
            
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
        This.IsRunning = True
    End If
    
End Sub


'Resets the 'SpeedUp' switches to their original state.
Sub TurnOff()
    
    Select Case This.Count
        Case 1                                               ' restore original values/settings.
            This.Count = 0
            This.IsRunning = False
            With Application
                .Calculation = This.CalculationMode
                .ScreenUpdating = This.ScreenUpdate
                .EnableEvents = This.EnableEvents(1)
                .DisplayAlerts = This.DisplayAlerts
                .EnableAnimations = This.EnableAnimations
                .Cursor = xlDefault
                If This.StatusBar = False Then
                    .StatusBar = False
                End If
                .EnableCancelKey = xlInterrupt
            End With
        Case Is > 1                                          ' peal-off one layer.
            This.Count = This.Count - 1
            ReDim Preserve This.EnableEvents(This.Count)
            Application.EnableEvents = This.EnableEvents(This.Count)
    End Select
    
End Sub


'Indicates the number of active run modes.
Property Get Count() As Byte
    Count = This.Count
End Property


'Resets application settings to their default values.
'Required if user terminates VBA code execution before TurnOff method is called.
Sub Reset()
    With Application
        This.Count = 0
        This.IsRunning = False
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
                "IsRunning:= " & This.IsRunning & VBA.vbCrLf & _
                "Count:= " & This.Count & VBA.vbCrLf & _
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
    '--------------------------------------------------------------------------
    This.Silent = True
    
    ShowMessage "|> Initializing:= " & Me.Name
End Sub


Private Sub SetStatusBar(ByVal Message As String)
    With Application
        'only change StatusBar text when there is none
        If .StatusBar = False Then
            This.StatusBar = False
            .StatusBar = Message
        Else
            This.StatusBar = True
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
    If Not This.Silent Then
        Debug.Print Msg
    End If
End Sub
