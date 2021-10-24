VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Option Compare Text
'
' Registry Class
' |* Brian J. Satola
' |* https://chejunkie.com
'
'==========================================================================================================================
' Public Subs, Functions and Properties - oh my!
'==========================================================================================================================
' Accessible in this class: (For additional details see http://docs.chejunkie.com/measure-performance-vba/)
'==========================================================================================================================
'   |> Get | --- About                :: String that describes the current class.
'   |> --------- DisplayPerformance   :: Displays 'PerformanceString' in a message-box, with optional [msg] and [title].
'   |> --------- GetHoursElapsed      :: Returns the elapsed time in [hours].
'   |> --------- GetMinutesElapsed    :: Returns the elapsed time in [minutes].
'   |> --------- GetSecondsElapsed    :: Returns the elapsed time in [seconds].
'   |> Get | --- IsRunning            :: Returns the status of the timer.
'   |> Get | --- Name                 :: Returns the name of the class defined by constant 'C_NAME'.
'   |> --------- PerformanceString    :: Returns a string defining the elapsed time in [seconds].
'   |> --------- PrintPerformance     :: Prints the 'PerformanceString' to Immediate, with optional [msg] and [title].
'   |> --------- Reset                :: Resets the start time to zero.
'   |> --------- StartTimer           :: Starts the clock/timer for the performance string.
'   |> --------- StopTimer            :: Stops the clock/timer for the performance string.
'   |> Get | --- Version              :: Returns version string for the class [e.g. #.# (year)].

'==========================================================================================================================
' API Declarations
'==========================================================================================================================
' Standard timer: http://www.pcreview.co.uk/forums/thread-985429.php
Private Declare PtrSafe Function TimeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Long

'==========================================================================================================================
' Application Constants
'==========================================================================================================================
Private Const C_NAME        As String = "StopWatch.cls"

Public Enum StopWatchDisplayFormat
    SWDisplaySeconds
    SWDisplayHoursMinutesSeconds
    SWDisplayHoursMinutesSecondsMilliseconds
End Enum

'==========================================================================================================================
' Private Variables
'==========================================================================================================================
Private m_beginTime         As Long
Private m_endTime           As Long
Private m_IsRunning         As Boolean

'''
''' Returns elapsed time in a message-box, with optional [msg], [title] and [displayFormat].
''' The default display format is in seconds.
'''
Public Sub DisplayPerformance(Optional Msg As String, Optional title As String, _
                                Optional displayFormat As StopWatchDisplayFormat = StopWatchDisplayFormat.SWDisplaySeconds)

    If (title = "") Then
        title = "Performance"
    ElseIf (VBA.InStr(UCase(title), "PERFORMANCE") = 0) Then
        title = "Performance | " & title
    End If
    If (Msg <> "") Then
        MsgBox Msg & VBA.vbCrLf & VBA.vbCrLf & Me.PerformanceString(returnFormat:=displayFormat), vbInformation, title
    Else
        MsgBox Me.PerformanceString(returnFormat:=displayFormat), vbInformation, title
    End If
End Sub


'''
''' Returns elapsed time in hours, with decimal remainder.
'''
Public Function GetHoursElapsed() As Double

    GetHoursElapsed = Me.GetSecondsElapsed / 3600
    
End Function

'''
''' Returns elapsed time in minutes, with decimal remainder.
'''
Public Function GetMinutesElapsed() As Double

    GetMinutesElapsed = Me.GetSecondsElapsed / 60
    
End Function

'''
''' Returns a formatted string in units based on the elapsed time.
''' Returns as hh:mm:ss, with an option to also [returnMilliseconds] as hh:mm:ss(ms).
'''
Private Function PerformanceStringX(Optional returnMilliseconds As Boolean = False) As String

Dim currentTime         As Double
Dim vint(4)             As Integer 'hh:mm:ss(ms)

    currentTime = Me.GetSecondsElapsed
    vint(1) = Int(currentTime / 3600)
    vint(2) = Int(((currentTime / 3600) - vint(1)) * 60)
    vint(3) = Int((((currentTime / 3600) - vint(1)) * 60 - vint(2)) * 60)
    vint(4) = Int(Replace(Format(currentTime - Int(currentTime), "#.000"), ".", ""))

    If returnMilliseconds Then
        PerformanceStringX = Format(vint(1), "00") & ":" & Format(vint(2), "00") & ":" & Format(vint(3), "00") & "(" & Format(vint(4), "000") & ")"
    Else
        '// hour(s):
        If vint(1) > 1 Then
            PerformanceStringX = vint(1) & " hours : "
        ElseIf vint(1) = 1 Then
            PerformanceStringX = "1 hour : "
        End If
        '// minute(s):
        If vint(2) > 1 Then
            PerformanceStringX = PerformanceStringX & vint(2) & " minutes : "
        ElseIf vint(2) = 1 Then
            PerformanceStringX = PerformanceStringX & "1 minute : "
        End If
        '// second(s):
        PerformanceStringX = PerformanceStringX & vint(3) & "." & Format(vint(4), "000") & " seconds elapsed."
    End If

End Function

'''
''' Prints the elapsed time to the Immediate window, with optional [title] and [displayFormat].
''' The default display format is in seconds.
'''
Public Sub PrintPerformance(Optional title As String, Optional displayFormat As StopWatchDisplayFormat = StopWatchDisplayFormat.SWDisplaySeconds)

    ' Print title.
    If (title = "") Then
        title = "Performance"
    ElseIf (VBA.InStr(UCase(title), "PERFORMANCE") = 0) Then
        title = "Performance | " & title
    End If
    Debug.Print ""
    Debug.Print "=========================================================================================================================="
    Debug.Print title
    Debug.Print "=========================================================================================================================="

    ' Print performance string
    Debug.Print Me.PerformanceString(returnFormat:=displayFormat)
    
End Sub

'''
''' Resets the start time to zero.
'''
Public Sub Reset()

    Me.StartTimer
    
End Sub

'''
''' Returns the status of the timer.
'''
Public Property Get IsRunning() As Boolean

    IsRunning = m_IsRunning
    
End Property




'''
''' Starts the clock/timer for the performance string.
'''
Public Sub StartTimer()

    m_IsRunning = True
    m_endTime = 0
    m_beginTime = TimeGetTime()
    
End Sub

'''
''' Stops the clock/timer for the performance string.
'''
Public Sub StopTimer()

    m_endTime = TimeGetTime()
    m_IsRunning = False
    
End Sub


'''
''' Returns a string defining the elapsed time, with optional [displayFormat].
''' Default display format is in seconds.
'''
Public Function PerformanceString(Optional returnFormat As StopWatchDisplayFormat = SWDisplaySeconds) As String

    Select Case returnFormat
        Case SWDisplaySeconds ' seconds
            PerformanceString = Me.GetSecondsElapsed & " seconds elapsed."
        Case SWDisplayHoursMinutesSeconds: ' hh:mm:ss
            PerformanceString = PerformanceStringX(returnMilliseconds:=False)
        Case SWDisplayHoursMinutesSecondsMilliseconds: ' hh:mm:ss(ms)
            PerformanceString = PerformanceStringX(returnMilliseconds:=True)
        Case Else ' invalid input type.
    End Select

End Function

'''
''' Returns the elapsed time in milliseconds (1/1000 of a second).
'''
Private Function ElapsedTime() As Long

    If m_IsRunning Then
        m_endTime = TimeGetTime()
    End If
    ElapsedTime = m_endTime - m_beginTime
    
End Function

'''
''' Returns a string describing the class.
'''
Public Property Get About() As String

    About = "ChE Junkie VBA StopWatch class module, " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details see  https://chejunkie.com/knowledge-base/measure-code-performance-vba/"
    
End Property
'''
''' Returns version information for the class.
'''
Public Property Get Version() As String

' Contains a list of (historical) changes to the class within the comments of the procedure.
'--------------------------------------------------------------------------------------------------------------------------
   'Version = "Version 1.0 (2015)"  ' Initial (official) release.
   'Version = "Version 1.5 (2015)"  '
                                    ' |> Cleaned up code, including procedure descriptions.
                                    ' |> Renamed class name from 'Timer.cls' to 'StopWatch.cls'.
                                    ' |> Added "Name" Get property (literal defintion).
                                    ' |> Added "About" Get property for use in e.g. display messages.
                                    ' |> Added "Version" Get property to track class changes.
                                    '
   'Version = "Version 2.0 (2015)"
                                    ' |> Syntax corrections to module-level variables (prefix:= "m_").
                                    ' |> Updated constructor print statement to correct class name.
                                    ' |> Changed m_EndTime from Variant to Long-type (inc. StartTimer method --> 0).
                                    ' |> Added new PerformanceStringX method formats results into hh:mm:ss (and ms).
                                    '
   'Version = "Version 3.0 (2017)"
                                    ' |> Created enumeration to handle display format.
                                    ' |> Now only PerformanceString function, using new enum.
                                    ' |--> PerformanceStringX is now private.
                                    ' |> Now only DisplayPerformance sub, using new enum.
                                    ' |--> DisplayPerformanceX was removed.
                                    ' |> Update About string to new https://chejunkie.com/ address.
                                    ' |> Renamed SecondsElapsed to GetSecondsElapsed (more descriptive).
                                    ' |> Renamed MinutesElapsed and HoursElapsed to GetMinutesElapsed/GetHoursElapsed.
                                    ' |> Other minor cosmetic changes to code comments.
    Version = "Version 3.1 (2018)"
                                    ' |> Added PrintPerformance method.
                                    
End Property


'''
''' Returns the name of the class.
'''
Public Property Get Name() As String

    Name = C_NAME
    
End Property
'''
''' Returns the number of elapsed seconds, with decimal remainder.
'''
Public Function GetSecondsElapsed() As Double

    GetSecondsElapsed = ElapsedTime() / 1000
    
End Function


Private Sub Class_Initialize()
'//// Class constructor.
'--------------------------------------------------------------------------------------------------------------------------
    Debug.Print "|> Initializing:= " & Me.Name
    
End Sub


