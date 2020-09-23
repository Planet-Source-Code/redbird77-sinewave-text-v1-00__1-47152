Attribute VB_Name = "mSineWaveHelper"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL As Long = 1
Public Const PI  As Double = 3.14159265358979
Public Const TPI As Double = 6.28318530717959

Public Function DegToRad(ByVal dDeg As Double) As Double
    DegToRad = dDeg * (PI / 180)
End Function

Public Function RadToDeg(ByVal dRad As Double) As Double
    ' 1 rad ~~ 57.2 deg
    ' PI = Atn(1) * 4
    RadToDeg = dRad * 180 / PI
End Function

Public Function LNGtoHEX(ByVal lColor As Long) As String

    Dim b(2) As Byte
    
    CopyMemory b(0), lColor, 3

    ' You can't just Hex$ a long to get a web-ready hex triplet color string,
    ' it'll be rerversed (i.e. ff6034 instead of 3460ff).
    LNGtoHEX = "#" & Right$("00000" & LCase$(Hex$(RGB(b(2), b(1), b(0)))), 6)
    
End Function

Public Function IsPrint(ByVal b As Byte) As Boolean
    IsPrint = IIf(b > 32 And (b < 127 Or b > 160), True, False)
End Function

Public Function Normalize(ByVal dUnVal As Double) As Double

' This function takes a number in the [-1,1] range and returns
' the corresponding number in the [0,255] range.  The ouput range can
' be a parameter also, but in this case it's just hardcoded.

' UNlo = 1, UNhi = -1, Nlo = 255, Nhi = 0
' Nval = Nlo + (UNval - UNlo) * (Nhi - Nlo) / (UNhi - UNlo)

    If dUnVal < -1 Then
        dUnVal = -1
    ElseIf dUnVal > 1 Then
        dUnVal = 1
    End If

    Normalize = 255 + (dUnVal - 1) * 255 / 2

End Function
