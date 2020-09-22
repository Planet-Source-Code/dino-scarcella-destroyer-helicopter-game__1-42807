Attribute VB_Name = "modDelay"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub delay(time As Long)
Dim T1, T2 As Long
Dim enddelay As Boolean

T2 = GetTickCount
 Do
 DoEvents 'DoEvents makes sure that our mouse and keyboard dont freeze-up
 T1 = GetTickCount
 
 'if 15MS has gone by, execute our next frame
 If (T1 - T2) >= time Then
  enddelay = True
  T2 = GetTickCount
 End If
'loop it until our sprite is off the screen...
Loop Until enddelay

End Sub
