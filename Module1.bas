Attribute VB_Name = "Module1"
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Sub Main()
 Load Form1
 
End Sub

Function Lesser(a As Integer, b As Integer)
 If a <= b Then
  Lesser = a
 Else
  Lesser = b
 End If
 
End Function

Function Greater(a As Integer, b As Integer)
 If a >= b Then
  Greater = a
 Else
  Greater = b
 End If
 
End Function


