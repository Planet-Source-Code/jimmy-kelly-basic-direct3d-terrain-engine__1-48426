Attribute VB_Name = "modVector"
Public Function DrawAngle(ByRef Frm As Form, _
ByVal Angle As Integer, _
ByVal XOrigin As Long, ByVal YOrigin As Long, _
ByVal Length As Integer)

Dim ScaleMode As Integer

ScaleMode = Frm.ScaleMode

Frm.ScaleMode = vbPixels

If Angle > 360 Then Err.Raise 1609, , "Angle greater than 360"
If Angle < 0 Then Err.Raise 1610, , "Angle lower than 0"

Frm.Line (XOrigin, YOrigin)- _
(XOrigin + Cos(Radian(Angle)) * Length, _
YOrigin + Sin(Radian(Angle)) * Length)

Frm.ScaleMode = ScaleMode

End Function

Public Function Radian(ByVal Deg As Long)
Dim Pi As Double
Pi = 3.14159268
 Radian = (Deg / 180) * Pi
End Function

Public Function Degrees(ByVal Rad As Long)
Dim Pi As Double
Pi = 3.14159268
 Degrees = (Rad / 360) * Pi
End Function
