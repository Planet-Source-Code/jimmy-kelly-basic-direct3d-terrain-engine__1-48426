Attribute VB_Name = "modGUI"
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim CPoint As POINTAPI

Public Function SetWindowSize(ByRef Frm As Form, Optional ByVal Width As Integer, _
Optional ByVal Height As Integer)

If IsMissing(Width) = False Then Frm.Width = Width * Screen.TwipsPerPixelX
If IsMissing(Height) = False Then Frm.Height = Height * Screen.TwipsPerPixelY

End Function

Public Function CursorShow()
 ShowCursor 1
End Function

Public Function CursorHide()
 ShowCursor 0
End Function

Public Function CursorX(ByVal lValue As Long)
 GetCursorPos CPoint
 SetCursorPos lValue, CPoint.y
End Function

Public Function GetCursorX() As Long
 GetCursorPos CPoint
 GetCursorX = CPoint.x
End Function

Public Function CursorY(ByVal lValue As Long)
 GetCursorPos CPoint
 SetCursorPos CPoint.x, lValue
End Function

Public Function GetCursorY() As Long
 GetCursorPos CPoint
 GetCursorY = CPoint.y
End Function
