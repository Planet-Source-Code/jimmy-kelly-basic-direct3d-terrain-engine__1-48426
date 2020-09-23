Attribute VB_Name = "modColor"
Public Enum RCLR
 R = 1
 G = 2
 B = 3
End Enum

Public Function LongToRGB(ColorValue As Long, ByVal ReturnColor As RCLR) As Long
    
    Dim rCol As Long, gCol As Long, bCol As Long
    rCol = ColorValue And &H10000FF 'this uses binary comparason
    gCol = (ColorValue And &H100FF00) / (2 ^ 8)
    bCol = (ColorValue And &H1FF0000) / (2 ^ 16)
    
   If rCol > 255 Then rCol = 255
   If rCol < 0 Then rCol = 0
    
   If gCol > 255 Then gCol = 255
   If gCol < 0 Then gCol = 0
    
   If bCol > 255 Then bCol = 255
   If bCol < 0 Then bCol = 0
    
    Select Case ReturnColor
    Case R Or 1
     LongToRGB = rCol
    Case G Or 2
     LongToRGB = gCol
    Case B Or 3
     LongToRGB = bCol
    End Select
    
End Function

Public Function GetR(ColorValue As Long) As Long
    
    Dim rCol As Long, gCol As Long, bCol As Long
    rCol = ColorValue And &H10000FF 'this uses binary comparason
    gCol = (ColorValue And &H100FF00) / (2 ^ 8)
    bCol = (ColorValue And &H1FF0000) / (2 ^ 16)

   If rCol > 255 Then rCol = 255
   If rCol < 0 Then rCol = 0

     GetR = rCol
    
End Function

Public Function GetG(ColorValue As Long) As Long
    
    Dim rCol As Long, gCol As Long, bCol As Long
    rCol = ColorValue And &H10000FF 'this uses binary comparason
    gCol = (ColorValue And &H100FF00) / (2 ^ 8)
    bCol = (ColorValue And &H1FF0000) / (2 ^ 16)

   If gCol > 255 Then gCol = 255
   If gCol < 0 Then gCol = 0

     GetG = gCol
    
End Function

Public Function GetB(ColorValue As Long) As Long
    
    Dim rCol As Long, gCol As Long, bCol As Long
    rCol = ColorValue And &H10000FF 'this uses binary comparason
    gCol = (ColorValue And &H100FF00) / (2 ^ 8)
    bCol = (ColorValue And &H1FF0000) / (2 ^ 16)

   If bCol > 255 Then bCol = 255
   If bCol < 0 Then bCol = 0

     GetB = bCol
    
End Function

Public Function RGBToLong(ByVal R, G, B As Long)

RGBToLong = RGB(R, G, B)

End Function

Public Function RGBToHex(ByVal R, G, B As Long)

Dim Result As String

Result = Hex(R) + Hex(G) + Hex(B)

RGBToHex = Val(Result)

End Function
