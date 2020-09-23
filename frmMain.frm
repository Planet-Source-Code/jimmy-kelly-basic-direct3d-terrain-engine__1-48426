VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Direct3D8 IM Example"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   239
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'Atomic™ Direct3D Engine (I changed the name. I like this better)
'©2003 by Backwoods Interactive. All Rights Reserved
'DirectX is a registered trademark of Microsoft
'====================================================
'This code is incomplete and mostly uncommented
'however it pretty much explains itself.
'====================================================

Dim WithEvents D3D As clsDirect3D8
Attribute D3D.VB_VarHelpID = -1
Dim Grass() As clsRectangle8
Attribute Grass.VB_VarHelpID = -1
Dim WithEvents Keyboard As clsKeyboard8
Attribute Keyboard.VB_VarHelpID = -1

Const Forward As Single = 0.5
Const Backward As Single = 0.5
Const Turn As Single = 0.05

Const ScreenWidth As Long = 640
Const ScreenHeight As Long = 480

Const UnitWidth As Long = 4
Const UnitHeight As Long = 4

Dim Fullscreen As Integer
Dim UseHardware As Integer
Dim bHW As Boolean

Private Sub D3D_FinishRender(ByVal FPS As Long, ByVal MS As Long)

 Me.Caption = "Direct3D8 IM Example (" & FPS & ", " & Fix(D3D.CamX) & ", " & _
 Fix(D3D.CamY) & ", " & Fix(D3D.CamZ) & ", " & Fix(D3D.CamAngle) & ")"

End Sub

Private Sub D3D_PreRender()

 D3D.Clear 0 'RGB(128, 0, 0)
 
 If Keyboard.GetKeyState(DIK_UPARROW) Then
  D3D.MoveForward Forward
 End If
 
 If Keyboard.GetKeyState(DIK_DOWNARROW) Then
  D3D.MoveBackward Backward
 End If
 
 If Keyboard.GetKeyState(DIK_RIGHTARROW) Then
  D3D.CamAngle = D3D.CamAngle - Turn
 End If
 
 If Keyboard.GetKeyState(DIK_LEFTARROW) Then
  D3D.CamAngle = D3D.CamAngle + Turn
 End If
 
 If Keyboard.GetKeyState(DIK_PGUP) Then
  D3D.CamPitch = D3D.CamPitch + Turn
 End If
 
 If Keyboard.GetKeyState(DIK_PGDN) Then
  D3D.CamPitch = D3D.CamPitch - Turn
 End If
 
 If Keyboard.GetKeyState(DIK_NUMPADPLUS) Then
  D3D.CamY = D3D.CamY - Forward
 End If
 
 If Keyboard.GetKeyState(DIK_NUMPADMINUS) Then
  D3D.CamY = D3D.CamY + Backward
 End If
 
 If Keyboard.GetKeyState(DIK_ESCAPE) Then
  Shutdown
  Unload Me
 End If
 
End Sub

Private Sub Form_Load()

'Be aware that not using Hardware mode
'especially when in fullscreen could lockup
'VB as I found out personally

Fullscreen = MsgBox("Use fullscreen mode?", vbYesNoCancel, "Direct3D Example")
UseHardware = MsgBox("Use hardware acceleration?", vbYesNo, "Direct3D Example")

SetWindowSize Me, ScreenWidth, ScreenHeight

Me.Show

Set Keyboard = New clsKeyboard8
Keyboard.Create Me.hWnd

Set D3D = New clsDirect3D8

Select Case UseHardware
 Case vbYes
  bHW = True
 Case vbNo
  bHW = False
End Select

Select Case Fullscreen
 Case vbYes
  Me.BorderStyle = 0
  CursorHide
  D3D.InitD3D Me.hWnd, ScreenWidth, ScreenHeight, True, bHW
 Case vbNo
  Me.BorderStyle = 2
  D3D.InitD3D Me.hWnd, ScreenWidth, ScreenHeight, False, bHW
 Case vbCancel
  Shutdown
  Unload Me
End Select

LoadTerrain

Me.Show

With D3D
 .CamX = 0
 .CamY = 8
 .CamZ = 0
 .FogColor = 0
 .FogStart = 0
 .FogEnd = 40
 .FogEnable = False
 .BeginRender
End With

End Sub

Public Function LoadTerrain()
 
 Dim picHeight As PictureBox
 Set picHeight = Controls.Add("VB.PictureBox", "picHeight")
 picHeight.AutoSize = True
 picHeight.BorderStyle = 0
 picHeight.ScaleMode = 3
 picHeight.Visible = True
 picHeight.Picture = LoadPicture(JoinPath(App.Path, "Height.bmp"))
 
 Dim I As Long
 Dim J As Long
 Dim W As Long
 Dim H As Long
 Dim C(3) As Long
 
 'picHeight.Width = picHeight.Picture.Width
 'picHeight.Height = picHeight.Picture.Height
 
 W = picHeight.ScaleWidth - UnitWidth * 2
 H = picHeight.ScaleHeight - UnitHeight * 2
 
 Dim Height As Long
 Height = 20
 
 For I = 0 To W Step UnitWidth
  For J = 0 To H Step UnitHeight
  C(0) = GetR(picHeight.Point(I, J)) / Height
  C(1) = GetR(picHeight.Point(I + UnitWidth, J)) / Height
  C(2) = GetR(picHeight.Point(I, J + UnitHeight)) / Height
  C(3) = GetR(picHeight.Point(I + UnitWidth, J + UnitHeight)) / Height
   CreateGrass C(0), C(1), C(2), C(3), I, J
   DoEvents
   Me.Caption = "Loading.. (" & I & "/" & W & ")"
  Next J
 Next I
 
 MsgBox UBound(Grass) * 2 & " polygons loaded" & " in " & UBound(Grass) & " objects"
 
 picHeight.Visible = False
 Set picHeight = Nothing
  
End Function

Public Sub Shutdown()
On Error Resume Next

 Keyboard.Destroy
 D3D.EndRender

 Set D3D = Nothing
 Erase Grass
 Set Tree = Nothing
 Set James = Nothing
 
 CursorShow
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Shutdown
End Sub

Public Function JoinPath(ByVal sPath As String, ByVal sFilename As String)
 If Left(sFilename, 1) = "\" Then sFilename = Mid(sFilename, 2, Len(sFilename))

 If Right(sPath, 1) = "\" Then
  JoinPath = sPath & sFilename
 Else
  JoinPath = sPath & "\" & sFilename
 End If
End Function

Public Sub CreateGrass(ByVal PY1 As Single, ByVal PY2 As Single, _
                       ByVal PY3 As Single, ByVal PY4 As Single, _
                       ByVal X As Single, ByVal Z As Single)

On Error Resume Next

Dim UGrass As Long
UGrass = UBound(Grass) + 1

If Err.Number = 9 Then UGrass = 0

If UGrass = 0 Then
 ReDim Grass(0)
Else
 ReDim Preserve Grass(UGrass)
End If

Set Grass(UGrass) = New clsRectangle8

With Grass(UGrass)

 .SetDevice D3D

 .SetVertexData ivX, qv1, X
 .SetVertexData ivY, qv1, PY1
 .SetVertexData ivZ, qv1, Z
 .SetVertexData ivC, qv1, RGB(0, PY1 * 10, 0)
 .SetVertexData ivU, qv1, 0
 .SetVertexData ivV, qv1, 0
 
 .SetVertexData ivX, qv2, X + UnitWidth
 .SetVertexData ivY, qv2, PY2
 .SetVertexData ivZ, qv2, Z
 .SetVertexData ivC, qv2, RGB(0, PY2 * 10, 0)
 .SetVertexData ivU, qv2, 1
 .SetVertexData ivV, qv2, 0
 
 .SetVertexData ivX, qv3, X
 .SetVertexData ivY, qv3, PY3
 .SetVertexData ivZ, qv3, Z + UnitHeight
 .SetVertexData ivC, qv3, RGB(0, PY3 * 10, 0)
 .SetVertexData ivU, qv3, 0
 .SetVertexData ivV, qv3, 1
 
 .SetVertexData ivX, qv4, X + UnitWidth
 .SetVertexData ivY, qv4, PY4
 .SetVertexData ivZ, qv4, Z + UnitHeight
 .SetVertexData ivC, qv4, RGB(0, PY4 * 10, 0)
 .SetVertexData ivU, qv4, 1
 .SetVertexData ivV, qv4, 1
 
 .Texture = JoinPath(App.Path, "Ground.bmp")
 
 .UseTexture = True
 .Render = True
 
 .Create
 
End With

End Sub
