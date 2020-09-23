VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirect3D8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim DX As New DirectX8
Dim D3DX As New D3DX8
Dim D3D As Direct3D8
Dim D3DDevice As Direct3DDevice8

Const FrameSkip As Single = 1
Dim FrameLength As Single

Dim WithEvents Game As clsGame
Attribute Game.VB_VarHelpID = -1
Dim VBuffer As Direct3DVertexBuffer8

Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Const Pi As Single = 3.141592653

Private Type Camera
 Angle As Single
 Pitch As Single
 X As Single
 Y As Single
 Z As Single
End Type

Private Type CustomVertex
    Position As D3DVECTOR
    Color As Long
    U As Single
    V As Single
End Type

Dim Camera As Camera
Dim CustomVertex As CustomVertex

Dim bFog As Boolean

 Dim matView As D3DMATRIX
 Dim matRotation As D3DMATRIX
 Dim matPitch As D3DMATRIX
 Dim matLook As D3DMATRIX
 Dim matPos As D3DMATRIX
 Dim matWorld As D3DMATRIX
 Dim matProj As D3DMATRIX

Event Rendering()
Event FinishRender(ByVal FPS As Long, ByVal MS As Long)
Event PreRender()

Public Property Get WorldMat() As D3DMATRIX
 WorldMat = matWorld
End Property

Public Function InitD3D(ByVal hWnd As Long, _
Optional ByVal Width As Long = 640, Optional ByVal Height As Long = 480, _
Optional ByVal Fullscreen As Boolean = False, _
Optional ByVal Hardware As Boolean = True) As Boolean
On Error Resume Next

Set D3D = DX.Direct3DCreate
If D3D Is Nothing Then Exit Function

Dim Mode As D3DDISPLAYMODE
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode

Dim d3dpp As D3DPRESENT_PARAMETERS

With d3dpp

If Fullscreen = True Then
 .Windowed = 0
Else
 .Windowed = 1
End If

 .BackBufferWidth = Width
 .BackBufferHeight = Height
 .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
 .BackBufferFormat = Mode.Format
 .BackBufferCount = 1
 .AutoDepthStencilFormat = D3DFMT_D16
 .EnableAutoDepthStencil = 1

End With

If Hardware = True Then
 Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                     hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
Else
 Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, _
                     hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
End If

If D3DDevice Is Nothing Then Exit Function

'I put this here insted of in the Render loop to prevent
'un-needed recalculations in Direct3D's engine

With D3DDevice

 .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
 .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
 
 .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE 'Use a texture-filter (looks better)
 .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE 'Use a texture-filter (looks better)
 
 .SetVertexShader D3DFVF_CUSTOMVERTEX
 
 .SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, 0
 .SetRenderState D3DRS_CULLMODE, D3DCULL_CW
 .SetRenderState D3DRS_ZENABLE, 1
 .SetRenderState D3DRS_LIGHTING, 0
 
 .SetRenderState D3DRS_ALPHAREF, &H8&
 .SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
 .SetRenderState D3DRS_ALPHATESTENABLE, 1
 
 .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID

End With

InitD3D = True
End Function

Private Function CreateVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
 CreateVector.X = X
 CreateVector.Y = Y
 CreateVector.Z = Z
End Function

Public Sub SetupMatrices()

 D3DXMatrixIdentity matWorld
 D3DDevice.SetTransform D3DTS_WORLD, matWorld

 D3DXMatrixIdentity matView
 D3DXMatrixIdentity matPos
 D3DXMatrixIdentity matRotation
 D3DXMatrixIdentity matLook
 
 D3DXMatrixRotationY matRotation, Camera.Angle
 D3DXMatrixRotationX matPitch, Camera.Pitch
 
 D3DXMatrixMultiply matLook, matRotation, matPitch
 
 D3DXMatrixTranslation matPos, -Camera.X, -Camera.Y, -Camera.Z
 
 D3DXMatrixMultiply matView, matPos, matLook
 
 D3DDevice.SetTransform D3DTS_VIEW, matView
 
 D3DXMatrixPerspectiveFovLH matProj, Pi / 3, 1, 1, 10000
 
 D3DDevice.SetTransform D3DTS_PROJECTION, matProj

End Sub

Public Property Get LookMat() As D3DMATRIX
 LookMat = matLook
End Property

Public Property Get RotMat() As D3DMATRIX
 RotMat = matRotation
End Property

Public Property Get PosMat() As D3DMATRIX
 PosMat = matPos
End Property

Public Property Get ViewMat() As D3DMATRIX
 ViewMat = matView
End Property

Public Property Get CamX() As Single
 CamX = Camera.X
End Property

Public Property Get CamY() As Single
 CamY = Camera.Y
End Property

Public Property Get CamZ() As Single
 CamZ = Camera.Z
End Property

Public Property Let CamX(ByVal sValue As Single)
 Camera.X = sValue
End Property

Public Property Let CamY(ByVal sValue As Single)
 Camera.Y = sValue
End Property

Public Property Let CamZ(ByVal sValue As Single)
 Camera.Z = sValue
End Property

Public Property Get CamAngle() As Single
 CamAngle = Camera.Angle
End Property

Public Property Get CamPitch() As Single
 CamPitch = Camera.Pitch
End Property

Public Property Let CamAngle(ByVal sValue As Single)
 Camera.Angle = sValue
End Property

Public Property Let CamPitch(ByVal sValue As Single)
 Camera.Pitch = sValue
End Property

Public Function BeginRender()
If Game Is Nothing Then Set Game = New clsGame
 Game.StartGame MaxFPS
End Function

Public Function EndRender()
If Game Is Nothing Then Set Game = New clsGame
 Game.EndGame
End Function

Private Sub Render()
RaiseEvent PreRender

'FrameLength = FrameLength + 1
'If FrameLength = FrameSkip Then Exit Sub

On Error Resume Next

If D3DDevice Is Nothing Then Exit Sub

Dim V As CustomVertex
Dim SizeOfVertex As Long
SizeOfVertex = Len(V)

SetupMatrices

D3DDevice.BeginScene

RaiseEvent Rendering

D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

RaiseEvent FinishRender(Game.FPS, Game.MS)
End Sub

Public Function Clear(ByVal Color As Long)
 D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Color, 1#, 0
End Function

Public Function MoveForward(ByVal Distance As Single)

Dim AngleConv As Single
AngleConv = Radian(360) - Camera.Angle

 Camera.X = Camera.X + (Sin(AngleConv) * Distance)
 Camera.Z = Camera.Z + (Cos(AngleConv) * Distance)

End Function

Public Function MoveBackward(ByVal Distance As Single)

Dim AngleConv As Single
AngleConv = Radian(360) - Camera.Angle

 Camera.X = Camera.X - (Sin(AngleConv) * Distance)
 Camera.Z = Camera.Z - (Cos(AngleConv) * Distance)

End Function

Public Property Get Device() As Direct3DDevice8
 Set Device = D3DDevice
End Property

Private Sub Game_GameUpdate(ByVal FPS As Long, ByVal MS As Long)
 Render
End Sub

'I would like to thank DirectX4VB.com for these 2 functions

Private Function FloatToDWord(f As Single) As Long
    'this function packs a 32bit floating point number
    'into a 32bit integer number; quite slow - dont overuse.
    'DXCopyMemory or CopyMemory() (win32 api) would
    'probably be faster...
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FloatToDWord = l
End Function

'//To check for Range Based Fog support (it looks better!!)
Private Function CheckForRangeBasedFog(adapter As Byte) As Boolean
On Local Error Resume Next
    Dim DX As New DirectX8
    Dim D3D As Direct3D8
    Dim Caps As D3DCAPS8
    
    Set D3D = DX.Direct3DCreate
    
    D3D.GetDeviceCaps adapter - 1, D3DDEVTYPE_HAL, Caps
    
    If Caps.RasterCaps And D3DPRASTERCAPS_FOGRANGE Then
        CheckForRangeBasedFog = True
    Else
        CheckForRangeBasedFog = False
    End If
End Function

Public Property Let FogEnable(ByVal bValue As Boolean)

 bFog = bValue
 
If bValue = True Then
 D3DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_NONE 'dont use table fog
 D3DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR 'use standard linear fog
 D3DDevice.SetRenderState D3DRS_RANGEFOGENABLE, CheckForRangeBasedFog(D3DADAPTER_DEFAULT) 'enable range based fog, hw dependent
 D3DDevice.SetRenderState D3DRS_FOGENABLE, 1 'set to 0 to disable
Else
 D3DDevice.SetRenderState D3DRS_FOGENABLE, 0 'set to 0 to disable
End If
 
End Property

Public Property Get FogEnable() As Boolean
 FogEnable = bFog
End Property

Public Property Let FogStart(ByVal bValue As Single)
 D3DDevice.SetRenderState D3DRS_FOGSTART, FloatToDWord(bValue)
End Property

Public Property Let FogEnd(ByVal bValue As Single)
 D3DDevice.SetRenderState D3DRS_FOGEND, FloatToDWord(bValue)
End Property

Public Property Let FogColor(ByVal lValue As Long)
 D3DDevice.SetRenderState D3DRS_FOGCOLOR, lValue
End Property

