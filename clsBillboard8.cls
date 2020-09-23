VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillboard8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim VertexBuffer As Direct3DVertexBuffer8
Dim TextureData As Direct3DTexture8

Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Const Pi As Single = 3.141592653

Dim bTexture As Boolean
Dim bChanged As Boolean
Dim bRender As Boolean
Dim bAlpha As Boolean
Dim bPitch As Boolean
Dim bAngle As Boolean

Dim cAngle As Single
Dim cPitch As Single

Private Type Polygon
 X(3) As Single
 Y(3) As Single
 Z(3) As Single
 U(3) As Single
 V(3) As Single
 C(3) As Long
 Texture As String
End Type

Dim D3DX As New D3DX8
Dim WithEvents D3D As clsDirect3D8
Attribute D3D.VB_VarHelpID = -1

Private Type CustomVertex
    Position As D3DVECTOR
    Color As Long
    U As Single
    V As Single
End Type

Dim sX As Single
Dim sY As Single
Dim sZ As Single
Dim sW As Single
Dim sH As Single

Dim CustomVertex As CustomVertex
Dim Polygon As Polygon

Event BeginDraw()
Event EndDraw()

Public Property Let Render(ByVal bValue As Boolean)
 bRender = bValue
End Property

Public Property Get Render() As Boolean
 Render = bRender
End Property

Public Property Let UseTexture(ByVal bValue As Boolean)
 bTexture = bValue
End Property

Public Property Get UseTexture() As Boolean
 UseTexture = bTexture
End Property

Public Property Let FacePitch(ByVal bValue As Boolean)
 bPitch = bValue
End Property

Public Property Get FacePitch() As Boolean
 FacePitch = bPitch
End Property

Public Property Let FaceAngle(ByVal bValue As Boolean)
 bAngle = bValue
End Property

Public Property Get FaceAngle() As Boolean
 FaceAngle = bAngle
End Property

Public Property Let Angle(ByVal sValue As Single)
 cAngle = sValue
End Property

Public Property Get Angle() As Single
 Angle = cAngle
End Property

Public Property Let Pitch(ByVal sValue As Single)
 cPitch = sValue
End Property

Public Property Get Pitch() As Single
 Pitch = cPitch
End Property

Public Property Let Texture(ByVal sValue As String)
' Set TextureData = D3DX.CreateTextureFromFile(D3D.Device, sValue)

 Set TextureData = D3DX.CreateTextureFromFileEx( _
 D3D.Device, sValue, D3DX_DEFAULT, D3DX_DEFAULT, _
 D3DX_DEFAULT, 0, D3DFMT_DXT5, D3DPOOL_MANAGED, _
 D3DX_FILTER_NONE, D3DX_FILTER_NONE, RGB(255, 0, 255), _
 ByVal 0, ByVal 0)
 
 bChanged = True
End Property

Public Property Get Texture() As String
 Texture = Polygon.Texture
End Property

Public Sub SetDevice(ByRef clsD3D As clsDirect3D8)
 Set D3D = clsD3D
 D3D.Device.SetVertexShader D3DFVF_CUSTOMVERTEX
End Sub

Public Property Let Width(ByVal sValue As Single)
 sW = sValue
 bChanged = True
End Property

Public Property Get Width() As Single
 Width = sW
End Property

Public Property Let Height(ByVal sValue As Single)
 sH = sValue
 bChanged = True
End Property

Public Property Get Height() As Single
 Height = sH
End Property

Public Property Let X(ByVal sValue As Single)
 sX = sValue
 bChanged = True
End Property

Public Property Get X() As Single
 X = sX
End Property

Public Property Let Y(ByVal sValue As Single)
 sY = sValue
 bChanged = True
End Property

Public Property Get Y() As Single
 Y = sY
End Property

Public Property Let Z(ByVal sValue As Single)
 sZ = sValue
 bChanged = True
End Property

Public Property Get Z() As Single
 Z = sZ
End Property

Private Function SetVertexData(ByRef DataType As InfoVertex, ByRef Index As QuadVertex, ByVal sValue As Variant)

Select Case DataType
 Case 1
  Polygon.X(Index) = sValue
 Case 2
  Polygon.Y(Index) = sValue
 Case 3
  Polygon.Z(Index) = sValue
 Case 4
  Polygon.U(Index) = sValue
 Case 5
  Polygon.V(Index) = sValue
 Case 6
  Polygon.C(Index) = sValue
 Case Else
  bChanged = False: Exit Function
End Select

 bChanged = True

End Function

Private Function GetVertexData(ByRef DataType As InfoVertex, ByRef Index As QuadVertex)
 
Select Case DataType
 Case 1
  GetVertexData = Polygon.X(Index)
 Case 2
  GetVertexData = Polygon.Y(Index)
 Case 3
  GetVertexData = Polygon.Z(Index)
 Case 4
  GetVertexData = Polygon.U(Index)
 Case 5
  GetVertexData = Polygon.V(Index)
 Case 6
  GetVertexData = Polygon.C(Index)
End Select

End Function

Private Sub Class_Initialize()

bChanged = True
bAlpha = False

End Sub

Public Property Let Alpha(ByVal bValue As Boolean)
 bAlpha = bValue
End Property

Public Property Get Alpha() As Boolean
 Alpha = bAlpha
End Property

Public Function Create()

    'create an array to hold the vertex values temporarily, until added to buffer
    Dim Vertices(0 To 3) As CustomVertex
    Dim VertexSizeInBytes As Long
    'get the size of a vertex
    VertexSizeInBytes = Len(Vertices(0))

    Vertices(0).Position = CreateVector(Polygon.X(0), Polygon.Y(0), Polygon.Z(0))
    Vertices(1).Position = CreateVector(Polygon.X(1), Polygon.Y(1), Polygon.Z(1))
    Vertices(2).Position = CreateVector(Polygon.X(2), Polygon.Y(2), Polygon.Z(2))
    Vertices(3).Position = CreateVector(Polygon.X(3), Polygon.Y(3), Polygon.Z(3))

    Vertices(0).Color = Polygon.C(0)
    Vertices(1).Color = Polygon.C(1)
    Vertices(2).Color = Polygon.C(2)
    Vertices(3).Color = Polygon.C(3)
    Vertices(0).U = Polygon.U(0)
    Vertices(1).U = Polygon.U(1)
    Vertices(2).U = Polygon.U(2)
    Vertices(3).U = Polygon.U(3)
    Vertices(0).V = Polygon.V(0)
    Vertices(1).V = Polygon.V(1)
    Vertices(2).V = Polygon.V(2)
    Vertices(3).V = Polygon.V(3)

    Set VertexBuffer = Nothing

    ' Create the vertex buffer.
    Set VertexBuffer = D3D.Device.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If VertexBuffer Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData VertexBuffer, 0, VertexSizeInBytes * 4, 0, Vertices(0)

End Function

Public Function Draw()

On Error Resume Next

If D3D Is Nothing Then Exit Function
 
  If bTexture = True Then
   D3D.Device.SetTexture 0, TextureData
  Else
   D3D.Device.SetTexture 0, Nothing
  End If

 Dim V As CustomVertex
 Dim SizeOfVertex As Long

 SizeOfVertex = Len(V)
 
  D3D.Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
  D3D.Device.SetRenderState D3DRS_ALPHATESTENABLE, 0
 
If bAlpha = True Then
 D3D.Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
 D3D.Device.SetRenderState D3DRS_ALPHATESTENABLE, 1
End If

RaiseEvent BeginDraw

If bPitch = True Then cPitch = D3D.CamPitch
If bAngle = True Then cAngle = D3D.CamAngle

ChangeMe

 'Dim rot As D3DMATRIX
 'D3DXMatrixRotationY rot, -D3D.CamAngle
 'D3D.Device.SetTransform D3DTS_WORLD, rot

  D3D.Device.SetStreamSource 0, VertexBuffer, SizeOfVertex
  D3D.Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
  
 D3D.SetupMatrices
  
RaiseEvent EndDraw
  
If bAlpha = True Then
  D3D.Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
  D3D.Device.SetRenderState D3DRS_ALPHATESTENABLE, 0
End If
  
  D3D.Device.SetTexture 0, Nothing
 
End Function

Private Function CreateVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
 CreateVector.X = X
 CreateVector.Y = Y
 CreateVector.Z = Z
End Function

Private Sub D3D_Rendering()
 If bChanged = True Then Create
 If bRender = True Then Draw
End Sub

Private Sub ChangeMe()

'I gave up on Direct3D's crazy rotation billboard stuff
'so I wrote my own code for doing it and it works just as well.
'Another note in case you think this is slower I tested it over
'and over and I didn't notice and diffrence in framerate

SetVertexData ivY, qv1, sY + Cos(cPitch) * sH / 2
SetVertexData ivY, qv2, sY + Cos(cPitch) * sH / 2
SetVertexData ivY, qv3, sY - Cos(cPitch) * sH / 2
SetVertexData ivY, qv4, sY - Cos(cPitch) * sH / 2

SetVertexData ivX, qv1, sX - Cos(cAngle) * sW / 2
SetVertexData ivZ, qv1, sZ - Sin(cAngle) * sW / 2
SetVertexData ivX, qv3, sX - Cos(cAngle) * sW / 2
SetVertexData ivZ, qv3, sZ - Sin(cAngle) * sW / 2

SetVertexData ivX, qv2, sX + Cos(cAngle) * sW / 2
SetVertexData ivZ, qv2, sZ + Sin(cAngle) * sW / 2
SetVertexData ivX, qv4, sX + Cos(cAngle) * sW / 2
SetVertexData ivZ, qv4, sZ + Sin(cAngle) * sW / 2

SetVertexData ivU, qv1, 0
SetVertexData ivV, qv1, 0
SetVertexData ivC, qv1, vbWhite

SetVertexData ivU, qv2, 1
SetVertexData ivV, qv2, 0
SetVertexData ivC, qv2, vbWhite

SetVertexData ivU, qv3, 0
SetVertexData ivV, qv3, 1
SetVertexData ivC, qv3, vbWhite

SetVertexData ivU, qv4, 1
SetVertexData ivV, qv4, 1
SetVertexData ivC, qv4, vbWhite

End Sub
