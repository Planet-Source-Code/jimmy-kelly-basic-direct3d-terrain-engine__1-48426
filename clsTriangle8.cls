VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTriangle8"
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

Private Type Polygon
 X(2) As Single
 Y(2) As Single
 Z(2) As Single
 U(2) As Single
 V(2) As Single
 C(2) As Long
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

Public Function SetVertexData(ByRef DataType As InfoVertex, ByRef Index As TriVertex, ByVal sValue As Variant)

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

Public Function GetVertexData(ByRef DataType As InfoVertex, ByRef Index As TriVertex)
 
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
    Vertices(2).Position = CreateVector(Polygon.X(2), Polygon.Y(2), Polygon.Z(2))

    Vertices(0).Color = Polygon.C(0)
    Vertices(1).Color = Polygon.C(1)
    Vertices(2).Color = Polygon.C(2)
    Vertices(2).Color = Polygon.C(2)
    Vertices(0).U = Polygon.U(0)
    Vertices(1).U = Polygon.U(1)
    Vertices(2).U = Polygon.U(2)
    Vertices(2).U = Polygon.U(2)
    Vertices(0).V = Polygon.V(0)
    Vertices(1).V = Polygon.V(1)
    Vertices(2).V = Polygon.V(2)
    Vertices(2).V = Polygon.V(2)

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

  D3D.Device.SetStreamSource 0, VertexBuffer, SizeOfVertex
  D3D.Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
  
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
