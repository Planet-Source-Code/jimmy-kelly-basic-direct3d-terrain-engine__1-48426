VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Windows API Functions
Private Declare Function GetTickCount _
Lib "kernel32" () As Long

Dim bRunning As Boolean
Dim MSRate As Long
Dim FPSRate As Long
Dim MaxFrameRate As Long

Event GameUpdate(ByVal FPS As Long, ByVal MS As Long)

Public Property Get FPS() As Long
 FPS = FPSRate
End Property

Public Property Get MS() As Long
 MS = MSRate
End Property

Public Function StartGame(ByVal MaxFPS As Long)

On Error Resume Next

MaxFrameRate = MaxFPS

bRunning = True

Static lStartTime As Long
Static lCounter As Long
Static lNewTime As Long

'main game loop
Do While bRunning
 lCounter = GetTickCount() - lStartTime
  If lCounter > lNewTime Then
  Game_Update lCounter
  lNewTime = lCounter + 1000 / MaxFrameRate
 End If
 DoEvents
Loop

End Function

Public Function EndGame()
 bRunning = False
End Function

Private Sub Game_Update(ByVal MS As Long)

Static lTimer As Long
Static lStart As Long
Static lCounter As Long

'Start counting miliseconds
lStart = GetTickCount

'count the frames per second
If MS > lTimer + 1000 Then
 lStart = GetTickCount - lStart
 FPSRate = lCounter
 MSRate = lStart
 lTimer = MS
 lCounter = 0
Else
 lCounter = lCounter + 1
End If

RaiseEvent GameUpdate(FPSRate, MSRate)

Exit Sub

End Sub
