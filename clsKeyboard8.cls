VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyboard8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim DX As New DirectX8
Dim DI As DirectInput8
Dim Keyboard As DirectInputDevice8
Dim KeyState As DIKEYBOARDSTATE

Dim WithEvents Game As clsGame
Attribute Game.VB_VarHelpID = -1

Event Keydown(ByVal KeyCode As CONST_DIKEYFLAGS)

Public Function Create(ByVal hWnd As Long)
On Error Resume Next
 Set DI = DX.DirectInputCreate()
 Set Keyboard = DI.CreateDevice("GUID_SysKeyboard")
 Keyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
 Keyboard.SetCooperativeLevel hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
 Keyboard.Acquire
End Function

Public Function Destroy()
On Error Resume Next
 Keyboard.Unacquire
 Set DI = Nothing
 Set Keyboard = Nothing
End Function

Public Function GetKeyState(ByRef KeyCode As CONST_DIKEYFLAGS) As Boolean

Keyboard.GetDeviceStateKeyboard KeyState

If KeyState.Key(KeyCode) <> 0 Then
 GetKeyState = True
Else
 GetKeyState = False
End If

End Function

Public Function BeginMonitor()
On Error Resume Next
 Set Game = New clsGame
 Game.StartGame 30
End Function

Public Function EndMonitor()
On Error Resume Next
 Game.EndGame
 Set Game = Nothing
End Function

Private Function TestKeys()

If GetKeyState(DIK_0) Then RaiseEvent Keydown(DIK_0)
If GetKeyState(DIK_1) Then RaiseEvent Keydown(DIK_1)
If GetKeyState(DIK_2) Then RaiseEvent Keydown(DIK_2)
If GetKeyState(DIK_3) Then RaiseEvent Keydown(DIK_3)
If GetKeyState(DIK_4) Then RaiseEvent Keydown(DIK_4)
If GetKeyState(DIK_5) Then RaiseEvent Keydown(DIK_5)
If GetKeyState(DIK_6) Then RaiseEvent Keydown(DIK_6)
If GetKeyState(DIK_7) Then RaiseEvent Keydown(DIK_7)
If GetKeyState(DIK_8) Then RaiseEvent Keydown(DIK_8)
If GetKeyState(DIK_9) Then RaiseEvent Keydown(DIK_9)
If GetKeyState(DIK_A) Then RaiseEvent Keydown(DIK_A)
If GetKeyState(DIK_ABNT_C1) Then RaiseEvent Keydown(DIK_ABNT_C1)
If GetKeyState(DIK_ABNT_C2) Then RaiseEvent Keydown(DIK_ABNT_C2)
If GetKeyState(DIK_ADD) Then RaiseEvent Keydown(DIK_ADD)
If GetKeyState(DIK_APOSTROPHE) Then RaiseEvent Keydown(DIK_APOSTROPHE)
If GetKeyState(DIK_APPS) Then RaiseEvent Keydown(DIK_APPS)
If GetKeyState(DIK_AT) Then RaiseEvent Keydown(DIK_AT)
If GetKeyState(DIK_AX) Then RaiseEvent Keydown(DIK_AX)
If GetKeyState(DIK_B) Then RaiseEvent Keydown(DIK_B)
If GetKeyState(DIK_BACK) Then RaiseEvent Keydown(DIK_BACK)
If GetKeyState(DIK_BACKSLASH) Then RaiseEvent Keydown(DIK_BACKSLASH)
If GetKeyState(DIK_BACKSPACE) Then RaiseEvent Keydown(DIK_BACKSPACE)
If GetKeyState(DIK_C) Then RaiseEvent Keydown(DIK_C)
If GetKeyState(DIK_CALCULATOR) Then RaiseEvent Keydown(DIK_CALCULATOR)
If GetKeyState(DIK_CAPITAL) Then RaiseEvent Keydown(DIK_CAPITAL)
If GetKeyState(DIK_CAPSLOCK) Then RaiseEvent Keydown(DIK_CAPSLOCK)
If GetKeyState(DIK_CIRCUMFLEX) Then RaiseEvent Keydown(DIK_CIRCUMFLEX)
If GetKeyState(DIK_COLON) Then RaiseEvent Keydown(DIK_COLON)
If GetKeyState(DIK_COMMA) Then RaiseEvent Keydown(DIK_COMMA)
If GetKeyState(DIK_CONVERT) Then RaiseEvent Keydown(DIK_CONVERT)
If GetKeyState(DIK_D) Then RaiseEvent Keydown(DIK_D)
If GetKeyState(DIK_DECIMAL) Then RaiseEvent Keydown(DIK_DECIMAL)
If GetKeyState(DIK_DELETE) Then RaiseEvent Keydown(DIK_DELETE)
If GetKeyState(DIK_DIVIDE) Then RaiseEvent Keydown(DIK_DIVIDE)
If GetKeyState(DIK_DOWN) Then RaiseEvent Keydown(DIK_DOWN)
If GetKeyState(DIK_DOWNARROW) Then RaiseEvent Keydown(DIK_DOWNARROW)
If GetKeyState(DIK_E) Then RaiseEvent Keydown(DIK_E)
If GetKeyState(DIK_END) Then RaiseEvent Keydown(DIK_END)
If GetKeyState(DIK_EQUALS) Then RaiseEvent Keydown(DIK_EQUALS)
If GetKeyState(DIK_ESCAPE) Then RaiseEvent Keydown(DIK_ESCAPE)
If GetKeyState(DIK_F) Then RaiseEvent Keydown(DIK_F)
If GetKeyState(DIK_F1) Then RaiseEvent Keydown(DIK_F1)
If GetKeyState(DIK_F2) Then RaiseEvent Keydown(DIK_F2)
If GetKeyState(DIK_F3) Then RaiseEvent Keydown(DIK_F3)
If GetKeyState(DIK_F4) Then RaiseEvent Keydown(DIK_F4)
If GetKeyState(DIK_F5) Then RaiseEvent Keydown(DIK_F5)
If GetKeyState(DIK_F6) Then RaiseEvent Keydown(DIK_F6)
If GetKeyState(DIK_F7) Then RaiseEvent Keydown(DIK_F7)
If GetKeyState(DIK_F8) Then RaiseEvent Keydown(DIK_F8)
If GetKeyState(DIK_F9) Then RaiseEvent Keydown(DIK_F9)
If GetKeyState(DIK_F10) Then RaiseEvent Keydown(DIK_F10)
If GetKeyState(DIK_F11) Then RaiseEvent Keydown(DIK_F11)
If GetKeyState(DIK_F12) Then RaiseEvent Keydown(DIK_F12)
If GetKeyState(DIK_G) Then RaiseEvent Keydown(DIK_G)
If GetKeyState(DIK_GRAVE) Then RaiseEvent Keydown(DIK_GRAVE)
If GetKeyState(DIK_H) Then RaiseEvent Keydown(DIK_H)
If GetKeyState(DIK_HOME) Then RaiseEvent Keydown(DIK_HOME)
If GetKeyState(DIK_I) Then RaiseEvent Keydown(DIK_I)
If GetKeyState(DIK_INSERT) Then RaiseEvent Keydown(DIK_INSERT)
If GetKeyState(DIK_J) Then RaiseEvent Keydown(DIK_J)
If GetKeyState(DIK_K) Then RaiseEvent Keydown(DIK_K)
If GetKeyState(DIK_KANA) Then RaiseEvent Keydown(DIK_KANA)
If GetKeyState(DIK_KANJI) Then RaiseEvent Keydown(DIK_KANJI)
If GetKeyState(DIK_L) Then RaiseEvent Keydown(DIK_L)
If GetKeyState(DIK_LALT) Then RaiseEvent Keydown(DIK_LALT)
If GetKeyState(DIK_LBRACKET) Then RaiseEvent Keydown(DIK_LBRACKET)
If GetKeyState(DIK_LCONTROL) Then RaiseEvent Keydown(DIK_LCONTROL)
If GetKeyState(DIK_LEFT) Then RaiseEvent Keydown(DIK_LEFT)
If GetKeyState(DIK_LEFTARROW) Then RaiseEvent Keydown(DIK_LEFTARROW)
If GetKeyState(DIK_LMENU) Then RaiseEvent Keydown(DIK_LMENU)
If GetKeyState(DIK_LSHIFT) Then RaiseEvent Keydown(DIK_LSHIFT)
If GetKeyState(DIK_LWIN) Then RaiseEvent Keydown(DIK_LWIN)
If GetKeyState(DIK_M) Then RaiseEvent Keydown(DIK_M)
If GetKeyState(DIK_MAIL) Then RaiseEvent Keydown(DIK_MAIL)
If GetKeyState(DIK_MEDIASELECT) Then RaiseEvent Keydown(DIK_MEDIASELECT)
If GetKeyState(DIK_MEDIASTOP) Then RaiseEvent Keydown(DIK_MEDIASTOP)
If GetKeyState(DIK_MINUS) Then RaiseEvent Keydown(DIK_MINUS)
If GetKeyState(DIK_MULTIPLY) Then RaiseEvent Keydown(DIK_MULTIPLY)
If GetKeyState(DIK_MUTE) Then RaiseEvent Keydown(DIK_MUTE)
If GetKeyState(DIK_MYCOMPUTER) Then RaiseEvent Keydown(DIK_MYCOMPUTER)
If GetKeyState(DIK_N) Then RaiseEvent Keydown(DIK_N)
If GetKeyState(DIK_NEXT) Then RaiseEvent Keydown(DIK_NEXT)
If GetKeyState(DIK_NEXTTRACK) Then RaiseEvent Keydown(DIK_NEXTTRACK)
If GetKeyState(DIK_NOCONVERT) Then RaiseEvent Keydown(DIK_NOCONVERT)
If GetKeyState(DIK_NUMLOCK) Then RaiseEvent Keydown(DIK_NUMLOCK)
If GetKeyState(DIK_NUMPAD0) Then RaiseEvent Keydown(DIK_NUMPAD0)
If GetKeyState(DIK_NUMPAD1) Then RaiseEvent Keydown(DIK_NUMPAD1)
If GetKeyState(DIK_NUMPAD2) Then RaiseEvent Keydown(DIK_NUMPAD2)
If GetKeyState(DIK_NUMPAD3) Then RaiseEvent Keydown(DIK_NUMPAD3)
If GetKeyState(DIK_NUMPAD4) Then RaiseEvent Keydown(DIK_NUMPAD4)
If GetKeyState(DIK_NUMPAD5) Then RaiseEvent Keydown(DIK_NUMPAD5)
If GetKeyState(DIK_NUMPAD6) Then RaiseEvent Keydown(DIK_NUMPAD6)
If GetKeyState(DIK_NUMPAD7) Then RaiseEvent Keydown(DIK_NUMPAD7)
If GetKeyState(DIK_NUMPAD8) Then RaiseEvent Keydown(DIK_NUMPAD8)
If GetKeyState(DIK_NUMPAD9) Then RaiseEvent Keydown(DIK_NUMPAD9)
If GetKeyState(DIK_NUMPADCOMMA) Then RaiseEvent Keydown(DIK_NUMPADCOMMA)
If GetKeyState(DIK_NUMPADENTER) Then RaiseEvent Keydown(DIK_NUMPADENTER)
If GetKeyState(DIK_NUMPADEQUALS) Then RaiseEvent Keydown(DIK_NUMPADEQUALS)
If GetKeyState(DIK_NUMPADMINUS) Then RaiseEvent Keydown(DIK_NUMPADMINUS)
If GetKeyState(DIK_NUMPADPERIOD) Then RaiseEvent Keydown(DIK_NUMPADPERIOD)
If GetKeyState(DIK_NUMPADPLUS) Then RaiseEvent Keydown(DIK_NUMPADPLUS)
If GetKeyState(DIK_NUMPADSLASH) Then RaiseEvent Keydown(DIK_NUMPADSLASH)
If GetKeyState(DIK_NUMPADSTAR) Then RaiseEvent Keydown(DIK_NUMPADSTAR)
If GetKeyState(DIK_O) Then RaiseEvent Keydown(DIK_O)
If GetKeyState(DIK_OEM_102) Then RaiseEvent Keydown(DIK_OEM_102)
If GetKeyState(DIK_P) Then RaiseEvent Keydown(DIK_P)
If GetKeyState(DIK_PAUSE) Then RaiseEvent Keydown(DIK_PAUSE)
If GetKeyState(DIK_PERIOD) Then RaiseEvent Keydown(DIK_PERIOD)
If GetKeyState(DIK_PGDN) Then RaiseEvent Keydown(DIK_PGDN)
If GetKeyState(DIK_PGUP) Then RaiseEvent Keydown(DIK_PGUP)
If GetKeyState(DIK_PLAYPAUSE) Then RaiseEvent Keydown(DIK_PLAYPAUSE)
If GetKeyState(DIK_POWER) Then RaiseEvent Keydown(DIK_POWER)
If GetKeyState(DIK_PREVTRACK) Then RaiseEvent Keydown(DIK_PREVTRACK)
If GetKeyState(DIK_PRIOR) Then RaiseEvent Keydown(DIK_PRIOR)
If GetKeyState(DIK_Q) Then RaiseEvent Keydown(DIK_Q)
If GetKeyState(DIK_R) Then RaiseEvent Keydown(DIK_R)
If GetKeyState(DIK_RALT) Then RaiseEvent Keydown(DIK_RALT)
If GetKeyState(DIK_RBRACKET) Then RaiseEvent Keydown(DIK_RBRACKET)
If GetKeyState(DIK_RCONTROL) Then RaiseEvent Keydown(DIK_RCONTROL)
If GetKeyState(DIK_RIGHT) Then RaiseEvent Keydown(DIK_RIGHT)
If GetKeyState(DIK_RIGHTARROW) Then RaiseEvent Keydown(DIK_RIGHTARROW)
If GetKeyState(DIK_RMENU) Then RaiseEvent Keydown(DIK_RMENU)
If GetKeyState(DIK_RSHIFT) Then RaiseEvent Keydown(DIK_RSHIFT)
If GetKeyState(DIK_RWIN) Then RaiseEvent Keydown(DIK_RWIN)
If GetKeyState(DIK_S) Then RaiseEvent Keydown(DIK_S)
If GetKeyState(DIK_SCROLL) Then RaiseEvent Keydown(DIK_SCROLL)
If GetKeyState(DIK_SEMICOLON) Then RaiseEvent Keydown(DIK_SEMICOLON)
If GetKeyState(DIK_SLASH) Then RaiseEvent Keydown(DIK_SLASH)
If GetKeyState(DIK_SLEEP) Then RaiseEvent Keydown(DIK_SLEEP)
If GetKeyState(DIK_SPACE) Then RaiseEvent Keydown(DIK_SPACE)
If GetKeyState(DIK_STOP) Then RaiseEvent Keydown(DIK_STOP)
If GetKeyState(DIK_SUBTRACT) Then RaiseEvent Keydown(DIK_SUBTRACT)
If GetKeyState(DIK_SYSRQ) Then RaiseEvent Keydown(DIK_SYSRQ)
If GetKeyState(DIK_T) Then RaiseEvent Keydown(DIK_T)
If GetKeyState(DIK_TAB) Then RaiseEvent Keydown(DIK_TAB)
If GetKeyState(DIK_U) Then RaiseEvent Keydown(DIK_U)
If GetKeyState(DIK_UNDERLINE) Then RaiseEvent Keydown(DIK_UNDERLINE)
If GetKeyState(DIK_UNLABELED) Then RaiseEvent Keydown(DIK_UNLABELED)
If GetKeyState(DIK_UP) Then RaiseEvent Keydown(DIK_UP)
If GetKeyState(DIK_UPARROW) Then RaiseEvent Keydown(DIK_UPARROW)
If GetKeyState(DIK_V) Then RaiseEvent Keydown(DIK_V)
If GetKeyState(DIK_VOLUMEDOWN) Then RaiseEvent Keydown(DIK_VOLUMEDOWN)
If GetKeyState(DIK_VOLUMEUP) Then RaiseEvent Keydown(DIK_VOLUMEUP)
If GetKeyState(DIK_W) Then RaiseEvent Keydown(DIK_W)
If GetKeyState(DIK_WAKE) Then RaiseEvent Keydown(DIK_WAKE)
If GetKeyState(DIK_WEBBACK) Then RaiseEvent Keydown(DIK_WEBBACK)
If GetKeyState(DIK_WEBFAVORITES) Then RaiseEvent Keydown(DIK_WEBFAVORITES)
If GetKeyState(DIK_WEBFORWARD) Then RaiseEvent Keydown(DIK_WEBFORWARD)
If GetKeyState(DIK_WEBHOME) Then RaiseEvent Keydown(DIK_WEBHOME)
If GetKeyState(DIK_WEBREFRESH) Then RaiseEvent Keydown(DIK_WEBREFRESH)
If GetKeyState(DIK_WEBSEARCH) Then RaiseEvent Keydown(DIK_WEBSEARCH)
If GetKeyState(DIK_WEBSTOP) Then RaiseEvent Keydown(DIK_WEBSTOP)
If GetKeyState(DIK_X) Then RaiseEvent Keydown(DIK_X)
If GetKeyState(DIK_Y) Then RaiseEvent Keydown(DIK_Y)
If GetKeyState(DIK_YEN) Then RaiseEvent Keydown(DIK_YEN)
If GetKeyState(DIK_Z) Then RaiseEvent Keydown(DIK_Z)

End Function

Private Sub Game_GameUpdate(ByVal FPS As Long, ByVal MS As Long)
 TestKeys
End Sub
