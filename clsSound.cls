VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" ( _
                            ByVal lpstrCommand As String, _
                            ByVal lpstrReturnString As String, _
                            ByVal uReturnLength As Long, _
                            ByVal hwndCallback As Long) As Long
                    
Private Declare PtrSafe Function mciGetErrorString Lib "winmm" Alias "mciGetErrorStringA" ( _
                            ByVal dwError As Long, _
                            ByVal lpstrBuffer As String, _
                            ByVal uLength As Long) As Long

Private pHasOpen As Boolean 'オープルされているか
Private pFile As String '音楽ファイル
Private pAlias As String '音楽エイリアス
Private pErrDescription As String 'APIエラーメッセージ
Private pCommand As String 'lpstrCommand作成用

'音楽ファイル
Public Property Get SoundFile() As String
    SoundFile = pFile
End Property
Public Property Let SoundFile(ByVal sFile As String)
    pFile = sFile
    '再生中等の場合もあるのでいったん停止させる
    On Error Resume Next
    Call StopSound
End Property

'エラーメッセージ
Public Property Get ErrDescription() As String
    ErrDescription = pErrDescription
End Property

'Openしているか
Public Property Get HasOpen() As Boolean
    HasOpen = pHasOpen
End Property

'mciSendStringのエラーメッセージ
Private Function GetMciError(ByVal aErrCode As Long, _
                             ByVal ErrRaise As Boolean) As String
    Dim buf As String * 256
    Dim rtnCode As Long
    rtnCode = mciGetErrorString(aErrCode, buf, Len(buf))
    If rtnCode = 0 Then
        pErrDescription = aErrCode & ":" & "不明なエラー"
    Else
        pErrDescription = aErrCode & ":" & Left(buf, InStr(buf, vbNullChar) - 1)
    End If
    If ErrRaise Then
        Err.Raise Number:=vbObjectError + 513, Description:=pErrDescription
    End If
End Function

'オープン：音楽ファイルのチェック兼用
Public Function OpenSound(Optional ByVal ErrRaise As Boolean = True) As Boolean
    OpenSound = True
  
    '現在の状態確認
    Dim rtnStatus As String
    rtnStatus = GetStatus
    Select Case rtnStatus
        Case "not ready"
            Exit Function '発生条件は不明
        Case "playing", "paused", "stopped"
            Exit Function
    End Select
  
    'オープン
    pAlias = "MySound" 'エイリアス作成
    Dim ret As Long
    pCommand = "open """ & SoundFile & """ alias " & pAlias
    ret = mciSendString(pCommand, "", 0, 0)
    pHasOpen = True
    If ret = 0 Then Exit Function
  
    'エラー
    Call GetMciError(ret, ErrRaise)
    OpenSound = False
    pHasOpen = False
End Function

'再生：引数は秒数指定
Public Sub Play(Optional ByVal aPosition As Long = 0)
    If Not pHasOpen Then Exit Sub
    If Not Me.OpenSound(False) Then Exit Sub
    If aPosition = 0 Then
        pCommand = "play " & pAlias
        Call mciSendString(pCommand, "", 0, 0)
    Else
        '指定位置から再生
        Call PlayPosition(aPosition)
    End If
End Sub

'指定位置再生：引数は秒数指定
Public Sub PlayPosition(ByVal aPosition As Long)
    If Not pHasOpen Then Exit Sub
    pCommand = "seek " & pAlias & " to " & aPosition * 1000
    Call mciSendString(pCommand, "", 0, 0)
    pCommand = "play " & pAlias
    Call mciSendString(pCommand, "", 0, 0)
End Sub

'指定秒数スキップ
Public Sub SkipPosition(ByVal aPostion As Long)
    If Not pHasOpen Then Exit Sub
    Dim curPos As Long
    curPos = GetPosition() * 200
    curPos = curPos + (aPostion * 200)
    pCommand = "seek " & pAlias & " to " & curPos
    Call mciSendString(pCommand, "", 0, 0)
    pCommand = "play " & pAlias
    Call mciSendString(pCommand, "", 0, 0)
End Sub

'一時停止
Public Sub Pause()
    If Not pHasOpen Then Exit Sub
    pCommand = "pause " & pAlias
    Call mciSendString(pCommand, "", 0, 0)
End Sub

'再生再開
Public Sub PlayResume()
    If Not pHasOpen Then Exit Sub
    pCommand = "resume " & pAlias
    Call mciSendString(pCommand, "", 0, 0)
End Sub

'停止
Public Sub StopSound()
    pCommand = "stop " & pAlias
    Call mciSendString(pCommand, "", 0, 0)
    Call CloseSound
End Sub

'クローズ
Public Sub CloseSound()
    pCommand = "close " & pAlias
    Call mciSendString(pCommand, "", 0, 0)
End Sub

'ステータス取得
'戻り値："not ready","paused","playing","stopped"
'　　　：未Open時は""空文字
Public Function GetStatus() As String
    If Not pHasOpen Then Exit Function
    Dim Mode As String
    Dim buf As String * 20
    pCommand = "status " & pAlias & " mode"
    Call mciSendString(pCommand, buf, Len(buf), 0)
    GetStatus = Replace(buf, Chr(0), "")
End Function

'再生位置取得：戻り値は秒数
Public Function GetPosition() As Double
    On Error Resume Next
    If Not pHasOpen Then Exit Function
  
    Dim buf As String * 20
    pCommand = "status " & pAlias & " position"
    Call mciSendString(pCommand, buf, Len(buf), 0)
    GetPosition = CLng(buf) / 1000
End Function

'再生時間取得：戻り値は秒数
Public Function GetLength(Optional ByVal ErrRaise As Boolean = False) As Double
    Dim ret As String
    ret = GetStatus '現在の状態を退避
  
    Call OpenSound(ErrRaise)
    If Not pHasOpen Then Exit Function
  
    Dim buf As String * 20
    pCommand = "status " & pAlias & " length"
    Call mciSendString(pCommand, buf, Len(buf), 0)
  
    '元々OpenしていなかったらClose
    If ret = "" Then Call CloseSound
  
    GetLength = CLng(buf) / 1000
End Function

'クラス終了処理
Private Sub Class_Terminate()
    'あくまで念の為
    On Error Resume Next
    Call StopSound
End Sub


