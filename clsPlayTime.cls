VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTime As Long

' 取得目前時間
Public Property Get Time() As Long
    Time = pTime
End Property

' 設定時間
Public Property Let Time(ByVal NewTime As Long)

    ' 限制時間範圍
    If NewTime < TrialDealTime Or NewTime > fTimeEnd Then Exit Property

    ' 若時間沒變就不做事
    If pTime = NewTime Then Exit Property

    ' 更新時間
    pTime = NewTime
    
    ' 同步全域變數（過渡用）
    CurrentTime = pTime

    ' 模擬事件 → 呼叫展示更新
    Call PlayTime_TimeChanged(pTime)

End Property
