Attribute VB_Name = "全域變數"
Option Explicit

' 交易時間常數
Public Const fTimeStart As Long = 84500   ' 期貨開始交易時間
Public Const fTimeEnd As Long = 134500    ' 期貨停止交易時間
Public Const sTimeStart As Long = 90000   ' 現貨開始交易時間
Public Const sTimeEnd As Long = 133000    ' 現貨停止交易時間
Public Const TrialDealTime As Long = 83000 '試搓合開始時間
Public Const JiangBoStepSec As Long = 10   '江波圖每10秒取樣一點(可在此變更)


' 模擬運行變數
Public runTime As clsPlayTime  ' 目前播放時間
Public CurrentRow As Long      ' 目前時間對應之資料整備列號
Public CurrentTime As Long      ' 目前時間
Public OneStep As Integer      ' 每次時間跳動秒數


'播放控制
Public SimRunning As Boolean  ' 控制播放的開關
Public TimeRatio As Integer   ' 播放速度倍數
Public KeyModeOnOff As Boolean '允許(取消)快速鍵

'初始化的參數
Public Const Default_CsvDir As String = "D:\看盤室Modify\TradeRecord\"



