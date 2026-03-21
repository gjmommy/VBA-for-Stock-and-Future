VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub cmdNewTimeSeq_Click()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("資料整備")

    Dim baseTime As Date, endTime As Date
    baseTime = TimeSerial(TrialDealTime \ 10000, (TrialDealTime \ 100) Mod 100, TrialDealTime Mod 100)
    endTime = TimeSerial(fTimeEnd \ 10000, (fTimeEnd \ 100) Mod 100, fTimeEnd Mod 100)

    ' 以「整秒」計算 N（包含起訖秒）
    Dim n As Long
    n = DateDiff("s", baseTime, endTime) + 1
    If n <= 0 Then
        MsgBox "時間範圍不合法（TrialDealTime > fTimeEnd）", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' 清掉舊資料（A:B 從第3列往下）
    ws.Range("A3:B" & ws.RowS.Count).ClearContents

    ' 用陣列一次寫入（快很多）
    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 2)

    Dim i As Long
    Dim t As Date
    Dim hh As Long, mm As Long, ss As Long

    For i = 1 To n
        t = DateAdd("s", i - 1, baseTime)

        arr(i, 1) = i ' Seq

        hh = Hour(t): mm = Minute(t): ss = Second(t)
        arr(i, 2) = hh * 10000 + mm * 100 + ss ' HHMMSS (Long)
    Next i

    ws.Range("A4").Resize(n, 2).Value = arr

CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
