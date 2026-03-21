Attribute VB_Name = "資料整理"

'由時間8:30:00~13:45:00每1秒1個序號 由序號得出對應時間
Public Function SeqToTime(ByVal seq As Long) As Long
    
    Dim baseTime As Date
    Dim targetTime As Date
    Dim endTime As Date
    Dim result As Long
    
    If seq <= 0 Then
        SeqToTime = 0
        Exit Function
    End If
    
    ' 起始時間 08:30:00
    baseTime = TimeSerial(TrialDealTime \ 10000, (TrialDealTime \ 100) Mod 100, TrialDealTime Mod 100)
    
    ' 加上 (seq - 1) 秒
    targetTime = DateAdd("s", seq - 1, baseTime)
    
    ' 超過 13:45:00 則無效
    endTime = TimeSerial(fTimeEnd \ 10000, (fTimeEnd \ 100) Mod 100, fTimeEnd Mod 100)
    If DateDiff("s", endTime, targetTime) > 0 Then
        SeqToTime = 0
        Exit Function
    End If
    
    ' 轉成 HHMMSS 格式數值
    result = Hour(targetTime) * 10000 + _
             Minute(targetTime) * 100 + _
             Second(targetTime)
    
    SeqToTime = result
    
End Function

'由時間8:30:00~13:45:00每1秒1個序號 由時間得出對應序號
Public Function TimeToSeq(ByVal t As Long) As Long
    
    Dim h As Long, m As Long, s As Long
    Dim baseTime As Date
    Dim inputTime As Date
    Dim diffSeconds As Long
    
    ' 拆解 HHMMSS
    h = t \ 10000
    m = (t \ 100) Mod 100
    s = t Mod 100
    
    ' 基本合法性檢查
    If h < 8 Or h > 13 Then Exit Function
    If m < 0 Or m > 59 Then Exit Function
    If s < 0 Or s > 59 Then Exit Function
    
    baseTime = TimeSerial(TrialDealTime \ 10000, (TrialDealTime \ 100) Mod 100, TrialDealTime Mod 100)
    inputTime = TimeSerial(h, m, s)
    
    ' 若早於 08:30:00 或晚於 13:45:00 則無效
    If inputTime < baseTime Then Exit Function
    If inputTime > TimeSerial(fTimeEnd \ 10000, (fTimeEnd \ 100) Mod 100, fTimeEnd Mod 100) Then Exit Function
    
    diffSeconds = DateDiff("s", baseTime, inputTime)
    
    ' Seq = 秒差 + 1
    TimeToSeq = diffSeconds + 1
    
End Function

'========================================================
' 通用：將 [原資料] 的「時間序列資料」整理到 [資料整備]（每秒一列）
'
' 規則：
' - 同秒多筆：成交價取該秒最後一個非空值；單量=該秒單量加總；
'             狀態(委買量~累計量)=該秒最後一筆（最後列）。
' - 缺秒：成交價、成交時間與狀態延續前一秒；單量=0
' - 新增：每秒同時記錄「對應原資料列號」
'
' 注意：
' - [資料整備] 的主時間軸在 B 欄，第 4 列為 Seq=1（TrialDealTime）
'========================================================
Public Sub BuildMarket_TimeAligned( _
    ByVal rawTimeCol As Long, _
    ByVal rawPxCol As Long, _
    ByVal rawQtyCol As Long, _
    ByVal rawStateFromCol As Long, _
    ByVal rawStateToCol As Long, _
    ByVal prepColRawRow As Long, _
    ByVal prepColTimeOut As Long, _
    ByVal prepColPx As Long, _
    ByVal prepColQty As Long, _
    ByVal prepColStateFrom As Long, _
    ByVal startRowRaw As Long _
)

    Dim wsRaw As Worksheet, wsPrep As Worksheet
    Set wsRaw = ThisWorkbook.Worksheets("原資料")
    Set wsPrep = ThisWorkbook.Worksheets("資料整備")

    Const PREP_TIME_COL As Long = 2          ' B：主時間軸
    Const PREP_FIRST_DATA_ROW As Long = 4    ' 第4列 = Seq=1

    Dim lastRaw As Long, lastPrep As Long
    lastRaw = wsRaw.Cells(wsRaw.RowS.Count, rawTimeCol).End(xlUp).row
    lastPrep = wsPrep.Cells(wsPrep.RowS.Count, PREP_TIME_COL).End(xlUp).row
    
    If lastRaw < startRowRaw Then Exit Sub
    If lastPrep < PREP_FIRST_DATA_ROW Then Exit Sub

    '--- 找原資料第一筆「有時間」的列 R0 ---
    Dim R0 As Long, T0 As Long, V As Variant
    R0 = startRowRaw
    
    Do While R0 <= lastRaw
        V = wsRaw.Cells(R0, rawTimeCol).Value2
        If Not IsEmpty(V) And IsNumeric(V) Then Exit Do
        R0 = R0 + 1
    Loop
    
    If R0 > lastRaw Then Exit Sub
    T0 = CLng(wsRaw.Cells(R0, rawTimeCol).Value2)

    '--- 找資料整備對應起始列 L0 ---
    Dim L0 As Long
    L0 = TimeToSeq(T0) + (PREP_FIRST_DATA_ROW - 1)
    If L0 < PREP_FIRST_DATA_ROW Then L0 = PREP_FIRST_DATA_ROW
    If L0 > lastPrep Then Exit Sub

    '--- 清除輸出區 ---
    Dim clearFromCol As Long, clearToCol As Long
    clearFromCol = prepColRawRow
    clearToCol = prepColStateFrom + 4
    wsPrep.Range(wsPrep.Cells(PREP_FIRST_DATA_ROW, clearFromCol), wsPrep.Cells(lastPrep, clearToCol)).ClearContents

    '--- 前值（缺秒延續用）---
    Dim prevPx As Variant
    prevPx = wsPrep.Cells(L0 - 1, prepColPx).Value                         ' 前一筆成交價

    Dim prevTradeTime As Variant
    prevTradeTime = wsPrep.Cells(L0 - 1, prepColTimeOut).Value             ' 前一筆成交時間

    Dim prevRawRow As Variant
    prevRawRow = wsPrep.Cells(L0 - 1, prepColRawRow).Value                 ' 前一筆原資料列號

    Dim prevState(1 To 5) As Variant
    Dim k As Long
    For k = 1 To 5
        prevState(k) = wsPrep.Cells(L0 - 1, prepColStateFrom + (k - 1)).Value
    Next k

    '--- 主迴圈 ---
    Dim i As Long, Ta As Long
    Dim curRawTime As Long
    Dim n As Long, lastRowInGroup As Long
    Dim sumQty As Double
    Dim lastNonEmptyPx As Variant, pxCell As Variant
    Dim rr As Long

    For i = L0 To lastPrep

        Ta = CLng(wsPrep.Cells(i, PREP_TIME_COL).Value2)                   ' 目前主時間軸時間

        ' 推進 R0 到 >= Ta
        Do While R0 <= lastRaw And IsNumeric(wsRaw.Cells(R0, rawTimeCol).Value2) _
              And CLng(wsRaw.Cells(R0, rawTimeCol).Value2) < Ta
            R0 = R0 + 1
        Loop

        ' 缺秒：原資料走完 or 下一筆不是數字
        If R0 > lastRaw Or Not IsNumeric(wsRaw.Cells(R0, rawTimeCol).Value2) Then
            wsPrep.Cells(i, prepColRawRow).Value = prevRawRow                              ' 原資料列號延續前一筆
            wsPrep.Cells(i, prepColTimeOut).Value = prevTradeTime                          ' 成交時間延續前一筆
            wsPrep.Cells(i, prepColPx).Value = prevPx                                      ' 成交價延續前一筆
            wsPrep.Cells(i, prepColQty).Value = 0                                          ' 單量=0
            For k = 1 To 5
                wsPrep.Cells(i, prepColStateFrom + (k - 1)).Value = prevState(k)          ' 狀態延續前一筆
            Next k
            GoTo NextSecond
        End If

        curRawTime = CLng(wsRaw.Cells(R0, rawTimeCol).Value2)

        ' 缺秒：下一筆時間 > Ta
        If curRawTime > Ta Then
            wsPrep.Cells(i, prepColRawRow).Value = prevRawRow                              ' 原資料列號延續前一筆
            wsPrep.Cells(i, prepColTimeOut).Value = prevTradeTime                          ' 成交時間延續前一筆
            wsPrep.Cells(i, prepColPx).Value = prevPx                                      ' 成交價延續前一筆
            wsPrep.Cells(i, prepColQty).Value = 0                                          ' 單量=0
            For k = 1 To 5
                wsPrep.Cells(i, prepColStateFrom + (k - 1)).Value = prevState(k)          ' 狀態延續前一筆
            Next k
            GoTo NextSecond
        End If

        ' 收集同秒群組 (R0 ~ R0+n-1)
        n = 0
        Do While (R0 + n) <= lastRaw _
              And IsNumeric(wsRaw.Cells(R0 + n, rawTimeCol).Value2) _
              And CLng(wsRaw.Cells(R0 + n, rawTimeCol).Value2) = Ta
            n = n + 1
        Loop
        lastRowInGroup = R0 + n - 1

        ' 寫入對應原資料列號
        wsPrep.Cells(i, prepColRawRow).Value = lastRowInGroup                              ' 寫入最後一筆原資料列號
        prevRawRow = lastRowInGroup                                                        ' 更新前一筆原資料列號

        ' 成交時間：同秒有成交資料就寫 Ta
        wsPrep.Cells(i, prepColTimeOut).Value = Ta                                         ' 寫入該秒成交時間
        prevTradeTime = Ta                                                                 ' 更新前一筆成交時間

        ' 成交價：取最後一個非空值，若都空則延續
        lastNonEmptyPx = Empty
        For rr = R0 To lastRowInGroup
            pxCell = wsRaw.Cells(rr, rawPxCol).Value2
            If Len(pxCell & vbNullString) > 0 Then lastNonEmptyPx = pxCell
        Next rr

        If Len(lastNonEmptyPx & vbNullString) > 0 Then
            wsPrep.Cells(i, prepColPx).Value = lastNonEmptyPx                              ' 寫入成交價
            prevPx = lastNonEmptyPx                                                        ' 更新前一筆成交價
        Else
            wsPrep.Cells(i, prepColPx).Value = prevPx                                      ' 若無成交價則延續前一筆
        End If

        ' 單量加總
        sumQty = 0
        For rr = R0 To lastRowInGroup
            If IsNumeric(wsRaw.Cells(rr, rawQtyCol).Value2) Then
                sumQty = sumQty + CDbl(wsRaw.Cells(rr, rawQtyCol).Value2)
            End If
        Next rr
        wsPrep.Cells(i, prepColQty).Value = sumQty                                         ' 寫入單量加總

        ' 狀態取最後一筆
        For k = 1 To 5
            wsPrep.Cells(i, prepColStateFrom + (k - 1)).Value = _
                wsRaw.Cells(lastRowInGroup, rawStateFromCol + (k - 1)).Value2             ' 寫入最後一筆狀態
            prevState(k) = wsPrep.Cells(i, prepColStateFrom + (k - 1)).Value              ' 更新前一筆狀態
        Next k

        ' 更新 R0
        R0 = R0 + n

NextSecond:
    Next i

End Sub

Public Sub Build_AllMarkets_TimeAligned()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    '-------------------------
    ' 現貨：原資料 B~I → 資料整備 C~K
    ' C(3)=原資料列號, D(4)=成交時間, E(5)=成交價, F(6)=單量, G~K(7~11)=狀態
    '-------------------------
    BuildMarket_TimeAligned _
        rawTimeCol:=2, rawPxCol:=3, rawQtyCol:=4, rawStateFromCol:=5, rawStateToCol:=9, _
        prepColRawRow:=3, prepColTimeOut:=4, _
        prepColPx:=5, prepColQty:=6, prepColStateFrom:=7, _
        startRowRaw:=6

    '-------------------------
    ' 期貨：原資料 L~S → 資料整備 M~U
    ' M(13)=原資料列號, N(14)=成交時間, O(15)=成交價, P(16)=單量, Q~U(17~21)=狀態
    '-------------------------
    BuildMarket_TimeAligned _
        rawTimeCol:=12, rawPxCol:=13, rawQtyCol:=14, rawStateFromCol:=15, rawStateToCol:=19, _
        prepColRawRow:=13, prepColTimeOut:=14, _
        prepColPx:=15, prepColQty:=16, prepColStateFrom:=17, _
        startRowRaw:=6

    '-------------------------
    ' 小型期貨：原資料 V~AC → 資料整備 W~AE
    ' W(23)=原資料列號, X(24)=成交時間, Y(25)=成交價, Z(26)=單量, AA~AE(27~31)=狀態
    '-------------------------
    BuildMarket_TimeAligned _
        rawTimeCol:=22, rawPxCol:=23, rawQtyCol:=24, rawStateFromCol:=25, rawStateToCol:=29, _
        prepColRawRow:=23, prepColTimeOut:=24, _
        prepColPx:=25, prepColQty:=26, prepColStateFrom:=27, _
        startRowRaw:=6

CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

