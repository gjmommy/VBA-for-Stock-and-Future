Attribute VB_Name = "交易展示"

Sub cmdForward_Click()
    ' 確保 runTime 已初始化
    If runTime Is Nothing Then
        Set runTime = New clsPlayTime
        runTime.Time = Sheets("看盤室").Range("E31").Value
    End If
    ' 更新 uTime，會自動觸發 runTime_Change()
    Forward
End Sub

Sub cmdBackward_Click()

    ' 確保 runTime 已初始化
    If runTime Is Nothing Then
        Set runTime = New clsPlayTime
        runTime.Time = ThisWorkbook.Sheets("看盤室").Range("E31").Value
    End If

    ' 後退一步
    Backward

End Sub

Sub cmdPlay_Click()

    If runTime Is Nothing Then
        Set runTime = New clsPlayTime
        runTime.Time = ThisWorkbook.Sheets("看盤室").Range("E31").Value
    End If

    SimRunning = True                                      ' 啟動播放
    TimeRatio = Sheets("看盤室").Range("G32").Value        ' 讀取播放速度

    If TimeRatio < 1 Then TimeRatio = 1
    If TimeRatio > 100 Then TimeRatio = 100

    RunSimulationLoop

End Sub
Sub cmdPause_Click()
    SimRunning = False  ' 停止播放
End Sub

' 數據更新展示
Sub UpdateDisplay()

    Dim wsPrep As Worksheet
    Dim wsView As Worksheet
    Dim lastPrepRow As Long

    Set wsPrep = ThisWorkbook.Sheets("資料整備")
    Set wsView = ThisWorkbook.Sheets("看盤室")
    
    ' 目前時間
    CurrentTime = runTime.Time                                     ' 目前播放時間
    CurrentRow = TimeToSeq(CurrentTime) + 3                        ' 由時間換算資料整備列號（第4列=Seq1）

    ' 防呆：列號超界則離開
    lastPrepRow = wsPrep.Cells(wsPrep.RowS.Count, 2).End(xlUp).row ' 資料整備B欄最後列
    If CurrentRow < 4 Or CurrentRow > lastPrepRow Then Exit Sub

    Application.EnableEvents = False
    ' ==================== 更新個股現貨資料 (Row 2) ====================
    wsView.Cells(2, 2).Value = wsPrep.Cells(CurrentRow, 5).Value          ' E欄 = 現貨成交價
    wsView.Cells(2, 15).Value = wsPrep.Cells(CurrentRow, 4).Value         ' D欄 = 現貨成交時間
    wsView.Cells(2, 9).Value = wsPrep.Cells(CurrentRow, 6).Value          ' F欄 = 現貨單量
    wsView.Cells(2, 10).Value = wsPrep.Cells(CurrentRow, 11).Value        ' K欄 = 現貨累計量
    wsView.Cells(2, 5).Resize(1, 4).Value = _
        wsPrep.Cells(CurrentRow, 7).Resize(1, 4).Value                    ' G:J = 現貨委買量/買價/賣價/委賣量

    ' ==================== 更新期貨資料 (Row 3) ====================
    wsView.Cells(3, 2).Value = wsPrep.Cells(CurrentRow, 15).Value         ' O欄 = 期貨成交價
    wsView.Cells(3, 15).Value = wsPrep.Cells(CurrentRow, 14).Value        ' N欄 = 期貨成交時間
    wsView.Cells(3, 9).Value = wsPrep.Cells(CurrentRow, 16).Value         ' P欄 = 期貨單量
    wsView.Cells(3, 10).Value = wsPrep.Cells(CurrentRow, 21).Value        ' U欄 = 期貨累計量
    wsView.Cells(3, 5).Resize(1, 4).Value = _
        wsPrep.Cells(CurrentRow, 17).Resize(1, 4).Value                   ' Q:T = 期貨委買量/買價/賣價/委賣量

    ' ==================== 更新小型期貨資料 (Row 4) ====================
    wsView.Cells(4, 2).Value = wsPrep.Cells(CurrentRow, 25).Value         ' Y欄 = 小型期貨成交價
    wsView.Cells(4, 15).Value = wsPrep.Cells(CurrentRow, 24).Value        ' X欄 = 小型期貨成交時間
    wsView.Cells(4, 9).Value = wsPrep.Cells(CurrentRow, 26).Value         ' Z欄 = 小型期貨單量
    wsView.Cells(4, 10).Value = wsPrep.Cells(CurrentRow, 31).Value        ' AE欄 = 小型期貨累計量
    wsView.Cells(4, 5).Resize(1, 4).Value = _
        wsPrep.Cells(CurrentRow, 27).Resize(1, 4).Value                   ' AA:AD = 小型期貨委買量/買價/賣價/委賣量
    wsView.Cells(6, 3).Value = runTime.Time                               '顯示指定時間
    wsView.Cells(6, 4).Value = TimeToSeq(runTime.Time) + 3                '顯示整備資料列號

    Application.EnableEvents = True
    
End Sub


Function FindLastRow(ws As Worksheet, startRow As Integer, col As Integer, CurrentTime As Long) As Long
    Dim h As Long, m As Long
    
    ' 設定搜尋範圍
    h = ws.Cells(ws.RowS.Count, col).End(xlUp).row ' 找到最後一列
    
    ' 執行二分搜尋
    Do While h - startRow > 1
        m = (startRow + h) \ 2
        If ws.Cells(m, col).Value = CurrentTime Then
            FindLastRow = m
            Exit Function
        ElseIf ws.Cells(m, col).Value < CurrentTime Then
            startRow = m ' 縮小搜尋範圍到 M-H
        Else
            h = m ' 縮小搜尋範圍到 L-M
        End If
    Loop
    
    ' 只有 1 或 2 行時，直接決定結果
    If ws.Cells(h, col).Value <= CurrentTime Then
        FindLastRow = h
    Else
        FindLastRow = startRow
    End If
End Function

' uTime 前進 (增加 OneStep)
Public Sub Forward()
    Dim h As Integer, m As Integer, s As Integer
    Dim tempTime As Double ' 避免溢位，使用 `Double` 來存儲時間
    
    ' 解析當前時間格式 (hhmmss)
    h = runTime.Time \ 10000          ' 取得小時 (hh)
    m = (runTime.Time Mod 10000) \ 100 ' 取得分鐘 (mm)
    s = runTime.Time Mod 100           ' 取得秒數 (ss)

    ' 增加秒數
    OneStep = Sheets("看盤室").Range("Onestep").Value
    If OneStep < 1 Then OneStep = 1
    s = s + OneStep

    ' 進位處理 (秒 → 分 → 時)
    If s >= 60 Then
        s = s - 60
        m = m + 1
    End If

    If m >= 60 Then
        m = 0
        h = h + 1
    End If

    ' 透過 TimeSerial 確保時間格式正確
    tempTime = TimeSerial(h, m, s) ' VBA 內建的時間函數，避免溢位

    ' 轉回 hhmmss 格式
    runTime.Time = Hour(tempTime) * 10000 + Minute(tempTime) * 100 + Second(tempTime)

    ' 確保時間不超過 `fTimeEnd`
    If runTime.Time > fTimeEnd Then runTime.Time = fTimeEnd

End Sub

' uTime 後退 (減少 OneStep)
Public Sub Backward()

    Dim h As Integer, m As Integer, s As Integer
    Dim tempTime As Double

    ' 每次重讀步長
    OneStep = ThisWorkbook.Sheets("看盤室").Range("OneStep").Value
    If OneStep < 1 Then OneStep = 1

    ' 已到起始時間就不再後退
    If runTime.Time <= TrialDealTime Then
        runTime.Time = TrialDealTime
        Exit Sub
    End If

    ' 解析目前時間格式 (hhmmss)
    h = runTime.Time \ 10000
    m = (runTime.Time Mod 10000) \ 100
    s = runTime.Time Mod 100

    ' 減少秒數
    s = s - OneStep

    ' 借位處理
    Do While s < 0
        s = s + 60
        m = m - 1
    Loop

    Do While m < 0
        m = m + 60
        h = h - 1
    Loop

    ' 用 TimeSerial 確保時間合法
    tempTime = TimeSerial(h, m, s)

    ' 寫回 runTime，並觸發畫面更新
    runTime.Time = Hour(tempTime) * 10000 + Minute(tempTime) * 100 + Second(tempTime)

    ' 下限保護
    If runTime.Time < TrialDealTime Then runTime.Time = TrialDealTime

End Sub

Sub RunSimulationLoop()

    Dim nextRun As Double
    Dim tempTime As Double
    Dim h As Integer, m As Integer, s As Integer
    
    nextRun = Timer + 1 / TimeRatio                             ' 依播放速度設定下次執行時間

    Do While SimRunning

        ' 到收盤時間即停止
        If runTime.Time >= fTimeEnd Then
            SimRunning = False
            Exit Do
        End If

        ' 解析目前時間 hhmmss
        h = runTime.Time \ 10000
        m = (runTime.Time Mod 10000) \ 100
        s = runTime.Time Mod 100

        ' 前進 1 秒
        s = s + 1

        ' 秒進位
        If s >= 60 Then
            s = s - 60
            m = m + 1
        End If

        ' 分進位
        If m >= 60 Then
            m = 0
            h = h + 1
        End If

        ' 用 TimeSerial 確保時間合法
        tempTime = TimeSerial(h, m, s)

        ' 寫回 runTime，並觸發畫面更新
        runTime.Time = Hour(tempTime) * 10000 + Minute(tempTime) * 100 + Second(tempTime)

        ' 允許 UI 更新
        DoEvents

        ' 控制播放速度
        Do While Timer < nextRun
            DoEvents
        Loop
        nextRun = Timer + 1 / TimeRatio

    Loop

End Sub



Sub CreateOrUpdateChart()

    Dim wsView As Worksheet, wsPrep As Worksheet, wsRaw As Worksheet
    Dim cht As ChartObject
    Dim rngX As Range, rngBuy As Range, rngSell As Range
    Dim chartName As String
    Dim stockName As String
    Dim buyPrice As String, buyQty As String
    Dim sellPrice As String, sellQty As String
    Dim prepRow As Long
    Dim rawAnchorRow As Long
    Dim startRawRow As Long, lastRawRow As Long

    Set wsView = ThisWorkbook.Sheets("看盤室")
    Set wsPrep = ThisWorkbook.Sheets("資料整備")
    Set wsRaw = ThisWorkbook.Sheets("原資料")
    
    chartName = "Chart_BS_Q"
    
    ' 由目前播放時間換算資料整備所在列
    CurrentRow = TimeToSeq(runTime.Time) + 3                           ' 第4列 = Seq1
    prepRow = CurrentRow
    
    ' 確保資料整備有足夠數據
    If prepRow < 4 Then Exit Sub
    
    ' 從資料整備 C 欄取出對應的原資料列號
    rawAnchorRow = Val(wsPrep.Cells(prepRow, 3).Value)                 ' C欄 = 現貨對應原資料列號
    If rawAnchorRow < 6 Then Exit Sub

    ' 設定顯示範圍：原資料往上含當前共 50 筆事件列
    startRawRow = Application.Max(6, rawAnchorRow - 49)               ' 最早不能早於第6列
    lastRawRow = rawAnchorRow                                          ' 目前對應的原資料列

    ' 設定數據範圍（現貨原資料）
    Set rngX = wsRaw.Range("B" & startRawRow & ":B" & lastRawRow)     ' B欄 = 原資料時間
    Set rngBuy = wsRaw.Range("E" & startRawRow & ":E" & lastRawRow)   ' E欄 = 委買量
    Set rngSell = wsRaw.Range("H" & startRawRow & ":H" & lastRawRow)  ' H欄 = 委賣量

    ' 股票名稱（顯示於圖表標題）
    stockName = wsView.Cells(2, 1).Value                               ' 看盤室 A2 = 現貨名稱
    
    ' 取得最新的委買 / 委賣價格與數量（原資料）
    buyPrice = wsRaw.Cells(lastRawRow, 6).Value                        ' F欄 = 買價
    buyQty = wsRaw.Cells(lastRawRow, 5).Value                          ' E欄 = 委買量
    sellPrice = wsRaw.Cells(lastRawRow, 7).Value                       ' G欄 = 賣價
    sellQty = wsRaw.Cells(lastRawRow, 8).Value                         ' H欄 = 委賣量

    ' 檢查圖表是否已存在
    For Each cht In wsView.ChartObjects
        If cht.Name = chartName Then
            
            ' 圖表已存在：直接更新數據範圍與標示
            With cht.Chart
                .SeriesCollection(1).XValues = rngX
                .SeriesCollection(1).Values = rngBuy
                .SeriesCollection(2).XValues = rngX
                .SeriesCollection(2).Values = rngSell
                
                .ChartTitle.Text = stockName
                .SeriesCollection(1).Name = "買 " & buyPrice & " / " & buyQty
                .SeriesCollection(2).Name = "賣 " & sellPrice & " / " & sellQty
            End With
            
            Exit Sub
        End If
    Next cht

    ' 若圖表不存在，則新建圖表
    Set cht = wsView.ChartObjects.Add(Left:=15, Top:=200, Width:=350, Height:=240)
    cht.Name = chartName
    cht.Chart.SetSourceData Source:=rngBuy

    ' 設定圖表類型為折線圖
    With cht.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = stockName
        
        ' X 軸不顯示
        .Axes(xlCategory).HasTitle = False
        .Axes(xlCategory).TickLabels.NumberFormat = " "
        .Axes(xlCategory).Format.Line.Visible = msoFalse
        
        ' Y 軸不顯示標題
        .Axes(xlValue).HasTitle = False
        
        ' Y 軸文字顏色
        .Axes(xlValue).TickLabels.Font.Color = RGB(255, 255, 255)
        
        ' 背景與繪圖區顏色
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(64, 64, 64)
        
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 16
        .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' 移除自動產生的數列
        On Error Resume Next
        .SeriesCollection(1).Delete
        On Error GoTo 0

        ' 設定買方數列（紅色）
        With .SeriesCollection.NewSeries
            .XValues = rngX
            .Values = rngBuy
            .Name = "買 " & buyPrice & " / " & buyQty
            .Border.Color = RGB(255, 0, 0)
            .MarkerStyle = xlNone
        End With

        ' 設定賣方數列（綠色）
        With .SeriesCollection.NewSeries
            .XValues = rngX
            .Values = rngSell
            .Name = "賣 " & sellPrice & " / " & sellQty
            .Border.Color = RGB(0, 128, 0)
            .MarkerStyle = xlNone
        End With
    End With

    ' 調整圖表大小
    With wsView.Shapes(chartName)
        .ScaleWidth 1.3, msoFalse, msoScaleFromTopLeft
        .ScaleHeight 1.3, msoFalse, msoScaleFromTopLeft
    End With
    
    ' 調整標題字體與位置
    With cht.Chart.ChartTitle
        .Format.TextFrame2.TextRange.Font.Size = 11
        .Left = 3
        .Top = 5
    End With
    
    ' 調整圖例位置
    With cht.Chart.Legend
        .Left = 200
        .Top = 1
    End With
    
    ' 調整繪圖區大小
    With cht.Chart.PlotArea
        .Width = .Width * 1.3
        .Height = .Height * 1.5
    End With
    
    ' 再次調整標題字型大小
    cht.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14
    
    ' 圖例文字顏色與大小
    If cht.Chart.HasLegend Then
        Dim legendEntry As Object
        For Each legendEntry In cht.Chart.Legend.LegendEntries
            legendEntry.Font.Color = RGB(255, 255, 200)
            legendEntry.Font.Size = 12
        Next legendEntry
    End If

End Sub

Sub CreateJiangBoChart(NowTime As Long)

    Dim ws As Worksheet
    Dim wsView As Worksheet
    Dim cht As ChartObject
    Dim rngTime As Range, rngPrice As Range, rngVolume As Range
    Dim lastRow As Long, ToRow As Long
    Dim maxVol As Double
    Dim chartName As String
    Dim hh As Long, mm As Long, ss As Long
    
    ' 名稱設定
    Set ws = ThisWorkbook.Sheets("原資料")
    Set wsView = ThisWorkbook.Sheets("看盤室")
    chartName = "JiangBoChart"

    ' 找到最後一筆江波圖資料
    lastRow = ws.Cells(ws.RowS.Count, 41).End(xlUp).row   ' AO欄 = 時間
    lastRow = lastRow + 1                                 ' 多留一列供成交量縮放參考

    ' 由 NowTime 直接計算 ToRow（09:00:00 起每 JiangBoStepSec 秒 1 筆，AO3 為第1筆）
    hh = NowTime \ 10000
    mm = (NowTime \ 100) Mod 100
    ss = NowTime Mod 100

    ToRow = ((hh - 9) * 3600 + mm * 60 + ss) \ JiangBoStepSec + 3

    ' 防呆：限制 ToRow 範圍
    If ToRow < 3 Then ToRow = 3
    If ToRow > lastRow Then ToRow = lastRow

    ' 一次性批量轉貼值
    ws.Range("AR3:AS" & lastRow).ClearContents
    ws.Range("AR3:AR" & ToRow).Value = ws.Range("AP3:AP" & ToRow).Value   ' AR = 成交價顯示
    ws.Range("AS3:AS" & ToRow).Value = ws.Range("AQ3:AQ" & ToRow).Value   ' AS = 區間成交量顯示
    
    Application.CutCopyMode = False

    ' 設定成交量最大值
    maxVol = Application.WorksheetFunction.Max(ws.Range("AQ3:AQ" & lastRow))
    maxVol = maxVol * 1.5

    ' 如果圖表已存在，直接更新數列
    For Each cht In wsView.ChartObjects
        If cht.Name = chartName Then
            With cht.Chart
                .SeriesCollection(1).Values = ws.Range("AR3:AR" & lastRow)
                .SeriesCollection(2).Values = ws.Range("AS3:AS" & lastRow)
                .Refresh
            End With
            DoEvents
            Exit Sub
        End If
    Next cht

    ' 設定數據範圍
    Set rngTime = ws.Range("AO3:AO" & lastRow)
    Set rngPrice = ws.Range("AR3:AR" & lastRow)
    Set rngVolume = ws.Range("AS3:AS" & lastRow)

    ' 建立新圖表
    Set cht = wsView.ChartObjects.Add(Left:=500, Top:=200, Width:=420, Height:=310)
    cht.Name = chartName
    cht.Chart.ChartType = xlLine

    ' 設定黑色背景
    With cht.Chart.ChartArea
        .Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With

    ' 設定成交價格線
    With cht.Chart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = rngTime
        .SeriesCollection(1).Values = rngPrice
        .SeriesCollection(1).Name = "成交價"
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 255, 255)
        .SeriesCollection(1).Format.Line.Weight = 0.75
        .SeriesCollection(1).Smooth = False
        .SeriesCollection(1).MarkerStyle = xlNone
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(64, 64, 64)
    End With

    ' 設定成交量柱狀圖
    With cht.Chart
        .SeriesCollection.NewSeries
        .SeriesCollection(2).XValues = rngTime
        .SeriesCollection(2).Values = rngVolume
        .SeriesCollection(2).Name = "成交量"
        .SeriesCollection(2).ChartType = xlColumnClustered
        .SeriesCollection(2).AxisGroup = xlSecondary
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 204, 0)
    End With

    ' 設定成交量 Y 軸
    With cht.Chart.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = maxVol
        .TickLabels.Font.Color = RGB(255, 255, 255)
        .TickLabels.Font.Size = 12
    End With

    ' 設定 X 軸格式
    With cht.Chart.Axes(xlCategory)
        .HasTitle = False
        .TickLabels.NumberFormat = "[hh]:mm"
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .Format.Line.Visible = msoFalse
    End With

    ' 設定成交價 Y 軸
    With cht.Chart.Axes(xlValue)
        .TickLabels.Font.Color = RGB(255, 255, 255)
        .TickLabels.Font.Size = 12
    End With

    ' 移除圖例
    cht.Chart.Legend.Delete

    ' 刪除 X 軸標籤
    With cht.Chart.Axes(xlCategory)
        .HasTitle = False
        .TickLabels.Delete
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
    End With

    ' 強制更新畫面
    Application.ScreenUpdating = True
    DoEvents

End Sub
' 設定江波圖之成交量座標範圍
Sub ChangeSecondaryAxisMax()
    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim maxScale As Double
    Dim userInput As String
    
    ' 設定工作表
    Set ws = ThisWorkbook.Sheets("看盤室")
    
    ' 設定圖表
    On Error Resume Next
    Set cht = ws.ChartObjects("JiangBoChart")
    On Error GoTo 0
    
    ' 確保圖表存在
    If cht Is Nothing Then
        MsgBox "找不到圖表 'JiangBoChart'，請確認名稱是否正確。", vbExclamation
        Exit Sub
    End If
    
    ' 取得使用者輸入的最大值
    userInput = InputBox("輸入欲設定" & vbNewLine & "成交量的最大值：", "設定最大值")
    
    ' 確保輸入為數字
    If IsNumeric(userInput) Then
        maxScale = CDbl(userInput)
        
        ' 設定次要座標軸最大值
        With cht.Chart.Axes(xlValue, xlSecondary)
            .MaximumScale = maxScale
        End With
        
        MsgBox "次要座標軸最大值已設定為 " & maxScale, vbInformation
    Else
        MsgBox "請輸入有效的數字。", vbExclamation
    End If
End Sub


'將 Sheets("原資料") 的成交資料填入 Sheets("即時成交") 中
Sub TransferTradeData()

    Dim wsData As Worksheet, wsTrade As Worksheet
    Dim lastRow As Long, lastTradeRow As Long
    Dim i As Long
    Dim YesterdayPrice As Double
    
    ' 設定工作表
    Set wsData = ThisWorkbook.Sheets("原資料")
    Set wsTrade = ThisWorkbook.Sheets("即時成交")
    
    ' 找到原資料 AE 欄最後一列
    lastRow = wsData.Cells(wsData.RowS.Count, "AE").End(xlUp).row
    If lastRow < 3 Then Exit Sub                                  ' 沒有成交資料就離開
    
    ' 即時成交頁資料庫區最後一列（AE3 -> K2，所以最後列 = lastRow - 1）
    lastTradeRow = lastRow - 1
    If lastTradeRow < 2 Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' (0) 清除舊資料內容（保留格式）
    wsTrade.Range("K2:P" & wsTrade.RowS.Count).ClearContents       ' 資料庫區 K:P
    wsTrade.Range("A2:F" & wsTrade.RowS.Count).ClearContents       ' 顯示區 A:F
    
    ' 設定背景色
    wsTrade.Range("K2:P" & lastTradeRow + 50).Interior.Color = RGB(64, 64, 64) ' 深灰色
    wsTrade.Range("A2:F" & lastTradeRow + 50).Interior.Color = RGB(64, 64, 64) ' 深灰色
    
    ' (1) 將原資料 AE:AJ 搬到 即時成交 K:P
    wsTrade.Range("K2").Resize(lastRow - 2, 1).Value = wsData.Range("AE3:AE" & lastRow).Value ' 時間
    wsTrade.Range("L2").Resize(lastRow - 2, 1).Value = wsData.Range("AH3:AH" & lastRow).Value ' 成交價
    wsTrade.Range("M2").Resize(lastRow - 2, 1).Value = wsData.Range("AJ3:AJ" & lastRow).Value ' 單量
    wsTrade.Range("N2").Resize(lastRow - 2, 1).Value = wsData.Range("AI3:AI" & lastRow).Value ' 漲跌
    wsTrade.Range("O2").Resize(lastRow - 2, 1).Value = wsData.Range("AF3:AF" & lastRow).Value ' 買價
    wsTrade.Range("P2").Resize(lastRow - 2, 1).Value = wsData.Range("AG3:AG" & lastRow).Value ' 賣價
    
    ' (2) 計算昨收 = 最後一筆成交價 - 最後一筆漲跌
    YesterdayPrice = wsTrade.Cells(lastTradeRow, 12).Value - wsTrade.Cells(lastTradeRow, 14).Value
    
    ' (3) 設定 L / O / P 欄字體顏色（成交價 / 買價 / 賣價）
    For i = 2 To lastTradeRow
        
        ' L欄 = 成交價
        If wsTrade.Cells(i, 12).Value > YesterdayPrice Then
            wsTrade.Cells(i, 12).Font.Color = RGB(255, 0, 0)              ' 紅色
        ElseIf wsTrade.Cells(i, 12).Value < YesterdayPrice Then
            wsTrade.Cells(i, 12).Font.Color = RGB(0, 176, 80)             ' 綠色
        Else
            wsTrade.Cells(i, 12).Font.Color = RGB(255, 192, 0)            ' 黃色
        End If
        
        ' O欄 = 買價
        If wsTrade.Cells(i, 15).Value > YesterdayPrice Then
            wsTrade.Cells(i, 15).Font.Color = RGB(255, 0, 0)
        ElseIf wsTrade.Cells(i, 15).Value < YesterdayPrice Then
            wsTrade.Cells(i, 15).Font.Color = RGB(0, 176, 80)
        Else
            wsTrade.Cells(i, 15).Font.Color = RGB(255, 192, 0)
        End If
        
        ' P欄 = 賣價
        If wsTrade.Cells(i, 16).Value > YesterdayPrice Then
            wsTrade.Cells(i, 16).Font.Color = RGB(255, 0, 0)
        ElseIf wsTrade.Cells(i, 16).Value < YesterdayPrice Then
            wsTrade.Cells(i, 16).Font.Color = RGB(0, 176, 80)
        Else
            wsTrade.Cells(i, 16).Font.Color = RGB(255, 192, 0)
        End If
        
    Next i
    
    ' (4) 設定 M 欄（單量）字體顏色
    For i = 2 To lastTradeRow
        If wsTrade.Cells(i, 12).Value >= wsTrade.Cells(i, 16).Value Then
            wsTrade.Cells(i, 13).Font.Color = RGB(255, 0, 0)              ' 外盤偏紅
        ElseIf wsTrade.Cells(i, 12).Value <= wsTrade.Cells(i, 15).Value Then
            wsTrade.Cells(i, 13).Font.Color = RGB(100, 250, 255)          ' 內盤偏藍
        Else
            wsTrade.Cells(i, 13).Font.Color = RGB(255, 192, 0)            ' 中間黃色
        End If
    Next i
    
    ' (5) K 欄（時間）字體改黃色
    wsTrade.Range("K2:K" & lastTradeRow).Font.Color = RGB(255, 192, 0)
    
    ' (6) 同步 A:F 的完整格式與 K:P 相同（只做一次）
    wsTrade.Range("K2:P" & lastTradeRow).Copy
    wsTrade.Range("A2").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' (7) 依目前播放時間更新顯示區 A:F
    
    If runTime Is Nothing Then
        Set runTime = New clsPlayTime
    End If
SetVisibleTradeData runTime.Time

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Set wsData = Nothing
    Set wsTrade = Nothing

End Sub


Sub SetVisibleTradeData(TradeTime As Long)

    Dim wsTrade As Worksheet
    Dim lastRow As Long
    Dim visibleRow As Long
    Dim rng As Range
    Dim srcRange As Range
    Dim dstRange As Range
    Dim visibleCount As Long

    ' 設定工作表
    Set wsTrade = ThisWorkbook.Sheets("即時成交")
    
    ' 找到 K 欄（成交時間）的最後一筆資料
    lastRow = wsTrade.Cells(wsTrade.RowS.Count, "K").End(xlUp).row
    If lastRow < 2 Then Exit Sub                                ' 沒有資料就離開
    
    ' 在 K2:KlastRow 中找 <= TradeTime 的最後一筆
    On Error Resume Next
    Set rng = wsTrade.Range("K2:K" & lastRow)
    visibleRow = Application.WorksheetFunction.Match(TradeTime, rng, 1) + 1
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 先清空顯示區
    wsTrade.Range("A2:F" & lastRow).ClearContents
    
    ' 若找不到符合條件的列，直接離開（畫面保持空白）
    If visibleRow < 2 Then GoTo SafeExit
    
    ' 計算要顯示的資料筆數
    visibleCount = visibleRow - 1                               ' 因為資料從第2列開始
    
    ' 來源範圍 K:P → 目標範圍 A:F
    Set srcRange = wsTrade.Range("K2:P" & visibleRow)
    Set dstRange = wsTrade.Range("A2").Resize(visibleCount, 6)
    
    ' 直接搬值（比 Copy/Paste 快很多）
    dstRange.Value = srcRange.Value
    
    
    ' 先把焦點移到即時成交視窗，避免捲錯視窗
    wsTrade.Activate
    wsTrade.Range("A" & visibleRow).Select
    
    ' 捲動到最新可見行
    If visibleRow > 38 Then
        Windows(2).ScrollRow = visibleRow - 35
    Else
        Windows(2).ScrollRow = 2
    End If

    ' 再把焦點還給看盤室控制區
    ThisWorkbook.Sheets("看盤室").Activate
    ThisWorkbook.Sheets("看盤室").Range("E30").Select
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
SafeExit:
    Set rng = Nothing
    Set srcRange = Nothing
    Set dstRange = Nothing
    Set wsTrade = Nothing

End Sub
Public Sub PlayTime_TimeChanged(ByVal NewTime As Long)

    ' 更新儲存格資料
    Call UpdateDisplay
    
    ' 以下逐步打開
    Call CreateOrUpdateChart
    Call CreateJiangBoChart(NewTime)
    Call SetVisibleTradeData(NewTime)

End Sub

Public Sub Test_RunTime()

    If runTime Is Nothing Then
        Set runTime = New clsPlayTime
    End If

    runTime.Time = 111030
    
End Sub
