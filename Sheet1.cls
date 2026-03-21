VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim isPlaying As Boolean ' 用來標記播放狀態

Private Sub cmdKeyMode_Click()
    ToKeyMode
    MsgBox "鍵盤快捷鍵說明：" & vbCrLf & vbNewLine & _
           "Shift + →  前進" & vbNewLine & _
           "Shift + ←  後退" _
        , vbInformation, "鍵盤操作說明"
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 若雙擊範圍在 E2:H4 內，則執行動作
    If Not Intersect(Target, Me.Range("E2:H4")) Is Nothing Then
        Cancel = True ' 取消 Excel 預設的編輯模式
        Call CreateOrUpdateChart
    End If
    
    ' 若雙擊範圍在 A2:A4，則重新載入商品名稱
    If Not Intersect(Target, Me.Range("A2:A4")) Is Nothing Then
        Cancel = True ' 取消 Excel 預設的編輯模式
        Dim wsData As Worksheet
        Set wsData = ThisWorkbook.Sheets("原資料")
        
        Dim i As Integer
        Dim prodName As String
        i = Target.row - 2 ' A2 → 0, A3 → 1, A4 → 2
        
        ' 讀取 A1/K1/U1 (對應於 i=0,1,2 時的 A1, K1, U1)
        prodName = wsData.Cells(1, i * 10 + 1).Value
        
        ' 若 A1/K1/U1 空白，則改取 A2/K2/U2
        If Trim(prodName) = "" Then
            prodName = wsData.Cells(2, i * 10 + 1).Value
        End If
        
        ' 填入對應的 A2:A4
        If Trim(prodName) <> "" Then Target.Value = prodName
    End If
    
    If Target.Address = "$E$32" Then 'ChangeSecondaryAxisMax
        Cancel = True
        ReFreshScreen
    End If
    
    If Target.Address = "$O$6" Then
        ChangeSecondaryAxisMax
        Cancel = True
    End If
    
    
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("原資料")

    On Error Resume Next ' 防止錯誤發生時 Excel 停止
    
    Select Case Target.Address
        ' ?? 當 E31 (時間輸入) 變動時，更新 runTime
        Case "$E$31"
            If IsNumeric(Target.Value) And Target.Value >= fTimeStart And Target.Value <= fTimeEnd Then
                If runTime Is Nothing Then
                    Set runTime = New clsPlayTime
                End If
                runTime.Time = CLng(Target.Value) ' 轉換為 Long 類型
            Else
                MsgBox "輸入時間無效！請輸入 " & fTimeStart & " ~ " & fTimeEnd & " 之間的數值。", vbExclamation
                Application.EnableEvents = False
                Target.Value = runTime.Time ' 恢復原本的時間
                Application.EnableEvents = True
            End If

        ' ?? 當 OneStep 變動時，更新變數
        Case Range("OneStep").Address
            If IsNumeric(Target.Value) And Target.Value > 0 Then
                OneStep = CLng(Target.Value)
            Else
                MsgBox "OneStep 必須為正數！", vbExclamation
                Application.EnableEvents = False
                Target.Value = OneStep ' 恢復原本的值
                Application.EnableEvents = True
            End If

        ' ?? 當 NoRowToChange 變動時，檢查範圍後更新變數
        Case Range("NoRowToChange").Address
            If IsNumeric(Target.Value) Then
                Dim maxRow As Long
                maxRow = ws.Cells(ws.RowS.Count, 2).End(xlUp).row
                
                If Target.Value < 1 Then
                    NoRowToChange = 1
                    Target.Value = 1
                ElseIf Target.Value > maxRow Then
                    NoRowToChange = maxRow
                Else
                    NoRowToChange = CLng(Target.Value)
                End If
            Else
                MsgBox "請輸入有效的行數 (5 ~ " & maxRow & ")！", vbExclamation
                Application.EnableEvents = False
                Target.Value = NoRowToChange ' 恢復原本的值
                Application.EnableEvents = True
            End If


        Case Else
            ' 其他變動不處理
    End Select

    On Error GoTo 0 ' 恢復預設錯誤處理
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    If Not Intersect(Target, Me.Range("A29:A31")) Is Nothing Then
        Application.EnableEvents = False
        Me.Range("A33").Select   ' 安全的可選格
        Application.EnableEvents = True
    End If

End Sub


Private Sub cmdForward_Click()
    Call 交易展示.cmdForward_Click
End Sub

Private Sub cmdBackward_Click()
    Call 交易展示.cmdBackward_Click
End Sub


Private Sub cmbYear_GotFocus() ' 清空值，確保使用者每次選擇時都會觸發 _Change
    cmbYear.Value = ""
End Sub
Private Sub cmbMonth_GotFocus() ' 清空值，確保使用者每次選擇時都會觸發 _Change
    cmbMonth.Value = ""
End Sub
Private Sub cmbDay_GotFocus() ' 清空值，確保使用者每次選擇時都會觸發 _Change
    cmbDay.Value = ""
End Sub
Private Sub cmbStock_GotFocus()
    cmbStock.Value = "" ' 清空值，確保使用者每次選擇時都會觸發 _Change
End Sub

'選擇資料檔所在子目錄
Private Sub cmdDirCsv_Click()
    Dim fd As fileDialog
    Dim selectedFolderPath As String

    Set fd = Application.fileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "選擇 CSV 資料所在的資料夾"
        .InitialFileName = Sheets("看盤室").Range("A31") ' 預設路徑
        .AllowMultiSelect = False

        If .Show = -1 Then
            selectedFolderPath = .SelectedItems(1)
            
            Dim probe As String
            probe = Dir(selectedFolderPath & "\*_????_??.csv")   ' 例：台積電_0718_25.csv
            
            If probe = "" Then
                MsgBox "子目錄中沒有資料檔", vbExclamation
                Exit Sub
            End If
            
            selectedFolderPath = selectedFolderPath & "\"
            Application.EnableEvents = False
            Sheets("看盤室").Cells(29, 1).Value = selectedFolderPath
            Application.EnableEvents = True
            
        End If
    End With
End Sub

'選擇個股與日期
Private Sub cmdChoice_Click()
    Dim folderPath As String
    Dim fileName As String
    Dim fileParts() As String
    Dim stockName As String
    Dim uniqueStocks As Collection
    Dim stockItem As Variant

    ' 設定目錄路徑
    folderPath = Sheets("看盤室").Range("A29") ' 修改為您的實際路徑

    ' 初始化
    Set uniqueStocks = New Collection
    
    cmbStock.Enabled = True
    ' 禁用其他控制項
    cmbYear.Enabled = False
    cmbMonth.Enabled = False
    cmbDay.Enabled = False
    cmdDownLoad.Enabled = False

    ' 清空 cmbStock
    cmbStock.Clear

    ' 搜尋子目錄中的 CSV 檔案
    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        fileParts = Split(fileName, "_")
        If UBound(fileParts) >= 2 Then
            ' 提取個股名稱
            stockName = fileParts(0)
            If InStr(stockName, "期貨") > 0 Then
                If Left(stockName, 2) = "小型" Then
                    stockName = Mid(stockName, 3)
                End If
                stockName = Left(stockName, InStr(stockName, "期貨") - 1)
            ElseIf InStr(stockName, "台指期近一") > 0 Then
                stockName = "台指期近一"
            ElseIf InStr(stockName, "小型電子近一") > 0 Then
                stockName = "小型電子近一"
            End If

            ' 添加到 uniqueStocks 集合
            On Error Resume Next
            uniqueStocks.Add stockName, stockName ' 使用 stockName 作為鍵，避免重複
            On Error GoTo 0
        End If
        fileName = Dir ' 取得下一個檔案
    Loop


    ' 將 uniqueStocks 添加到 cmbStock
    For Each stockItem In uniqueStocks
        cmbStock.AddItem stockItem
    Next stockItem

    MsgBox "OK~ 請點選欲載入的個股", vbInformation
End Sub

'選擇個股
Private Sub cmbStock_Change()
    Dim folderPath As String
    Dim fileName As String
    Dim fileParts() As String
    Dim stockName As String
    Dim selectedStock As String
    Dim year As String
    Dim uniqueYears As Collection
    Dim yearItem As Variant

    ' 設定目錄路徑
    folderPath = Sheets("看盤室").Range("A29") ' 修改為您的實際路徑

    ' 初始化
    Set uniqueYears = New Collection
    selectedStock = cmbStock.Value

    ' 如果未選擇個股則退出
    If selectedStock = "" Then Exit Sub

    ' 禁用月份和日期下拉選單
    cmbMonth.Enabled = False
    cmbDay.Enabled = False

    ' 清空年份下拉選單
    cmbYear.Clear

    ' 搜尋子目錄中的 CSV 檔案
    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        fileParts = Split(fileName, "_")
        If UBound(fileParts) >= 2 Then
            ' 提取檔案中的個股名稱
            stockName = fileParts(0)
            If InStr(stockName, "期貨") > 0 Then
                stockName = "台積電" ' 合併期貨至「台積電」
            ElseIf InStr(stockName, "台指期近一") > 0 Then
                stockName = "台指期近一"
            ElseIf InStr(stockName, "小型電子近一") > 0 Then
                stockName = "小型電子近一"
            End If

            ' 如果檔案個股名稱與選中的個股相符
            If stockName = selectedStock Then
                year = Left(fileParts(2), 2) ' 提取年份（例如 "25" 表示 2025 年）

                ' 添加年份到集合（避免重複）
                On Error Resume Next
                uniqueYears.Add year, year ' 使用年份作為鍵
                On Error GoTo 0
            End If
        End If
        fileName = Dir ' 取得下一個檔案
    Loop

    ' 將 uniqueYears 添加到 cmbYear
    For Each yearItem In uniqueYears
        cmbYear.AddItem yearItem
    Next yearItem

    ' 啟用年份下拉選單
    cmbYear.Enabled = True

    MsgBox "OK~ 請點選欲載入的年份", vbInformation
End Sub

'選擇年份
Private Sub cmbYear_Change()
    Dim folderPath As String
    Dim fileName As String
    Dim fileParts() As String
    Dim stockName As String
    Dim selectedStock As String
    Dim selectedYear As String
    Dim month As String
    Dim uniqueMonths As Collection
    Dim monthItem As Variant

    ' 設定目錄路徑
    folderPath = Sheets("看盤室").Range("A29") ' 修改為您的實際路徑

    ' 初始化
    Set uniqueMonths = New Collection
    selectedStock = cmbStock.Value
    selectedYear = cmbYear.Value

    ' 如果未選擇個股或年份則退出
    If selectedStock = "" Or selectedYear = "" Then Exit Sub

    cmbMonth.Enabled = True
    ' 禁用日期下拉選單
    cmbDay.Enabled = False

    ' 清空月份下拉選單
    cmbMonth.Clear

    ' 搜尋子目錄中的 CSV 檔案
    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        fileParts = Split(fileName, "_")
        If UBound(fileParts) >= 2 Then
            ' 提取檔案中的個股名稱
            stockName = fileParts(0)
            If InStr(stockName, "期貨") > 0 Then
                stockName = "台積電" ' 合併期貨至「台積電」
            ElseIf InStr(stockName, "台指期近一") > 0 Then
                stockName = "台指期近一"
            ElseIf InStr(stockName, "小型電子近一") > 0 Then
                stockName = "小型電子近一"
            End If

            ' 如果檔案個股名稱與選中的個股相符，且年份匹配
            If stockName = selectedStock And Left(fileParts(2), 2) = selectedYear Then
                month = Left(fileParts(1), 2) ' 提取月份（例如 "01" 表示 1 月）

                ' 添加月份到集合（避免重複）
                On Error Resume Next
                uniqueMonths.Add month, month ' 使用月份作為鍵
                On Error GoTo 0
            End If
        End If
        fileName = Dir ' 取得下一個檔案
    Loop

    ' 將 uniqueMonths 添加到 cmbMonth
    For Each monthItem In uniqueMonths
        cmbMonth.AddItem monthItem
    Next monthItem

    ' 啟用月份下拉選單
    cmbMonth.Enabled = True

    MsgBox "OK~ 請點選欲載入的月份", vbInformation
End Sub

'選擇月份
Private Sub cmbMonth_Change()
    Dim folderPath As String
    Dim fileName As String
    Dim fileParts() As String
    Dim stockName As String
    Dim selectedStock As String
    Dim selectedYear As String
    Dim selectedMonth As String
    Dim day As String
    Dim uniqueDays As Collection
    Dim dayItem As Variant

    ' 設定目錄路徑
    folderPath = Sheets("看盤室").Range("A29") ' 修改為您的實際路徑

    ' 初始化
    Set uniqueDays = New Collection
    selectedStock = cmbStock.Value
    selectedYear = cmbYear.Value
    selectedMonth = cmbMonth.Value

    ' 如果未選擇個股、年份或月份則退出
    If selectedStock = "" Or selectedYear = "" Or selectedMonth = "" Then Exit Sub

    cmbDay.Enabled = True
    ' 清空日期下拉選單
    cmbDay.Clear

    ' 搜尋子目錄中的 CSV 檔案
    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        fileParts = Split(fileName, "_")
        If UBound(fileParts) >= 2 Then
            ' 提取檔案中的個股名稱
            stockName = fileParts(0)
            If InStr(stockName, "期貨") > 0 Then
                stockName = "台積電" ' 合併期貨至「台積電」
            ElseIf InStr(stockName, "台指期近一") > 0 Then
                stockName = "台指期近一"
            ElseIf InStr(stockName, "小型電子近一") > 0 Then
                stockName = "小型電子近一"
            End If

            ' 如果檔案個股名稱、年份和月份均匹配
            If stockName = selectedStock And Left(fileParts(2), 2) = selectedYear And Left(fileParts(1), 2) = selectedMonth Then
                day = Right(fileParts(1), 2) ' 提取日期（例如 "13" 表示 13 日）

                ' 添加日期到集合（避免重複）
                On Error Resume Next
                uniqueDays.Add day, day ' 使用日期作為鍵
                On Error GoTo 0
            End If
        End If
        fileName = Dir ' 取得下一個檔案
    Loop

    ' 將 uniqueDays 添加到 cmbDay
    For Each dayItem In uniqueDays
        cmbDay.AddItem dayItem
    Next dayItem

    ' 啟用日期下拉選單
    cmbDay.Enabled = True

    MsgBox "OK~ 請點選欲載入的日期", vbInformation
End Sub

'選擇 日
Private Sub cmbDay_Change()
    Dim selectedStock As String
    Dim selectedYear As String
    Dim selectedMonth As String
    Dim selectedDay As String
    Dim formattedDate As String

    ' 獲取選擇的個股、年份、月份和日期
    selectedStock = cmbStock.Value
    selectedYear = cmbYear.Value
    selectedMonth = cmbMonth.Value
    selectedDay = cmbDay.Value

    ' 如果任一項為空則退出
    If selectedStock = "" Or selectedYear = "" Or selectedMonth = "" Or selectedDay = "" Then Exit Sub

    ' 格式化日期為 MMDD_YY
    formattedDate = selectedMonth & selectedDay & "_" & selectedYear

    ' 將資料填入指定儲存格
    Range("A27").Value = selectedStock    ' 個股名稱
    Range("B27").Value = formattedDate   ' 格式化日期 MMDD_YY
    
    cmdDownLoad.Enabled = True

    MsgBox "按下[載入資料]鍵 取得資料！", vbInformation
End Sub
' 將指定資料載入 [原資料] 工作表，並重建展示所需資料
Private Sub cmdDownLoad_Click()

    Dim folderPath As String
    Dim selectedStock As String
    Dim selectedDate As String
    Dim fileName As String
    Dim fullFilePath As String
    Dim wsRaw As Worksheet
    Dim wbCSV As Workbook
    Dim colOffset As Integer
    Dim csvFileVariants As Collection
    Dim variantItem As Variant
    Dim YYMM As String
    Dim tradeFileName As String
    Dim tradeFullPath As String
    Dim cell As Range

    '========================
    ' (0) 讀取使用者設定
    '========================
    folderPath = ThisWorkbook.Sheets("看盤室").Cells(29, 1).Value
    If folderPath = "" Then
        MsgBox "子目錄路徑未指定！", vbExclamation
        Exit Sub
    End If

    selectedStock = cmbStock.Value
    selectedDate = cmbMonth.Value & cmbDay.Value & "_" & cmbYear.Value
    
    If selectedStock = "" Or selectedDate = "" Then
        MsgBox "請先選擇個股和交易日期！", vbExclamation
        Exit Sub
    End If

    ' 當月代碼（例如 2603）
    YYMM = cmbYear.Value & cmbMonth.Value

    '========================
    ' (1) 清空 [原資料]
    '========================
    Set wsRaw = ThisWorkbook.Sheets("原資料")
    wsRaw.Cells.Clear

    '========================
    ' (2) 定義要載入的 3 份議價資料
    '     現貨 / 當月期貨 / 當月小型期貨
    '========================
    Set csvFileVariants = New Collection
    
    csvFileVariants.Add selectedStock & "_" & selectedDate & ".csv"                                 ' 現貨
    csvFileVariants.Add selectedStock & "期貨近一(" & YYMM & ")_" & selectedDate & ".csv"          ' 當月期貨
    csvFileVariants.Add "小型" & selectedStock & "期貨近一(" & YYMM & ")_" & selectedDate & ".csv" ' 當月小型期貨

    '========================
    ' (3) 依序載入 3 份議價資料到 [原資料]
    '     A:J / K:T / U:AD
    '========================
    colOffset = 0
    
    For Each variantItem In csvFileVariants
        
        fileName = CStr(variantItem)
        fullFilePath = folderPath & "\" & fileName
        
        If Dir(fullFilePath) <> "" Then
            
            Set wbCSV = Workbooks.Open(fullFilePath)
            
            ' 將 CSV 資料貼到對應區塊（從第2列開始）
            wbCSV.Sheets(1).UsedRange.Copy
            wsRaw.Cells(2, 1 + colOffset).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            
            wbCSV.Close SaveChanges:=False
            
        Else
            MsgBox "檔案不存在：" & fullFilePath, vbExclamation
        End If
        
        colOffset = colOffset + 10
        
    Next variantItem

    '========================
    ' (4) 載入成交資料檔到 [原資料] AE 起
    '========================
    tradeFileName = selectedStock & "_成交_" & selectedDate & ".csv"
    tradeFullPath = folderPath & "\成交\" & tradeFileName

    If Dir(tradeFullPath) <> "" Then
        
        Set wbCSV = Workbooks.Open(tradeFullPath)
        
        ' 將成交資料貼到 AE 欄起（從第2列開始）
        wbCSV.Sheets(1).UsedRange.Copy
        wsRaw.Cells(2, 31).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        wbCSV.Close SaveChanges:=False

        ' 將 AE 欄時間轉成 hhmmss 整數格式
        For Each cell In wsRaw.Range("AE3:AE" & Trim(Str(TheLastRow("AE"))))
            If IsNumeric(cell.Value) And cell.Value > 0 Then
                cell.Value = Format(cell.Value, "hhmmss")
                cell.NumberFormat = "0"
            End If
        Next cell
        
    Else
        MsgBox "成交資料檔案不存在：" & tradeFullPath, vbExclamation
    End If

    '========================
    ' (5) 將成交資料關鍵欄搬到固定區
    '     AE → AT   AH → AU   AJ → AV
    '========================
    wsRaw.Range("AE:AE").Copy
    wsRaw.Range("AT1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    wsRaw.Range("AH:AH").Copy
    wsRaw.Range("AU1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    wsRaw.Range("AJ:AJ").Copy
    wsRaw.Range("AV1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '========================
    ' (6) 將商品名稱填到 [看盤室]
    '     A2 / A3 / A4
    '========================
    With ThisWorkbook.Sheets("看盤室")
        .Range("A2").Value = wsRaw.Range("A2").Value    ' 現貨商品名稱
        .Range("A3").Value = wsRaw.Range("K2").Value    ' 期貨商品名稱
        .Range("A4").Value = wsRaw.Range("U2").Value    ' 小型期貨商品名稱
    End With

    '========================
    ' (7) 製作 20秒1點江波圖資料
    '========================
    ProcessSampledTradeData

    '========================
    ' (8) [原資料] → [資料整備]
    '========================
    Build_AllMarkets_TimeAligned

    '========================
    ' (9) [原資料] → [即時成交] K:P
    '========================
    TransferTradeData

    '========================
    ' (10) 初始化播放時間並觸發畫面更新
    '========================
    Set 全域變數.runTime = New clsPlayTime
    全域變數.runTime.Time = ThisWorkbook.Sheets("看盤室").Range("E31").Value

    MsgBox "資料下載完成！", vbInformation

End Sub
'回傳指定欄位之最底下一列有值之列號
Private Function TheLastRow(ColumnAssigned) 'ColumnAssigned=指定欄位(數字或字母)
    If IsNumeric(ColumnAssigned) = True Then
        TheLastRow = Cells(ActiveSheet.RowS.Count, ColumnAssigned).End(xlUp).row
    End If
    If IsNumeric(ColumnAssigned) = False Then
        TheLastRow = ActiveSheet.Range(ColumnAssigned & Trim(Str(ActiveSheet.RowS.Count))).End(xlUp).row
    End If
End Function


' 製作江波圖資料（每 JiangBoStepSec 秒 1 點）
Sub ProcessSampledTradeData()

    Dim wsRaw As Worksheet
    Dim lastRow As Long, lastSampleRow As Long
    Dim sampleRow As Long
    Dim CurrentTime As Long
    Dim startRow As Long, endRow As Long
    Dim sumVolume As Double
    Dim maxTradeRow As Long

    Set wsRaw = ThisWorkbook.Sheets("原資料")
    
    ' 清除舊數據
    wsRaw.Range("AO:AS").ClearContents
    wsRaw.Range("AO2") = "時間"
    wsRaw.Range("AP2") = "成交"
    wsRaw.Range("AQ2") = JiangBoStepSec & "秒量"
    
    ' 產生 AO 欄時間戳
    GenerateTimeStamps
    
    ' 找到原始成交時間最後一列（AT欄）
    lastRow = wsRaw.Cells(wsRaw.RowS.Count, 46).End(xlUp).row
    
    ' 找到 AO 欄最後一列（取樣時間軸）
    lastSampleRow = wsRaw.Cells(wsRaw.RowS.Count, 41).End(xlUp).row
    
    sampleRow = 4   ' 從 AO4 開始
    
    ' 由 AO 欄的時間軸逐點取樣
    Do While sampleRow <= lastSampleRow
        
        CurrentTime = wsRaw.Cells(sampleRow, 41).Value
        
        ' AP 欄（成交價）：找 <= CurrentTime 的最後一筆成交
        maxTradeRow = FindLastRow(wsRaw, 3, 46, CurrentTime)
        If maxTradeRow > 0 Then
            wsRaw.Cells(sampleRow, 42).Value = wsRaw.Cells(maxTradeRow, 47).Value
        ElseIf sampleRow > 4 Then
            wsRaw.Cells(sampleRow, 42).Value = wsRaw.Cells(sampleRow - 1, 42).Value
        Else
            wsRaw.Cells(sampleRow, 42).Value = 0
        End If

        ' AQ 欄（區間成交量）：前一取樣點 ~ 本取樣點
        If sampleRow = 4 Then
            startRow = 3
        Else
            startRow = FindLastRow(wsRaw, 3, 46, wsRaw.Cells(sampleRow - 1, 41).Value)
        End If
        
        endRow = maxTradeRow

        If startRow > 0 And endRow >= startRow Then
            sumVolume = Application.WorksheetFunction.Sum(wsRaw.Range("AV" & startRow & ":AV" & endRow))
        Else
            sumVolume = 0
        End If

        wsRaw.Cells(sampleRow, 43).Value = sumVolume

        sampleRow = sampleRow + 1
    Loop

    ' 清除顯示區
    wsRaw.Range("AR:AS").ClearContents

End Sub

' 在 [原資料] 的 AO 欄填入江波圖所需取樣時間點
Sub GenerateTimeStamps()

    Dim wsRaw As Worksheet
    Dim n As Long
    Dim baseTime As Date, endTime As Date, t As Date
    
    Set wsRaw = ThisWorkbook.Sheets("原資料")
    wsRaw.Range("AO:AO").ClearContents                               ' 清除 AO 欄

    n = 3                                                            ' 從第 3 列開始填

    ' 以現貨交易時間為江波圖時間軸
    baseTime = TimeSerial(sTimeStart \ 10000, (sTimeStart \ 100) Mod 100, sTimeStart Mod 100) ' 09:00:00
    endTime = TimeSerial(sTimeEnd \ 10000, (sTimeEnd \ 100) Mod 100, sTimeEnd Mod 100)         ' 13:30:00

    t = baseTime

    Do While t <= endTime
        wsRaw.Cells(n, 41).Value = Hour(t) * 10000 + Minute(t) * 100 + Second(t)   ' AO 欄 = hhmmss
        n = n + 1
        t = DateAdd("s", JiangBoStepSec, t)                                          ' 每 JiangBoStepSec 秒一點
    Loop

End Sub

' 取得下一期的期貨月份代碼
Private Function GetNextMonthCode(currentYYMM As String) As String
    Dim yearPart As String
    Dim monthPart As String
    Dim nextMonth As Integer

    ' 分割 YYMM 為年份和月份
    yearPart = Left(currentYYMM, 2) ' 前兩位為年份
    monthPart = Right(currentYYMM, 2) ' 後兩位為月份
    nextMonth = CInt(monthPart) + 1 ' 加 1 表示下一期

    ' 處理月份進位
    If nextMonth > 12 Then
        nextMonth = 1
        yearPart = CStr(CInt(yearPart) + 1) ' 進位至下一年
    End If

    ' 返回新的 YYMM
    GetNextMonthCode = yearPart & Format(nextMonth, "00")
End Function

Sub cmdPlay_Click()
    
    Call 交易展示.cmdPlay_Click

End Sub

Sub cmdPause_Click()
    SimRunning = False  ' 停止播放
End Sub

Sub TestSleep()
    Dim i As Integer
    For i = 1 To 10
        Sheet1.Cells(32, 1) = i
        DoEvents ' 確保 UI 更新不卡住
        Sleep 500 ' 暫停 500 毫秒（0.5 秒）
    Next i
End Sub
