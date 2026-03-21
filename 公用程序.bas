Attribute VB_Name = "公用程序"
Public startRow As Long, endRow As Long, lastRow As Long
Public nextUpdateTime  As Date

Sub 安排視窗()
    Dim appWidth As Double
    Dim appHeight As Double
    Dim leftWidth As Double
    Dim rightWidth As Double
    Dim windowHeight As Double

    ' 獲取 Excel 視窗的寬度和高度
    appWidth = Application.Width
    appHeight = Application.Height

    ' 計算左邊和右邊的寬度，以及視窗高度
    leftWidth = appWidth * 0.25
    rightWidth = appWidth * 0.75
    windowHeight = appHeight * 0.9

    ' 若有Windows("Book1")關閉之
    On Error Resume Next
    If Not Application.Windows("Book1") Is Nothing Then
        Application.Windows("Book1").Close
    End If

    ' 確保已開啟兩個視窗
    If Windows.Count < 2 Then
        ActiveWindow.NewWindow
    End If

    ' 設置左邊的視窗 (:1)
    With Windows(1)
        .WindowState = xlNormal ' 設置為非最大化
        .Top = 0
        .Left = 0
        .Width = leftWidth
        .Height = windowHeight
        .Activate
        Sheets("即時成交").Select
        Range("A2").Select
        .FreezePanes = True
    End With

    ' 設置右邊的視窗 (:2)
    With Windows(2)
        .WindowState = xlNormal ' 設置為非最大化
        .Top = 0
        .Left = leftWidth
        .Width = rightWidth
        .Height = windowHeight
        .Activate
        Sheets("看盤室").Select
        Cells(1, 1).Select
    End With
End Sub

Sub ReFreshScreen()
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    DoEvents ' 立即更新畫面
End Sub

Sub ToKeyMode()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("看盤室")

    ' 1. 切換布林變數的狀態 (True 變 False，False 變 True)
    KeyModeOnOff = Not KeyModeOnOff
    
    ' 2. 根據現在的狀態，決定接上還是切斷熱鍵神經
    If KeyModeOnOff = True Then
        ' ?? 啟用模式：綁定熱鍵到指定的巨集
        Application.OnKey "%{Left}", "Backward"
        Application.OnKey "%{Right}", "Forward"
        Application.EnableEvents = False
        ws.Cells(29, 5).Value = "熱鍵ON"
        Application.EnableEvents = True
    Else
        ' ?? 停用模式：省略第二個參數，把熱鍵控制權還給 Excel 預設功能
        Application.OnKey "%{Left}", ""
        Application.OnKey "%{Right}", ""
        Application.EnableEvents = False
        ws.Cells(29, 5).Value = "熱鍵OFF"
        Application.EnableEvents = True
    End If
    ActiveCell.Select
End Sub


Private Sub cmdForward_Click()
    Call 交易展示.cmdForward_Click
End Sub

Private Sub cmdBackward_Click()
    Call 交易展示.cmdBackward_Click
End Sub

'初始化各廣域變數 並定義快速鍵功能
Sub InitializeVariable()

    Set runTime = New clsPlayTime
    runTime.Time = Sheets("看盤室").Range("E31").Value
    
    OneStep = Sheets("看盤室").Range("OneStep").Value
    
    ToKeyMode

End Sub


Public Sub Test_RunTime_UpdateDisplay()

    ' 若 runTime 尚未建立，先建立並掛接事件
    If runTime Is Nothing Then
        Set runTime = New clsPlayTime
        ThisWorkbook.HookRunTime
    End If

    ' 指定測試時間（例：09:00:14）
    runTime.Time = 90014

End Sub
