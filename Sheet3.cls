VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' **當使用者選擇 AO 欄範圍時，自動將選擇範圍轉到 AP 欄**
    If Not Intersect(Target, Me.Range("AO3:AO797")) Is Nothing Then
        Application.EnableEvents = False ' 暫時關閉事件，避免無限循環
        Me.Cells(Target.row, 42).Select ' **將選取範圍轉到 AP 欄 (成交價)**
        Application.EnableEvents = True ' 重新開啟事件
    End If
End Sub

