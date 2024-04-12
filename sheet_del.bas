Attribute VB_Name = "Module2"
Sub sheet_del()
  Dim sheet_cnt As Integer
  For sheet_cnt = Worksheets.Count To 2 Step -1
    Application.DisplayAlerts = False ' メッセージを非表示
    Sheets(sheet_cnt).Delete
    Application.DisplayAlerts = True  ' メッセージを表示
  Next
End Sub
