Attribute VB_Name = "Module2"
Sub sheet_del()
  Dim sheet_cnt As Integer
  For sheet_cnt = Worksheets.Count To 2 Step -1
    Application.DisplayAlerts = False ' ���b�Z�[�W���\��
    Sheets(sheet_cnt).Delete
    Application.DisplayAlerts = True  ' ���b�Z�[�W��\��
  Next
End Sub
