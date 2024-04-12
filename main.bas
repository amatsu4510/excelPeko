Attribute VB_Name = "Module1"
'binaryデータ格納変数
Dim bData() As Byte
Dim cnt As Integer
Dim vFilePath() As Variant
Type rgbquad
    rgbBule     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

Sub bmp_read()
    'マクロ実行中画面を更新しないようにする
    Application.ScreenUpdating = False
    
    '押されたボタンの名称を取得
    Dim buuton_name As Integer
    button_name = CInt(Worksheets("TOP").Buttons(Application.Caller).Text)
    
    '枚数を取得
    Dim sheets_num As Integer
    sheets_num = Worksheets("TOP").Cells(button_name, 6)
    
    'ファイルのフルパスを格納変数
    ReDim vFilePath(0 To sheets_num - 1)
    
    Dim buf As String
    cnt = 0
    Const Path As String = "D:\プライベート\excel\BMPファイル読み取り\02_成果物"
    Dim second_path As String
    second_path = Worksheets("TOP").Cells(button_name, 4)
    buf = Dir(Path & second_path)
    Do While buf <> ""
        vFilePath(cnt) = Path & second_path & buf
        buf = Dir()
        cnt = cnt + 1
    Loop
    
    '画像枚数分ループ
    For cnt = 0 To sheets_num - 1
      'ファイルがなかったら終了
      If IsEmpty(vFilePath(cnt)) = True Then
        Application.ScreenUpdating = True
        Exit Sub
      End If
      'ファイルサイズ格納変数
      Dim nFileLen As Long
      'ファイルのサイズをバイト数で取得
      nFileLen = FileLen(vFilePath(cnt))
      'ファイルサイズが0バイトの場合は処理終了
      If nFileLen = 0 Then
          'マクロ実行中画面を更新するようにする
          Application.ScreenUpdating = True
          End
      End If
    
      '選択されたbmp画像をバイナリデータで取得
      Dim iFile As Integer
    
      'Openメソッドで指定する、使用可能なファイル番号を返します
      iFile = FreeFile
    
      '指定されたファイルを取得したファイル番号としてバイナリモードで開く
      Open vFilePath(cnt) For Binary As #iFile
    
      'ファイルサイズ分の配列を生成
      ReDim bData(0 To nFileLen - 1)
      '第1引数:ファイルNo 第3引数:データを読み込んだ内容を格納する変数
      Get #iFile, , bData
      'ファイルを閉じる
      Close #iFile
  
      '0x42 == B ==66,0x4D == M == 77(16進 == 文字コード == 10進)
      If Not (bData(0) = 66 And bData(1) = 77) Then
          MsgBox "bmp画像を選択してください。"
          Exit Sub
      End If
      
      'シートに描画
      bmp_text_output
      
    Next
    
      'マクロ実行中画面を更新するようにする
      Application.ScreenUpdating = True
End Sub
Sub bmp_text_output()
    'マクロ実行中画面を更新しないようにする
    Application.ScreenUpdating = False
    
    'オブジェクト型:Worksheetを格納する変数
    Dim ExportSheet As Worksheet
    'ワークシートを追加する
    Set ExportSheet = Sheets.add(After:=Sheets(Sheets.Count))
    'ワークシートの名前を変更
    Dim SheetsName As Integer
    SheetsName = Sheets.Count - 2 + 1
    If SheetsName < 10 Then
        ExportSheet.name = "00" & CStr(SheetsName)
    ElseIf SheetsName < 100 Then
        ExportSheet.name = "0" & CStr(SheetsName)
    Else
        ExportSheet.name = CStr(SheetsName)
    End If
    
    '画像幅[px]格納変数
    Dim ImageRow As Long
    '画像の高さ[px]格納変数
    Dim ImageCol As Long
    'カウンタ
    Dim k As Long
    'Rデータ格納変数
    Dim R As Integer
    'Gデータ格納変数
    Dim G As Integer
    'Bデータ格納変数
    Dim B As Integer
    'セルの幅[px]入力値格納変数
    Dim input_height As Integer
    'セルの高さ[px]入力値格納変数
    Dim input_width As Integer
    
    'TOPに入力されたセルの指定幅を取得
    input_height = Worksheets("TOP").Cells(11, 8)
    'TOPに入力されたセルの高さ定を取得
    input_width = Worksheets("TOP").Cells(12, 8)
    
    '画像の幅を取得
    ImageCol = bData(18) + (bData(19) * 256) + (bData(20) * 256) + (bData(21) * 256)
    '画像の高さを取得
    ImageRow = bData(22) + (bData(23) * 256) + (bData(24) * 256) + (bData(25) * 256)
    
    '色配列のサイズを取得
    Dim elm As Long
    'RGBデータ格納配列
    Dim RGBDatas() As Byte
    'ループカウンタ
    Dim i As Long
    
    'bmpファイル色24bitの場合
    If bData(28) = 24 Then
    '画像の幅×画像の高さ×3(3色分のデータを格納するため×3)
    elm = ImageCol * ImageRow * 3
    '0番目は空にする
    ReDim RGBDatas(elm)
    'バイナリデータのデータ部情報を格納
    'RGB格納変数のサイズ分ループ
    For i = 1 To UBound(RGBDatas)
        'データ部が54バイト目以降のため+53する
        RGBDatas(i) = bData(i + 53)
    Next i
    'カウンタを1で初期化
    k = 1
    'ExportSheetに対して連続で処理を実行する
    With ExportSheet
        'セルの書式をクリアする(色、罫線、条件付き書式など)
        .Cells.ClearFormats
        'セルの列と行の幅高さを設定する
        .Range(Columns(1), Columns(ImageCol)).ColumnWidth = 0.077 * input_height
        .Range(Rows(1), Rows(ImageRow)).RowHeight = 0.75 * input_width
        
        '画像の高さ分ループ(y)
        For i = ImageRow To 1 Step -1
            '画像の幅分ループ(x)
            For j = 1 To ImageCol
            'Rデータを取得
            R = Int(RGBDatas((k - 1) * 3 + 1 + 2))
            'Gデータを取得
            G = Int(RGBDatas((k - 1) * 3 + 1 + 1))
            'Bデータを取得
            B = Int(RGBDatas((k - 1) * 3 + 1))
            
            '「セルの書式が多すぎます」回避用
            Dim format_cnt
            Dim s
            On Error Resume Next
            For format_cnt = ActiveWorkbook.Styles.Count To 1 Step -1
              s = ActiveWorkbook.Styles(i).Delete
            Next
            
            'セルの色を指定
            .Cells(i, j).Interior.Color = RGB(R, G, B)
            'カウンタをカウントアップ
            k = k + 1
            Next j
        Next i
    
    End With
    
    ElseIf bData(28) = 8 Then
    '画像の幅×画像の高さ×3(3色分のデータを格納するため×3)
    elm = ImageCol * ImageRow
    'RGBデータ格納配列
    '0番目は空にする
    ReDim RGBDatas(elm)
    'バイナリデータのデータ部情報を格納
    'RGB格納変数のサイズ分ループ
    For i = 1 To UBound(RGBDatas)
        'データ部が54バイト目以降のため+53する
        RGBDatas(i) = bData(i + 1077)
    Next i
    
    'RGBパレットデータ
    Dim palette() As rgbquad
    ReDim palette(255)
    k = 1
    For i = 0 To UBound(palette)
      palette(i).rgbBule = bData(k + 53)
      palette(i).rgbGreen = bData(k + 53 + 1)
      palette(i).rgbRed = bData(k + 53 + 2)
      palette(i).rgbReserved = bData(k + 53 + 3)
      k = k + 4
    Next
    
    'カウンタを1で初期化
    k = 1
        'ExportSheetに対して連続で処理を実行する
    With ExportSheet
        'セルの書式をクリアする(色、罫線、条件付き書式など)
        .Cells.ClearFormats
        'セルの列と行の幅高さを設定する
        .Range(Columns(1), Columns(ImageCol)).ColumnWidth = 0.077 * input_height
        .Range(Rows(1), Rows(ImageRow)).RowHeight = 0.75 * input_width
        
        '画像の高さ分ループ(y)
        For i = ImageRow To 1 Step -1
            '画像の幅分ループ(x)
            For j = 1 To ImageCol
            'セルの色を指定
            'Rデータを取得
            R = Int(palette(RGBDatas(k)).rgbRed)
            'Gデータを取得
            G = Int(palette(RGBDatas(k)).rgbGreen)
            'Bデータを取得
            B = Int(palette(RGBDatas(k)).rgbBule)
            'セルの色を指定
            .Cells(i, j).Interior.Color = RGB(R, G, B)
            'カウンタをカウントアップ
            k = k + 1
            Next j
        Next i
    
    End With
    
    Else
    
    End If
    
    'マクロ実行中画面を更新するようにする
    Application.ScreenUpdating = True
End Sub
Sub Music_Start()
    Dim clsSound As clsSound
    Set clsSound = New clsSound
    '押されたボタンの名称を取得
    Dim buuton_name As Integer
    button_name = CInt(Worksheets("TOP").Buttons(Application.Caller).Text)
    '音楽再生
    clsSound.SoundFile = "D:\プライベート\excel\BMPファイル読み取り\02_成果物\music" & Worksheets("TOP").Cells(button_name, 9)
    clsSound.GetLength '再生時間取得(Playの前にコールしないと再生しない)
    clsSound.Play      '音楽再生
    
    'シートの移動処理
    Playback
  
    '音楽停止
    clsSound.StopSound
  
    Set clsSound = Nothing
End Sub
Sub Playback()
   '押されたボタンの名称を取得
   Dim buttonName As Integer
   buttonName = CInt(Worksheets("TOP").Buttons(Application.Caller).Text)
   Dim sheetNum As Integer
   sheetNum = Worksheets("TOP").Cells(buttonName, 6)
   Dim sheetLoopCnt As Integer
   sheetLoopCnt = Worksheets("TOP").Cells(buttonName, 8)
    
    'シート移動処理
    Dim i As Integer, j As Integer
    For i = 1 To sheetLoopCnt
      For j = 1 To sheetNum
        Sheets(Format(j, "000")).Select
        Application.Wait [Now()] + 83 / 86400000
      Next
    Next
End Sub
Sub test()
  Workbooks.Open "D:\excel\BMPファイル読み取り\02_成果物\book.xlsx"
End Sub
