Attribute VB_Name = "Module1"
'binary�f�[�^�i�[�ϐ�
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
    '�}�N�����s����ʂ��X�V���Ȃ��悤�ɂ���
    Application.ScreenUpdating = False
    
    '�����ꂽ�{�^���̖��̂��擾
    Dim buuton_name As Integer
    button_name = CInt(Worksheets("TOP").Buttons(Application.Caller).Text)
    
    '�������擾
    Dim sheets_num As Integer
    sheets_num = Worksheets("TOP").Cells(button_name, 6)
    
    '�t�@�C���̃t���p�X���i�[�ϐ�
    ReDim vFilePath(0 To sheets_num - 1)
    
    Dim buf As String
    cnt = 0
    Const Path As String = "D:\�v���C�x�[�g\excel\BMP�t�@�C���ǂݎ��\02_���ʕ�"
    Dim second_path As String
    second_path = Worksheets("TOP").Cells(button_name, 4)
    buf = Dir(Path & second_path)
    Do While buf <> ""
        vFilePath(cnt) = Path & second_path & buf
        buf = Dir()
        cnt = cnt + 1
    Loop
    
    '�摜���������[�v
    For cnt = 0 To sheets_num - 1
      '�t�@�C�����Ȃ�������I��
      If IsEmpty(vFilePath(cnt)) = True Then
        Application.ScreenUpdating = True
        Exit Sub
      End If
      '�t�@�C���T�C�Y�i�[�ϐ�
      Dim nFileLen As Long
      '�t�@�C���̃T�C�Y���o�C�g���Ŏ擾
      nFileLen = FileLen(vFilePath(cnt))
      '�t�@�C���T�C�Y��0�o�C�g�̏ꍇ�͏����I��
      If nFileLen = 0 Then
          '�}�N�����s����ʂ��X�V����悤�ɂ���
          Application.ScreenUpdating = True
          End
      End If
    
      '�I�����ꂽbmp�摜���o�C�i���f�[�^�Ŏ擾
      Dim iFile As Integer
    
      'Open���\�b�h�Ŏw�肷��A�g�p�\�ȃt�@�C���ԍ���Ԃ��܂�
      iFile = FreeFile
    
      '�w�肳�ꂽ�t�@�C�����擾�����t�@�C���ԍ��Ƃ��ăo�C�i�����[�h�ŊJ��
      Open vFilePath(cnt) For Binary As #iFile
    
      '�t�@�C���T�C�Y���̔z��𐶐�
      ReDim bData(0 To nFileLen - 1)
      '��1����:�t�@�C��No ��3����:�f�[�^��ǂݍ��񂾓��e���i�[����ϐ�
      Get #iFile, , bData
      '�t�@�C�������
      Close #iFile
  
      '0x42 == B ==66,0x4D == M == 77(16�i == �����R�[�h == 10�i)
      If Not (bData(0) = 66 And bData(1) = 77) Then
          MsgBox "bmp�摜��I�����Ă��������B"
          Exit Sub
      End If
      
      '�V�[�g�ɕ`��
      bmp_text_output
      
    Next
    
      '�}�N�����s����ʂ��X�V����悤�ɂ���
      Application.ScreenUpdating = True
End Sub
Sub bmp_text_output()
    '�}�N�����s����ʂ��X�V���Ȃ��悤�ɂ���
    Application.ScreenUpdating = False
    
    '�I�u�W�F�N�g�^:Worksheet���i�[����ϐ�
    Dim ExportSheet As Worksheet
    '���[�N�V�[�g��ǉ�����
    Set ExportSheet = Sheets.add(After:=Sheets(Sheets.Count))
    '���[�N�V�[�g�̖��O��ύX
    Dim SheetsName As Integer
    SheetsName = Sheets.Count - 2 + 1
    If SheetsName < 10 Then
        ExportSheet.name = "00" & CStr(SheetsName)
    ElseIf SheetsName < 100 Then
        ExportSheet.name = "0" & CStr(SheetsName)
    Else
        ExportSheet.name = CStr(SheetsName)
    End If
    
    '�摜��[px]�i�[�ϐ�
    Dim ImageRow As Long
    '�摜�̍���[px]�i�[�ϐ�
    Dim ImageCol As Long
    '�J�E���^
    Dim k As Long
    'R�f�[�^�i�[�ϐ�
    Dim R As Integer
    'G�f�[�^�i�[�ϐ�
    Dim G As Integer
    'B�f�[�^�i�[�ϐ�
    Dim B As Integer
    '�Z���̕�[px]���͒l�i�[�ϐ�
    Dim input_height As Integer
    '�Z���̍���[px]���͒l�i�[�ϐ�
    Dim input_width As Integer
    
    'TOP�ɓ��͂��ꂽ�Z���̎w�蕝���擾
    input_height = Worksheets("TOP").Cells(11, 8)
    'TOP�ɓ��͂��ꂽ�Z���̍�������擾
    input_width = Worksheets("TOP").Cells(12, 8)
    
    '�摜�̕����擾
    ImageCol = bData(18) + (bData(19) * 256) + (bData(20) * 256) + (bData(21) * 256)
    '�摜�̍������擾
    ImageRow = bData(22) + (bData(23) * 256) + (bData(24) * 256) + (bData(25) * 256)
    
    '�F�z��̃T�C�Y���擾
    Dim elm As Long
    'RGB�f�[�^�i�[�z��
    Dim RGBDatas() As Byte
    '���[�v�J�E���^
    Dim i As Long
    
    'bmp�t�@�C���F24bit�̏ꍇ
    If bData(28) = 24 Then
    '�摜�̕��~�摜�̍����~3(3�F���̃f�[�^���i�[���邽�߁~3)
    elm = ImageCol * ImageRow * 3
    '0�Ԗڂ͋�ɂ���
    ReDim RGBDatas(elm)
    '�o�C�i���f�[�^�̃f�[�^�������i�[
    'RGB�i�[�ϐ��̃T�C�Y�����[�v
    For i = 1 To UBound(RGBDatas)
        '�f�[�^����54�o�C�g�ڈȍ~�̂���+53����
        RGBDatas(i) = bData(i + 53)
    Next i
    '�J�E���^��1�ŏ�����
    k = 1
    'ExportSheet�ɑ΂��ĘA���ŏ��������s����
    With ExportSheet
        '�Z���̏������N���A����(�F�A�r���A�����t�������Ȃ�)
        .Cells.ClearFormats
        '�Z���̗�ƍs�̕�������ݒ肷��
        .Range(Columns(1), Columns(ImageCol)).ColumnWidth = 0.077 * input_height
        .Range(Rows(1), Rows(ImageRow)).RowHeight = 0.75 * input_width
        
        '�摜�̍��������[�v(y)
        For i = ImageRow To 1 Step -1
            '�摜�̕������[�v(x)
            For j = 1 To ImageCol
            'R�f�[�^���擾
            R = Int(RGBDatas((k - 1) * 3 + 1 + 2))
            'G�f�[�^���擾
            G = Int(RGBDatas((k - 1) * 3 + 1 + 1))
            'B�f�[�^���擾
            B = Int(RGBDatas((k - 1) * 3 + 1))
            
            '�u�Z���̏������������܂��v���p
            Dim format_cnt
            Dim s
            On Error Resume Next
            For format_cnt = ActiveWorkbook.Styles.Count To 1 Step -1
              s = ActiveWorkbook.Styles(i).Delete
            Next
            
            '�Z���̐F���w��
            .Cells(i, j).Interior.Color = RGB(R, G, B)
            '�J�E���^���J�E���g�A�b�v
            k = k + 1
            Next j
        Next i
    
    End With
    
    ElseIf bData(28) = 8 Then
    '�摜�̕��~�摜�̍����~3(3�F���̃f�[�^���i�[���邽�߁~3)
    elm = ImageCol * ImageRow
    'RGB�f�[�^�i�[�z��
    '0�Ԗڂ͋�ɂ���
    ReDim RGBDatas(elm)
    '�o�C�i���f�[�^�̃f�[�^�������i�[
    'RGB�i�[�ϐ��̃T�C�Y�����[�v
    For i = 1 To UBound(RGBDatas)
        '�f�[�^����54�o�C�g�ڈȍ~�̂���+53����
        RGBDatas(i) = bData(i + 1077)
    Next i
    
    'RGB�p���b�g�f�[�^
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
    
    '�J�E���^��1�ŏ�����
    k = 1
        'ExportSheet�ɑ΂��ĘA���ŏ��������s����
    With ExportSheet
        '�Z���̏������N���A����(�F�A�r���A�����t�������Ȃ�)
        .Cells.ClearFormats
        '�Z���̗�ƍs�̕�������ݒ肷��
        .Range(Columns(1), Columns(ImageCol)).ColumnWidth = 0.077 * input_height
        .Range(Rows(1), Rows(ImageRow)).RowHeight = 0.75 * input_width
        
        '�摜�̍��������[�v(y)
        For i = ImageRow To 1 Step -1
            '�摜�̕������[�v(x)
            For j = 1 To ImageCol
            '�Z���̐F���w��
            'R�f�[�^���擾
            R = Int(palette(RGBDatas(k)).rgbRed)
            'G�f�[�^���擾
            G = Int(palette(RGBDatas(k)).rgbGreen)
            'B�f�[�^���擾
            B = Int(palette(RGBDatas(k)).rgbBule)
            '�Z���̐F���w��
            .Cells(i, j).Interior.Color = RGB(R, G, B)
            '�J�E���^���J�E���g�A�b�v
            k = k + 1
            Next j
        Next i
    
    End With
    
    Else
    
    End If
    
    '�}�N�����s����ʂ��X�V����悤�ɂ���
    Application.ScreenUpdating = True
End Sub
Sub Music_Start()
    Dim clsSound As clsSound
    Set clsSound = New clsSound
    '�����ꂽ�{�^���̖��̂��擾
    Dim buuton_name As Integer
    button_name = CInt(Worksheets("TOP").Buttons(Application.Caller).Text)
    '���y�Đ�
    clsSound.SoundFile = "D:\�v���C�x�[�g\excel\BMP�t�@�C���ǂݎ��\02_���ʕ�\music" & Worksheets("TOP").Cells(button_name, 9)
    clsSound.GetLength '�Đ����Ԏ擾(Play�̑O�ɃR�[�����Ȃ��ƍĐ����Ȃ�)
    clsSound.Play      '���y�Đ�
    
    '�V�[�g�̈ړ�����
    Playback
  
    '���y��~
    clsSound.StopSound
  
    Set clsSound = Nothing
End Sub
Sub Playback()
   '�����ꂽ�{�^���̖��̂��擾
   Dim buttonName As Integer
   buttonName = CInt(Worksheets("TOP").Buttons(Application.Caller).Text)
   Dim sheetNum As Integer
   sheetNum = Worksheets("TOP").Cells(buttonName, 6)
   Dim sheetLoopCnt As Integer
   sheetLoopCnt = Worksheets("TOP").Cells(buttonName, 8)
    
    '�V�[�g�ړ�����
    Dim i As Integer, j As Integer
    For i = 1 To sheetLoopCnt
      For j = 1 To sheetNum
        Sheets(Format(j, "000")).Select
        Application.Wait [Now()] + 83 / 86400000
      Next
    Next
End Sub
Sub test()
  Workbooks.Open "D:\excel\BMP�t�@�C���ǂݎ��\02_���ʕ�\book.xlsx"
End Sub
