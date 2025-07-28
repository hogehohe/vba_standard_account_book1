Option Explicit

'---------------------------------------------
'�p���]���C���V�[�g�Ŏg���萔
'---------------------------------------------
'1�}�X�̕b�����`
Const UNIT_TIME                         As Double = 0.1

'0�b�̗�
Const COLUMN_ZERO_NUM                   As Long = 6

'�s
'�M������[
Const ROW_RELIABILITY_TOP               As Long = 2
'�M�������[
Const ROW_RELIABILITY_BOTTOM            As Long = 7
'�p���_��[
Const ROW_POSTURE_SCORE_TOP             As Long = 9
'�p���_���[
Const ROW_POSTURE_SCORE_KOSHIMAGEOTTOM  As Long = 17

'A_�p���_
Const ROW_POSTURE_SCORE_KOBUSHIAGE      As Long = 12
'B_�p���_
Const ROW_POSTURE_SCORE_KOSHIMAGE       As Long = 14
'C_�p���_
Const ROW_POSTURE_SCORE_HIZAMAGE        As Long = 16

'---------------------------------------------
'�|�C���g�v�Z�V�[�g�̗�
'---------------------------------------------

'�p���_���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_ALL          As Long = 203

'����
Const COLUMN_MEASURE_SECTION            As Long = 204
'����
Const COLUMN_PREDICT_SECTION            As Long = 205
'���O���
Const COLUMN_REMOVE_SECTION             As Long = 206
'�������
Const COLUMN_FORCED_SECTION_TOTAL       As Long = 207
'���f�[�^
Const COLUMN_BASE_SCORE                 As Long = 208
'�p���f�_�ΐF
Const COLUMN_POSTURE_GREEN              As Long = 209
'�p���f�_���F
Const COLUMN_POSTURE_YELLOW             As Long = 210
'�p���f�_�ԐF
Const COLUMN_POSTURE_RED                As Long = 211

'����
Const COLUMN_MISSING_SECTION            As Long = 219

'���㋭�����
Const COLUMN_FORCED_SECTION_KOBUSHIAGE  As Long = 223
'���Ȃ��������
Const COLUMN_FORCED_SECTION_KOSHIMAGE   As Long = 228
'�G�Ȃ��������
Const COLUMN_FORCED_SECTION_HIZAMAGE    As Long = 233

'����A(����)���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_KOBUSHIAGE   As Long = 245
'����B(���Ȃ�)���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_KOSHIMAGE    As Long = 247
'����C(�G�Ȃ�)���ۑ�����Ă����
Const COLUMN_POSTURE_SCORE_HIZAMAGE     As Long = 249

'---------------------------------------------
'�p���]���C���V�[�g�@�֘A
'---------------------------------------------
'LIMIT_COLUMN�̐ݒ�l��3�̔{���Ƃ���K�v������
'30fps�~60�b�~9����16200
'�p���]���C���V�[�g��9�����Ɏ��̃V�[�g�𐶐�����
Const LIMIT_COLUMN                      As Long = 16200

Const SHEET_LIMIT_COLUMN                As Long = LIMIT_COLUMN + COLUMN_ZERO_NUM

'�����\���Z���̕�
Const TIME_WIDTH                        As Long = 30
'�����\���Z�������݂���s
Const TIME_ROW                          As Long = 22
'��ڂ̎����\���Z���̍��[
Const TIME_COLUMN_LEFT                  As Long = 22
'��ڂ̎����\���Z���̉E�[
Const TIME_COLUMN_RIGHT                 As Long = 51
'�f�[�^�����p�̃e�[�u���̉��[
Const BOTTOM_OF_TABLE                   As Long = 22

'�񕝗p�̗�
Private Enum widthSize
    Small = 1
    Medium = 2
    Large = 4
    LL = 6
End Enum

'�񕝒����{�^�����O
Const EXPANDBTN_NAME                    As String = "expandBtn"
Const REDUCEBTN_NAME                    As String = "reduceBtn"

'---------------------------------------------
'�������W���[���Ŏg�p����ϐ�
'---------------------------------------------
'�Đ��E��~�{�^���Ŏg�p
'�w�肵�����Ԃ��o�߂���Ə��������s����
Private ResTime As Date
Private scrollTime As Double


'�������ԒZ�k�̂��߁A�X�V���X�g�b�v
' ����1 �F�Ȃ�
' �߂�l�F�Ȃ�
Function stopUpdate()
    '�\���E�X�V���I�t�ɂ���
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Function


'�������ԒZ�k�̂��߁A�X�V�����X�^�[�g
' ����1 �F�Ȃ�
' �߂�l�F�Ȃ�
Function restartUpdate()
    '�\���E�X�V���I���ɖ߂�
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function


'�p���]���C���V�[�g
'�r���Ɖf���������I�[�g�t�B��
Sub autoFillTemplate()
    '���x���̗�
    Dim startColumnNum      As Long
    '10�b�̗�
    Dim unit10SecColumnNum  As Long

    '�ϐ���`
    Dim workTime            As Double
    Dim fps                 As Double
    Dim maxFrameNum         As Long
    Dim ruleLineColumnNum   As Long
    Dim ruleLineColumnAlf   As String

    startColumnNum = COLUMN_ZERO_NUM + 1
    unit10SecColumnNum = 10 / UNIT_TIME

    '��Ǝ��Ԃ��擾����
    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
        '�t���[�����[�g���擾
        fps = getFps()
        '�ŏI�s�̒l���擾
        maxFrameNum = getLastRow()
    End With
End Sub


'�r���̕��� "G2:EZ25"�͈̔͂��x�[�X�Ƃ���B
'����1�F���[�N�V�[�g
'����2�F�r�����`�悳���Ō�̗�
Private Sub autoFillLine(ws As Worksheet, endline As Long)
    Dim ruleLineColumnNum  As Long
    Dim ruleLineColumnAlf  As String

    ruleLineColumnNum = endline
    If ruleLineColumnNum > SHEET_LIMIT_COLUMN Then
        ruleLineColumnNum = SHEET_LIMIT_COLUMN
    End If
    Dim frame30Mod As Long
    frame30Mod = (ruleLineColumnNum + 21) Mod 30

    '�I�[�g�t�B���֐���RC�\�L�œ��삳������@��������Ȃ�����
    '�I�[�g�t�B���̏I������A���t�@�x�b�g�\�L�ɕϊ�
    ruleLineColumnAlf = Split(Cells(1, ruleLineColumnNum).Address(True, False), "$")(0)

    '���ꂢ�ɂ��Ă���I�[�g�t�B������(=�F�܂ŃR�s�[����邽��)
    Call clear(ws)

    ws.Range("G2:EZ21").AutoFill Destination:=Range("G2:" & ruleLineColumnAlf & 21), Type:=xlFillDefault
    ruleLineColumnAlf = Split(Cells(1, ruleLineColumnNum + 1).Address(True, False), "$")(0)
    Range(ruleLineColumnAlf & 2 & ":XFD21").Borders.LineStyle = xlLineStyleNone ' �㉺���E�̌r��������

End Sub

'���������ԃZ���ɑ}������
'����1�F���[�N�V�[�g
'����2�F��
'����3�F�ŏI��
Private Sub autoFillTime(ws As Worksheet, min As Long, endclm As Long)
    Dim tmp            As Long

    Dim boldcnt        As Long: boldcnt = 0
    Dim r              As Range

    Dim timeStr        As String

    Dim frame30Mod     As Long

    '�ϐ���`
    Dim i As Long
    tmp = endclm

    If 30 <= tmp - TIME_COLUMN_LEFT Then
        If tmp > LIMIT_COLUMN Then
            tmp = LIMIT_COLUMN
        End If
    End If

    '�I�[�g�t�B������ꏊ�ɃZ������������ƃG���[���o�邽��
    '�Z����������������
    ws.Range(Cells(TIME_ROW, 12), Cells(TIME_ROW, 16384)).clear

    For i = TIME_COLUMN_LEFT To SHEET_LIMIT_COLUMN Step TIME_WIDTH
        Set r = ws.Range(Cells(TIME_ROW, i), Cells(TIME_ROW, i + TIME_WIDTH - 1))
        boldcnt = boldcnt + 1

        '�Z���̏������܂Ƃ߂Đݒ肷��B
        With r
            .Merge True
            .Orientation = -90
            .ReadingOrder = xlContext
            .HorizontalAlignment = xlCenter
            .NumberFormatLocal = "hh:mm:ss"
            If boldcnt = 5 Then
                .Font.FontStyle = "bold"
                boldcnt = 0
            End If
        End With
    Next i

    timeStr = "00:" + CStr(min) + ":01"
     With ws.Range(Cells(TIME_ROW, TIME_COLUMN_LEFT), Cells(TIME_ROW, TIME_COLUMN_RIGHT))
        .Value = timeStr
    End With

    timeStr = "00:" + CStr(min) + ":02"
    With ws.Range(Cells(TIME_ROW, TIME_COLUMN_LEFT + TIME_WIDTH), Cells(TIME_ROW, TIME_COLUMN_RIGHT + TIME_WIDTH))
        .Value = timeStr
    End With

    frame30Mod = (tmp - TIME_COLUMN_LEFT) Mod TIME_WIDTH

    If frame30Mod Then
        tmp = tmp + TIME_WIDTH - frame30Mod
    End If

    If (TIME_COLUMN_LEFT + TIME_WIDTH) < tmp Then
        ws.Range( _
            Cells(TIME_ROW, TIME_COLUMN_LEFT), _
            Cells(TIME_ROW, TIME_COLUMN_RIGHT + TIME_WIDTH) _
        ).AutoFill _
        Destination:=Range( _
            Cells(TIME_ROW, TIME_COLUMN_LEFT), _
            Cells(TIME_ROW, tmp - 1) _
        ), _
        Type:=xlFillValues
    End If

End Sub


Sub test_cancel()
    Call paintPostureScore(1)
End Sub


'�P�ʎ��ԓ�����ł������p���_�E�M�����𒲂ׂăZ���ɐF��h��
'processingRange�@1:�I��͈́i�����I�ɋ������L�����Z���j 2:�S�� else:�����1�Z������
Sub paintPostureScore(processingRange As Long)
    '---------------------------------------------
    'RGB���w�肷�邽�߂̕ϐ����`
    '---------------------------------------------
    '�M����
    Dim colorMeasureSection    As String '���F
    Dim colorPredictSection    As String '���F
    Dim colorMissingSection    As String '�s���N
    Dim colorForcedSection     As String '�F
    Dim colorRemoveSection     As String '�O���[

    '�p���_
    Dim colorResultGreen       As String '�ΐF
    Dim colorResultYellow      As String '���F
    Dim colorResultRed         As String '�ԐF
    Dim colorResultGlay        As String '�O���[
    Dim colorResultWhite       As String '���F 20221219_����

    '---------------------------------------------
    '�ϐ��ɐF���Z�b�g
    '---------------------------------------------
    '1:����A2:����A3:�����A4:�����A5:���O
    '�M����
    colorMeasureSection = RGB(0, 176, 240)   '���F
    colorPredictSection = RGB(252, 246, 0)   '���F
    colorMissingSection = RGB(255, 124, 128) '�s���N
    colorForcedSection  = RGB(0, 51, 204)    '�F
    colorRemoveSection  = RGB(191, 191, 191) '�O���[
    '�p���_
    colorResultGreen    = RGB(0, 176, 80)    '�ΐF
    colorResultYellow   = RGB(255, 192, 0)   '���F
    colorResultRed      = RGB(192, 0, 0)     '�ԐF
    colorResultGlay     = RGB(191, 191, 191) '�O���[
    colorResultWhite    = RGB(255, 255, 255) '���F

    '---------------------------------------------
    '�z��
    '---------------------------------------------
    '�|�C���g�v�Z�V�[�g�̎p���_��ۊ�
    Dim postureScoreDataArray()    As Long

    '~~~~~~~~~~~~~~~~�ǉ�~~~~~~~~~~~~~~~~~~~

    '����A,B,C,D,E���Ƃ̎p���_��ۊ�
    Dim postureScoreDataArray_A()  As Long
    Dim postureScoreDataArray_B()  As Long
    Dim postureScoreDataArray_C()  As Long

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '0�`1�_�̃t���[���������ꂼ�ꍇ�v

    '~~~~~~~~~~~~~~~~�ǉ�~~~~~~~~~~~~~~~~~~~
    '����A����C�̃t���[���������ꂼ�ꍇ�v
    Dim postureScoreCounterArray_A(0 To 1)      As Long
    Dim postureScoreCounterArray_B(0 To 1)      As Long
    Dim postureScoreCounterArray_C(0 To 1)      As Long

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '�|�C���g�v�Z�V�[�g�̐M������ۊ�
    '1:����A2:����A3:����
    Dim reliabilityDataArray()     As Long
    '�M���� 1 ~ 3 �̃t���[���������ꂼ�ꍇ�v
    Dim reliabilityCounterArray(3) As Long

    '---------------------------------------------
    '���̑��̕ϐ�
    '---------------------------------------------
    '�|�C���g�v�Z�V�[�g�ő�s���̕ϐ���`
    Dim maxRowNum               As Long

    '�ϐ���`
    Dim i                       As Long
    Dim j                       As Long

    Dim fps                     As Double

    '�P�ʎ��Ԃ̌J��Ԃ������̊J�n�I���n�_���`
    Dim wholeStart              As Long
    Dim wholeEnd                As Long

    '�p���_�ꎞ�L���p�̕ϐ�
    Dim postureScoreFlag        As Long

    '~~~~~~~~~~~~~~~~�ǉ�~~~~~~~~~~~~~~~~~~~
    Dim postureScoreFlag_A      As Long
    Dim postureScoreFlag_B      As Long
    Dim postureScoreFlag_C      As Long

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '�P�ʎ��Ԃ̒��ň�ԑ����p���_��ۊ�
    Dim mostOftenPostureScore   As Long

    '~~~~~~~~~~~~~~~~�ǉ�~~~~~~~~~~~~~~~~~~~

    'A����E�̎p���_�ꎞ�L���p�̕ϐ�
    Dim mostOftenPostureScore_A As Long
    Dim mostOftenPostureScore_B As Long
    Dim mostOftenPostureScore_C As Long

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '�M�����ꎞ�L���p�̕ϐ�
    Dim reliabilityFlag         As Long
    '�P�ʎ��Ԃ̒��ň�ԑ����M������ۊ�
    Dim mostOftenReliability    As Long

    '���y�[�W�ɂ�������
    Dim thisPageLimit           As Long
    thisPageLimit = LIMIT_COLUMN

    '�O�̃y�[�W�̍ŏI���ۑ�����
    Dim preClm      As Long
    preClm = 0
    Call stopUpdate

    Dim baseClm     As Long
    Dim shtPage     As Long

    '~~~~~~~�ǉ�~~~~~~~~~~~~~~~~~~~
    '�J���[��ێ�����ϐ�
    Dim colorStr    As String
    Dim colorStr1   As String '����A
    Dim colorStr2   As String '����B
    Dim colorStr3   As String '����C

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '���掞��(�b)�ɂ���̏�������ύX����

    Dim wSize       As widthSize

    '---------------------------------------------
    '�ϐ��A�z��ɒl�����
    '---------------------------------------------
    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
        '�ŏI�s���擾
        maxRowNum = .Cells(1, 3).End(xlDown).row
        '�z��̍Ō��
'        �]�����폜
        maxRowNum = maxRowNum - 1
        '�z����Ē�`
        ReDim postureScoreDataArray_A(maxRowNum, 0)
        ReDim postureScoreDataArray_B(maxRowNum, 0)
        ReDim postureScoreDataArray_C(maxRowNum, 0)

        '�M������ԗp
        ReDim reliabilityDataArray(maxRowNum, 0)

        '�z��̒��ɒl������
        For i = 1 To maxRowNum

            '����A����̔z�������
            postureScoreDataArray_A(i - 1, 0) = .Cells(i + 1, COLUMN_POSTURE_SCORE_KOBUSHIAGE - 1).Value
            postureScoreDataArray_B(i - 1, 0) = .Cells(i + 1, COLUMN_POSTURE_SCORE_KOSHIMAGE - 1).Value
            postureScoreDataArray_C(i - 1, 0) = .Cells(i + 1, COLUMN_POSTURE_SCORE_HIZAMAGE - 1).Value

            '�M������z��ɓ����
            '1:����A2:����A3:����

            If .Cells(i + 1, COLUMN_MEASURE_SECTION).Value > 0 Then
                reliabilityDataArray(i, 0) = 1
            End If
            If .Cells(i + 1, COLUMN_PREDICT_SECTION).Value > 0 Then
                reliabilityDataArray(i, 0) = 2
            End If
            If .Cells(i + 1, COLUMN_MISSING_SECTION).Value > 0 Then
                reliabilityDataArray(i, 0) = 3
            End If
        Next
        '�t���[�����[�g���擾
        fps = getFps()
        Dim video_sec As Double: video_sec = wholeEnd / fps

    End With

    '---------------------------------------------
    '�����͈͂����߂�
    '---------------------------------------------
    '�L�����Z��(�߂�)�{�^������Ă΂ꂽ�Ƃ�

    If processingRange = 1 Then
        '�A�N�e�B�u�Z���̈�ԍ���6��ڈȉ��̎�
        '�G���[���b�Z�[�W���o���ď�������߂�

        shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
        baseClm = LIMIT_COLUMN * shtPage

        'pageLimit�����̃y�[�W�ƂȂ�臒l�܂ōX�V
        thisPageLimit = (shtPage + 1) * LIMIT_COLUMN
        preClm = (LIMIT_COLUMN * shtPage) * -1

        Dim lCol As Long, rCol As Long
        If Not CropSelectionToDataArea(lCol, rCol) Then
            MsgBox "�͈͊O�ł�", vbCritical
            Exit Sub
        End If

        wholeStart = lCol - COLUMN_ZERO_NUM + baseClm
        wholeEnd = rCol - COLUMN_ZERO_NUM + baseClm

        If wholeStart < 1 Then
            wholeStart = 1
        End If

    '���C���̏�������Ă΂ꂽ�Ƃ�
    ElseIf processingRange = 2 Then

        '�擪����
        wholeStart = 1
        '�����܂�
        wholeEnd = maxRowNum

        '���workSheet�A���킹�ď�����
        ThisWorkbook.Sheets("�p���]���C���V�[�g").Activate
        preClm = 0
        If maxRowNum >= 150 Then
            Call autoFillLine(ActiveSheet, wholeEnd + COLUMN_ZERO_NUM) '230206 + COLUMN_ZERO_NUM��ǉ�
            Call autoFillTime(Worksheets("�p���]���C���V�[�g"), 0, wholeEnd)
        End If

        Call addPageShape(ActiveSheet, False, True)

        '15�b�ȉ����2�Ƃ���
        If video_sec <= 15 Then
            wSize = LL
            Call changeBtnState(EXPANDBTN_NAME, False)
            Call changeBtnState(REDUCEBTN_NAME, True)
        Else
            wSize = Small
            Call changeBtnState(REDUCEBTN_NAME, False)
            Call changeBtnState(EXPANDBTN_NAME, True)
        End If

        Call DataAjsSht.SetCellsHW(CInt(wSize), Worksheets("�p���]���C���V�[�g"))

    '���O������t���[���ɋ������㏑�������Ƃ��i�P�Z�������s�j
    Else
        shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
        baseClm = LIMIT_COLUMN * shtPage

        'pageLimit�����̃y�[�W�ƂȂ�臒l�܂ōX�V
        thisPageLimit = (shtPage + 1) * LIMIT_COLUMN
        preClm = (LIMIT_COLUMN * shtPage) * -1

        wholeStart = processingRange - COLUMN_ZERO_NUM + baseClm

        wholeEnd = wholeStart
    End If

    For i = wholeStart To wholeEnd

        '�p���_�̃J�E���^�[�����Z�b�g
        'A����C�̎p���_�̃J�E���^�[�����Z�b�g
        Erase postureScoreCounterArray_A
        Erase postureScoreCounterArray_B
        Erase postureScoreCounterArray_C
'
        '�M�����̃J�E���^�[�����Z�b�g
        Erase reliabilityCounterArray

        '�p���_���m�F
        'A����C�̎p���_���m�F
        postureScoreFlag_A = postureScoreDataArray_A(i - 1, 0)
        postureScoreFlag_B = postureScoreDataArray_B(i - 1, 0)
        postureScoreFlag_C = postureScoreDataArray_C(i - 1, 0)

        '�p���_�t���O�𗧂Ă�
        'A����C�̎p���_�t���O�𗧂Ă�
        postureScoreCounterArray_A(postureScoreFlag_A) = 1
        postureScoreCounterArray_B(postureScoreFlag_B) = 1
        postureScoreCounterArray_C(postureScoreFlag_C) = 1

        '�M�������m�F
        reliabilityFlag = reliabilityDataArray(i, 0)
        '�M�����t���O�𗧂Ă�
        reliabilityCounterArray(reliabilityFlag) = 1

        '---------------------------------------------
        '�t���[�������ł��������̂�T��
        '---------------------------------------------
        mostOftenPostureScore = 0
        mostOftenPostureScore_A = 0
        mostOftenPostureScore_B = 0
        mostOftenPostureScore_C = 0

        '�p���_ 0 ~ 1 �̐擪���珇�ɔ�r
        For j = 0 To 1
            '�t���[�����̍��v�������p���_��I��
            '���v�������ꍇ�͐h���p����D�悷��

            '����A����C
            If postureScoreCounterArray_A(mostOftenPostureScore_A) <= postureScoreCounterArray_A(j) Then
                mostOftenPostureScore_A = j
            End If

            If postureScoreCounterArray_B(mostOftenPostureScore_B) <= postureScoreCounterArray_B(j) Then
                mostOftenPostureScore_B = j
            End If

            If postureScoreCounterArray_C(mostOftenPostureScore_C) <= postureScoreCounterArray_C(j) Then
                mostOftenPostureScore_C = j
            End If

        Next

        '������1
        mostOftenReliability = 1
            '�M����1�`3�̐擪���珇�ɔ�r
            '1:����A2:����A3:����
        For j = 2 To 3
            '�t���[�����̍��v�������p���_��I��
            '���v�������ꍇ�͐M�������Ⴂ����D�悷��
            If reliabilityCounterArray(mostOftenReliability) <= reliabilityCounterArray(j) Then
                mostOftenReliability = j
            End If
        Next

        'active sheet��ύX����
        If i <= thisPageLimit Then
            '�������Ȃ�
        Else
            ThisWorkbook.ActiveSheet.Next.Activate
            If InStr(ThisWorkbook.ActiveSheet.Name, "�p���]���C���V�[�g") > 0 Then
                '�������Ȃ�
            Else
                '�߂�
                ThisWorkbook.ActiveSheet.Previous.Activate
                Call createSheet(wholeEnd - i)
            End If
            '�X�V
            thisPageLimit = thisPageLimit + LIMIT_COLUMN
            preClm = preClm - LIMIT_COLUMN
            Call clear(ActiveSheet)
            Call autoFillLine(ActiveSheet, wholeEnd - i)
            Call autoFillTime(ThisWorkbook.ActiveSheet, ((thisPageLimit / LIMIT_COLUMN) - 1) * 9, wholeEnd - i)
            Call addPageShape(ActiveSheet, True, True)
        End If
        '---------------------------------------------
        '�p���]���C���V�[�g�̃Z���ɐF��h��
        '---------------------------------------------
        With ThisWorkbook.ActiveSheet

            '-------------����A
            '0�_�̏ꍇ�A��
            If mostOftenPostureScore_A = 0 Then
                colorStr1 = colorResultWhite

            '1�_�̏ꍇ�A��
            ElseIf mostOftenPostureScore_A = 1 Then
                colorStr1 = colorResultRed
            End If

'            -------------����B
            '0�_�̏ꍇ�A��
            If mostOftenPostureScore_B = 0 Then
                colorStr2 = colorResultWhite

            '1�_�̏ꍇ�A��
            ElseIf mostOftenPostureScore_B = 1 Then
                colorStr2 = colorResultRed
            End If

'            -------------����C
            '0�_�̏ꍇ�A��
            If mostOftenPostureScore_C = 0 Then
                colorStr3 = colorResultWhite

            '1�_�̏ꍇ�A��
            ElseIf mostOftenPostureScore_C = 1 Then
                colorStr3 = colorResultRed
            End If

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            '�F���N���A
            .Range _
            ( _
                .Cells(ROW_POSTURE_SCORE_KOSHIMAGEOTTOM, COLUMN_ZERO_NUM + i + preClm), _
                .Cells(ROW_POSTURE_SCORE_TOP, COLUMN_ZERO_NUM + i + preClm) _
            ) _
            .Interior.ColorIndex = 0

            '~~~~~~~~~~~~~~~�ǉ�~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '�F��h��
            '����A
            If mostOftenPostureScore_A = 0 Or 1 Then
                .Range _
                ( _
                    .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, COLUMN_ZERO_NUM + i + preClm), _
                    .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, COLUMN_ZERO_NUM + i + preClm) _
                ) _
                .Interior.Color = colorStr1
            End If

            '����B
            If mostOftenPostureScore_B = 0 Or 1 Then

                .Range _
                ( _
                    .Cells(ROW_POSTURE_SCORE_KOSHIMAGE, COLUMN_ZERO_NUM + i + preClm), _
                    .Cells(ROW_POSTURE_SCORE_KOSHIMAGE, COLUMN_ZERO_NUM + i + preClm) _
                ) _
                .Interior.Color = colorStr2
            End If

            '����C
            If mostOftenPostureScore_C = 0 Or 1 Then

                .Range _
                ( _
                    .Cells(ROW_POSTURE_SCORE_HIZAMAGE, COLUMN_ZERO_NUM + i + preClm), _
                    .Cells(ROW_POSTURE_SCORE_HIZAMAGE, COLUMN_ZERO_NUM + i + preClm) _
                ) _
                .Interior.Color = colorStr3
            End If

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            '�ł����������M�����ɉ�����
            '�F��ύX
            '1:����A2:����A3:����
            If mostOftenReliability = 1 Then
                colorStr = colorMeasureSection
            ElseIf mostOftenReliability = 2 Then
                colorStr = colorPredictSection
            ElseIf mostOftenReliability = 3 Then
                colorStr = colorMissingSection
            End If

            .Range _
            ( _
                .Cells(ROW_RELIABILITY_TOP, COLUMN_ZERO_NUM + i + preClm), _
                .Cells(ROW_RELIABILITY_BOTTOM, COLUMN_ZERO_NUM + i + preClm) _
            ) _
            .Interior.Color = colorStr

        End With
    Next

    ' �L�����Z���{�^���ȊO����̏����̎�
    If 1 < processingRange Then
        If calcSheetNamePlace(ThisWorkbook.ActiveSheet) = 0 Then ' 0 = Base sheet
            Call addPageShape(ActiveSheet, False, False)
        Else
            Call addPageShape(ActiveSheet, True, False)
        End If
    End If

    '�e�V�[�g���X�V
    Call checkReliabilityRatio
    Call restartUpdate

End Sub


'�w�S�̂������x�{�^���������ꂽ�Ƃ�
'�S�̂̎p���_���v�Z���āA�F��h��
Sub paintAll()
    Call paintPostureScore(2)
End Sub


'�wCancel�x�{�^���������ꂽ�Ƃ�
'�I��͈͂̎p���_���v�Z���āA�F��h��i�����{�^���̃L�����Z���j
Sub paintSelected()
    '����1:processingRange As Long �����͈͂����߂�

    '��ŏC�����邱�ƂɂȂ邪�A7��ڂ�菬�����񂪑I������Ă����珈�������Ȃ�
    If DataAjsSht.activeCells <= COLUMN_ZERO_NUM Then
        Exit Sub
    End If

    Call paintPostureScore(1)
End Sub


'�h��Ԃ���S�ăN���A
Sub clear(ws As Worksheet)
    ws _
    .Range _
    ( _
        Cells(ROW_RELIABILITY_TOP, COLUMN_ZERO_NUM + 1), _
        Cells( _
            ROW_POSTURE_SCORE_KOSHIMAGEOTTOM, _
            Cells(ROW_POSTURE_SCORE_KOSHIMAGEOTTOM, COLUMN_ZERO_NUM + 1).End(xlToRight).Column _
        ) _
    ) _
    .Interior.ColorIndex = 0
End Sub


'���ʂ̏C���{�^��
'�p���_�������I�ɕύX����
'�{�^���ʂň���postureScorebutton���ς��
Sub forceResult(postureScorebutton As Long)
    '---------------------------------------------
    'RGB���w�肷�邽�߂̕ϐ����`
    '---------------------------------------------
    '�M����
    Dim colorMeasureSection    As String '���F
    Dim colorPredictSection    As String '���F
    Dim colorMissingSection    As String '�s���N
    Dim colorForcedSection     As String '�F
    Dim colorRemoveSection     As String '�O���[

    '�p���_
    Dim colorResultGreen       As String '�ΐF
    Dim colorResultYellow      As String '���F
    Dim colorResultRed         As String '�ԐF
    Dim colorResultGlay        As String '�O���[
    Dim colorResultWhite       As String '���F 20221219_����
    Dim colorResultBrown       As String '���F 20221222_����
    Dim colorResultOFFGlay     As String '�O���[ 20221222_����

    '---------------------------------------------
    '�ϐ��ɐF���Z�b�g
    '---------------------------------------------
    '1:����A2:����A3:�����A4:�����A5:���O
    '�M����
    colorMeasureSection = RGB(0, 176, 240)   '���F
    colorPredictSection = RGB(252, 246, 0)   '���F
    colorMissingSection = RGB(255, 124, 128) '�s���N
    colorForcedSection  = RGB(0, 51, 204)    '�F
    colorRemoveSection  = RGB(191, 191, 191) '�O���[
    '�p���_
    colorResultGreen    = RGB(0, 176, 80)    '�ΐF
    colorResultYellow   = RGB(255, 192, 0)   '���F
    colorResultRed      = RGB(192, 0, 0)     '�ԐF
    colorResultGlay     = RGB(191, 191, 191) '�O���[
    colorResultWhite    = RGB(255, 255, 255) '���F
    colorResultBrown    = RGB(64, 0, 0)      '���F
    colorResultOFFGlay  = RGB(217, 217, 217) '����I�t�p�̃O���[

    Dim baseClm As Long
    Dim shtPage As Long
    shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
    baseClm = LIMIT_COLUMN * shtPage

    '�I��͈͓��̃Z���ǂݍ��ݗp�@20221222_����
    Dim SelectCells  As Variant
    Dim MaxRightCell As Variant
    Dim MinLeftCell  As Variant

    Dim lCol As Long
    Dim rCol As Long

    '�ꎞ�I��Selection.row�̉���ۑ����Ă����ϐ�
    Dim postur_row As Long

    '�ϐ���`
    Dim k As Long
    Dim m As Long
    '---------------------------------------------
    '�������狭������
    '---------------------------------------------
    With ThisWorkbook.ActiveSheet
        '�C���V�[�g�̑I��͈͂̓|�C���g�v�Z�V�[�g����͂ݏo���Ȃ��͈͂ɂ��邱��
        '�C���V�[�g�̑I��͈͂͐F�h��ł���͈͂ɂ��邱��
        If CropSelectionToDataArea(lCol, rCol) Then

            '�I��͈͂̍��[�ƉE�[���擾
            MinLeftCell = lCol
            MaxRightCell = rCol

            '�߂�(Remove�{�^��)
            If postureScorebutton = -1 Then
                Call postureUpdate(MinLeftCell + baseClm, MaxRightCell + baseClm, 0, CInt(postureScorebutton))
                '�������܂Ŗ߂�{�^�����������Ƃ��ɃL�b�N�����}�N��
                Call paintPostureScore(1)

            '����(0�`�P�P�̎p���_�{�^��)
            ElseIf postureScorebutton >= 0 Then

                Call postureUpdate(MinLeftCell + baseClm, MaxRightCell + baseClm, 1, CInt(postureScorebutton))

                If postureScorebutton = 99 Then
                    '���O��99�ɕύX�@20221219_����
                    '�ŏ��ɔw�i�h��Ԃ������ɂ��Ă���̂ŁA���������Ȃ�
                    '�M�����̃Z���ɏ��O�̐F��h��
                    .Range _
                    ( _
                        .Cells(ROW_RELIABILITY_TOP, MinLeftCell), _
                        .Cells(ROW_RELIABILITY_BOTTOM, MaxRightCell) _
                    ) _
                    .Interior.Color = colorRemoveSection

                    For k = 1 To 3
                        .Range _
                        ( _
                            .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE - 2 + 2 * k, MinLeftCell), _
                            .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE - 2 + 2 * k, MaxRightCell) _
                        ) _
                        .Interior.Color = colorResultGlay
                    Next

                Else

                    postur_row = Selection.row

                    '===�������A�ق��̗�̏��O���������A���O�������ꏊ�Ɍ��̃f�[�^�ǂ���ɐF��t���Ȃ�������===

                    For m = MinLeftCell To MaxRightCell
                        If .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, m).Interior.Color = colorResultGlay Then
                            Call paintPostureScore(m)
                        End If
                    Next
                    '===========================================================================================

                    '�p���_�̃Z���ɉ����ꂽ�{�^���̎p���_
                    '1�_�̏ꍇ�͐�
                    If postureScorebutton = 1 Then
                        .Range _
                        ( _
                            .Cells(postur_row, MinLeftCell), _
                            .Cells(postur_row, MaxRightCell) _
                        ) _
                        .Interior.Color = colorResultBrown

                    '0�_�̏ꍇ�͔�
                    ElseIf postureScorebutton = 0 Then
                        .Range _
                        ( _
                            .Cells(postur_row, MinLeftCell), _
                            .Cells(postur_row, MaxRightCell) _
                        ) _
                        .Interior.Color = colorResultOFFGlay
                    End If

                    '�M�����̃Z���ɋ����F���ʂ�
                    .Range _
                    ( _
                        .Cells(ROW_RELIABILITY_TOP, MinLeftCell), _
                        .Cells(ROW_RELIABILITY_BOTTOM, MaxRightCell) _
                    ) _
                    .Interior.Color = colorForcedSection

                End If
                '�����̂Ƃ��͒P�ƂŎ��s
                Call checkReliabilityRatio
            End If
        Else
            MsgBox "�͈͂̓O���t������I�����Ă�������", vbOKOnly + vbCritical, "�͈͑I���G���["
        End If
    End With

    Call checkReliabilityRatio

End Sub


'�_���ɉ����Đ��l���w��̗�ɒl��}������
'����1�F�I��͈͂̍��[�̃Z��
'����2�F�I��͈͂̉E�[�̃Z��
'����3�F�߂邩��Ă΂ꂽ��0�A����ȊO��1
'����4�F�ǂ̃{�^������Ă΂ꂽ������ʂ���ID
'       �߂�   �F-1
'       ����ON �F1
'       ����OFF�F0
'       ���O   �F99
Sub postureUpdate(sclm As Long, fclm As Long, bit As Long, score As Long)

    Dim s    As Long
    Dim last As Long
    Dim i    As Long
    Dim fbit As Long
    Dim vle  As Long

    Dim column_forced_num As Long

    If Selection.row = ROW_POSTURE_SCORE_KOBUSHIAGE Then
        column_forced_num = COLUMN_POSTURE_SCORE_KOBUSHIAGE
    ElseIf Selection.row = ROW_POSTURE_SCORE_KOSHIMAGE Then
        column_forced_num = COLUMN_POSTURE_SCORE_KOSHIMAGE
    ElseIf Selection.row = ROW_POSTURE_SCORE_HIZAMAGE Then
        column_forced_num = COLUMN_POSTURE_SCORE_HIZAMAGE
    End If

    '�|�C���g�v�Z�V�[�g�ł�1�s�ڂ���l�𐔂��Ȃ���2�s�ڂ���ƂȂ邽��+1
    s = sclm - COLUMN_ZERO_NUM + 1
    last = fclm - COLUMN_ZERO_NUM + 1

    For i = s To last

        With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
            fbit = .Cells(i, COLUMN_FORCED_SECTION_TOTAL).Value

            If bit = 0 Then
                If fbit = 0 Then
                    vle = .Cells(i, COLUMN_POSTURE_SCORE_ALL).Value
                Else
                    vle = .Cells(i, COLUMN_BASE_SCORE).Value
                End If

                '�p���f�_���O��ԂɃr�b�g�������Ă���
                If .Cells(i, COLUMN_REMOVE_SECTION).Value = 1 Then
                    vle = .Cells(i, COLUMN_BASE_SCORE).Value
                End If
            Else
                vle = score
            End If

            Call baseScore(i, bit)
            .Cells(i, COLUMN_POSTURE_SCORE_ALL).Value = vle
        End With

        Call reliabilityUpdate(i, bit, vle, column_forced_num)
    Next

End Sub

'�p���f�_������ԂɃr�b�g�𗧂Ă鏈��
'����1�F�|�C���g�v�Z�V�[�g�̏C������Z���̍s
'����2�F���Z�b�g����Ă΂ꂽ��0�A����ȊO��1
'����3�F�ǂ̃{�^������Ă΂ꂽ������ʂ���ID
'       ���Z�b�g �F-1
'       ����ON �@�F1
'       ����OFF�@�F0
'       ���O   �@�F99
'����4�F�|�C���g�v�Z�V�[�g�̏C������Z���̗�
Sub reliabilityUpdate(row As Long, bit As Long, vle As Long, column_forced_num As Long)
    '�ϐ���`
    Dim column_reliability_forced_num As Long


    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
        '���O
        If vle = 99 And bit = 1 Then
            '�p���f�_���O���
            .Cells(row, COLUMN_REMOVE_SECTION).Value = bit

            '�p���̋���������
            .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = 0
            .Cells(row, COLUMN_POSTURE_SCORE_KOSHIMAGE).Value = 0
            .Cells(row, COLUMN_POSTURE_SCORE_HIZAMAGE).Value = 0


            '�M�����̋���������
            .Cells(row, COLUMN_FORCED_SECTION_KOBUSHIAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_KOSHIMAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_HIZAMAGE).Value = 0

        '���Z�b�g
        ElseIf bit = 0 Then
            '�p���f�_������ԂƎp���f�_���O��Ԃ̃r�b�g������
            .Cells(row, COLUMN_FORCED_SECTION_TOTAL).Value = bit
            .Cells(row, COLUMN_REMOVE_SECTION).Value = bit

            '�p�������Z�b�g
            .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE - 1).Value
            .Cells(row, COLUMN_POSTURE_SCORE_KOSHIMAGE).Value = .Cells(row, COLUMN_POSTURE_SCORE_KOSHIMAGE - 1).Value
            .Cells(row, COLUMN_POSTURE_SCORE_HIZAMAGE).Value = .Cells(row, COLUMN_POSTURE_SCORE_HIZAMAGE - 1).Value


            '�M�����̋���������
            .Cells(row, COLUMN_FORCED_SECTION_KOBUSHIAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_KOSHIMAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_HIZAMAGE).Value = 0


        '����
        Else
            '�M�����������ɂ��������߂�
            If column_forced_num = COLUMN_POSTURE_SCORE_KOBUSHIAGE Then
                column_reliability_forced_num = COLUMN_FORCED_SECTION_KOBUSHIAGE
            ElseIf column_forced_num = COLUMN_POSTURE_SCORE_KOSHIMAGE Then
                column_reliability_forced_num = COLUMN_FORCED_SECTION_KOSHIMAGE
            ElseIf column_forced_num = COLUMN_POSTURE_SCORE_HIZAMAGE Then
                column_reliability_forced_num = COLUMN_FORCED_SECTION_HIZAMAGE
            End If

            '���O������
            .Cells(row, COLUMN_REMOVE_SECTION).Value = 0
            '���㍘�Ȃ��G�Ȃ��̂����ꂩ������
            .Cells(row, column_forced_num).Value = vle
            '�S�̂̐M��������
            .Cells(row, COLUMN_FORCED_SECTION_TOTAL).Value = bit
            '���㍘�Ȃ��G�Ȃ��̂����ꂩ�M�����������ɂ���
            .Cells(row, column_reliability_forced_num).Value = 1
        End If
    End With

End Sub


'���f�[�^��֑}������
'����1�F�f�[�^��}������Z���̍s
'����2�F�߂邩��Ă΂ꂽ��0�A����ȊO��1
Sub baseScore(row As Long, bit As Long)
    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
        If bit = 1 Then
            If .Cells(row, COLUMN_BASE_SCORE).Value = "" Then
                .Cells(row, COLUMN_BASE_SCORE).Value = .Cells(row, COLUMN_POSTURE_SCORE_ALL).Value
            End If
        Else
            .Cells(row, COLUMN_POSTURE_SCORE_ALL).Value = .Cells(row, COLUMN_BASE_SCORE).Value
        End If
    End With
End Sub


'�w�������x�{�^���������ꂽ�Ƃ�
Sub reset()
    Call forceResult(-1)
End Sub


'�p���_�w0�x�����{�^���������ꂽ�Ƃ�
Sub force0()
    Call forceResult(0)
End Sub


'�p���_�w1�x�����{�^���������ꂽ�Ƃ�
Sub force1()
    Call forceResult(1)
End Sub


'�p���_�w���O�x�����{�^���������ꂽ�Ƃ�
Sub jogai()
    Call forceResult(99)
End Sub


'�f�[�^��Ԃ̊������v�Z
Sub checkReliabilityRatio()
    '�ϐ���`
    Dim i                               As Long
    '�t���[�����[�g
    Dim fps                             As Double
    '�|�C���g�v�Z�V�[�g�ŏI�s
    Dim maxRowNum                       As Long
    '�p���]���C���V�[�g�̍ŏI��
    Dim ColumnNum                       As Long
    '�z��̍ŏI�l
    Dim maxArrayNum                     As Long
    '�M�����̔ԍ�
    '1:����A2:����A3:�����A4:�����A5:���O
    Dim reliabilityFlag                 As Long
    '�M�����̊���
    Dim measurementSectionRatio         As Double
    Dim predictSectionRatio             As Double
    Dim missingSectionRatio             As Double
    Dim coercionSectionRatio            As Double
    Dim exclusionSectionRatio           As Double
    Dim totalRatio                      As Double

    '�z���`
    '�F��ۑ�����z��
    Dim reliabilityColorDataArray()     As Long
    '�F���J�E���g����z��
    '�M�����P�`�R�̃t���[���������ꂼ�ꍇ�v
    '1:����A2:����A3:�����A4:�����A5:���O
    Dim reliabilityColorCounterArray(5) As Long

    '---------------------------------------------
    'RGB���w�肷�邽�߂̕ϐ����`
    '---------------------------------------------
    '�M����
    Dim colorMeasureSection    As String '���F
    Dim colorPredictSection    As String '���F
    Dim colorMissingSection    As String '�s���N
    Dim colorForcedSection     As String '�F
    Dim colorRemoveSection     As String '�O���[

    '---------------------------------------------
    '�ϐ��ɐF���Z�b�g
    '---------------------------------------------
    '�M����
    colorMeasureSection = RGB(0, 176, 240)   '���F
    colorPredictSection = RGB(252, 246, 0)   '���F
    colorMissingSection = RGB(255, 124, 128) '�s���N
    colorForcedSection = RGB(0, 51, 204)     '�F
    colorRemoveSection = RGB(191, 191, 191)  '�O���[

    '---------------------------------------------
    '�ϐ��E�z�񏀔�
    '---------------------------------------------
    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
        '�t���[�����[�g���擾
        fps = getFps()
        '�ŏI�s���擾
        maxRowNum = getLastRow()
    End With

    '�p���]���C���V�[�g
    Dim sName()  As String
    Dim n        As Long
    Dim actSheet As Worksheet

    '�]��������
    maxRowNum = maxRowNum - 1

    '��ԉE�̗񐔂��擾
    With ThisWorkbook.Sheets("�p���]���C���V�[�g")
        ColumnNum = Columns.Count - 6
    End With
    '�ŏ��̗�(6��܂�)����ǉ�����
    ColumnNum = 16206

    maxArrayNum = ColumnNum - 1

    '�z����Ē�`
    ReDim reliabilityColorDataArray(maxArrayNum, 0)

    '�J�E���^�[��������
    Erase reliabilityColorCounterArray

    '---------------------------------------------
    '��������M�����̊������v�Z
    '---------------------------------------------

    For i = 2 To maxRowNum + 1 '230208

        With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

            '���O
            If .Cells(i, COLUMN_REMOVE_SECTION).Value > 0 Then
                reliabilityColorCounterArray(5) = reliabilityColorCounterArray(5) + 1
                GoTo CONTINUE:
            '����
            ElseIf .Cells(i, COLUMN_FORCED_SECTION_TOTAL).Value > 0 Then
                reliabilityColorCounterArray(4) = reliabilityColorCounterArray(4) + 1
                GoTo CONTINUE:
            '����
            ElseIf .Cells(i, COLUMN_MISSING_SECTION).Value > 0 Then
                reliabilityColorCounterArray(3) = reliabilityColorCounterArray(3) + 1
                GoTo CONTINUE:
            '����
            ElseIf .Cells(i, COLUMN_PREDICT_SECTION).Value > 0 Then
                reliabilityColorCounterArray(2) = reliabilityColorCounterArray(2) + 1
                GoTo CONTINUE:
            '����
            ElseIf .Cells(i, COLUMN_MEASURE_SECTION).Value > 0 Then
                reliabilityColorCounterArray(1) = reliabilityColorCounterArray(1) + 1
                GoTo CONTINUE:

            End If
        End With

CONTINUE:
    Next

    '�������v�Z
    '����
    predictSectionRatio = reliabilityColorCounterArray(2) / maxRowNum * 100
    '����
    missingSectionRatio = reliabilityColorCounterArray(3) / maxRowNum * 100
    '���O
    exclusionSectionRatio = reliabilityColorCounterArray(5) / maxRowNum * 100
    '����
    measurementSectionRatio = reliabilityColorCounterArray(1) / maxRowNum * 100
    '����
    coercionSectionRatio = reliabilityColorCounterArray(4) / maxRowNum * 100


    Set actSheet = ActiveSheet
    sName() = call_GetSheetNameToArrayspecific(ThisWorkbook, "�p���]���C���V�[�g")
    For n = 1 To UBound(sName)
        '�������Z���ɓ���
        With ThisWorkbook.Sheets(sName(n))
            '����
            .Cells(3, 4) = Round(measurementSectionRatio, 1) & "%"
            '����
            .Cells(4, 4) = Round(coercionSectionRatio, 1) & "%"
            '���O
            .Cells(5, 4) = Round(exclusionSectionRatio, 1) & "%"
            '����
            .Cells(6, 4) = Round(predictSectionRatio, 1) & "%"
            '����
            .Cells(7, 4) = Round(missingSectionRatio, 1) & "%"
            '����+����+���O
            .Cells(3, 5) = Round(measurementSectionRatio + coercionSectionRatio + exclusionSectionRatio, 1) & "%"
            '����+����
            .Cells(6, 5) = Round(predictSectionRatio + missingSectionRatio, 1) & "%"

        End With
    Next
End Sub


'�g��{�^���A�k���{�^���������ꂽ�Ƃ��Ɏ��s����鏈��
    '�����FexpansionFlag As Long�@���̊g��or�k�������߂�
    'False�F�k���@True:�g��

    '�R�[�h�ِ̑̍����Ă������ꂽ�̂ŏ�����Ԃɖ߂��Ă��܂�230213
Sub adjustWidth(expansionFlag As Boolean)
    Dim columnWidth0 As Double
    Const EXPANSION_RATIO As Long = 100
    Static initFin As Boolean
    Static wSize As widthSize

    Call stopUpdate
    '�g��E�k���ǂ���̃t���O���m�F�i�{�^����������󂯎��j
    '�k���{�^��

    '���߂ČĂ΂ꂽ����������
    If (initFin = False) Then
        initFin = initFin + True
        Dim initSize As Long
        initSize = DataAjsSht.GetWidthPoints
        Select Case initSize
            Case Is < widthSize.Medium
                wSize = Small
            Case Is < widthSize.Large
                wSize = Medium
            Case Is < widthSize.LL
                wSize = Large
            Case Else
                wSize = LL
        End Select
    End If

    wSize = sizeNext(wSize, expansionFlag)

    Dim sName() As String
    Dim n As Long
    Dim actSheet As Worksheet
    Set actSheet = ActiveSheet
    sName() = call_GetSheetNameToArrayspecific(ThisWorkbook, "�p���]���C���V�[�g")
    For n = 1 To UBound(sName)
        Call DataAjsSht.SetCellsHW(CInt(wSize), ThisWorkbook.Sheets(sName(n)))
    Next
    actSheet.Activate
    Call restartUpdate
End Sub


'�w���g��x�{�^���������ꂽ�Ƃ�
Sub expandWidth()
    '�����FexpansionFlag As Long�@���̊g��or�k�������߂�
    'False�F�k���@True:�g��
    Call adjustWidth(True)
End Sub


'�w���k���x�{�^���������ꂽ�Ƃ�
Sub reduceWidth()
    '�����FexpansionFlag As Boolean�@���̊g��or�k�������߂�
    'False�F�k���@True:�g��
    Call adjustWidth(False)
End Sub


'1��ʍ��փX�N���[��
Sub scrollToLeftPage()
        ActiveWindow.LargeScroll ToLeft:=1
End Sub


'1��ʉE�փX�N���[��
Sub scrollToRightPage()
        If ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count <= _
        ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column Then
            ActiveWindow.LargeScroll ToRight:=1
        End If
End Sub


'�ł����փX�N���[��
Sub scrollToLeftEnd()
    Dim scrclm As Long
    If getClm(1) Then
        If getPageShapeState(ActiveSheet, "prevPage") Then
            Call prevPage_Click
        Else
            Call initCellPlace(ActiveSheet)
        End If
    Else
        Call initCellPlace(ActiveSheet)
    End If

End Sub


'�ł��E�փX�N���[��
Sub scrollToRightEnd()
    '�����������̈�ԉE�̗���o���Ă����B
    '�����l�ŉE�ɍs���Ƃ��́A���V�[�g������΂�����ֈڍs����B
    Dim keepColumn As Long

    If getClm(ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column) Then
        If getPageShapeState(ActiveSheet, "nextPage") Then
            Call nextPage_Click
        End If
    Else
        keepColumn = keepColumn * 0 + ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column

        ActiveWindow.SmallScroll ToLeft:=ActiveWindow.Panes(2).VisibleRange.Cells.Columns.Count

        '�ȉ��̕���͍���͂���Ȃ��\��������
        '�����E��
        If ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column = 16192 Then
            ActiveWindow.SmallScroll ToRight:=5
        Else
            '3�b�����炷(=30fps * 3)
            ActiveWindow.SmallScroll ToRight:=90
        End If

        Call finCellPlace(ActiveSheet)

    End If
End Sub


'���݂̃J������ێ�����
Private Function getClm(clm As Long)
    Static keepColumn As Long
    Dim ret As Boolean: ret = False

    If keepColumn = clm Then
        ret = True
    Else
        keepColumn = keepColumn * 0 + clm
    End If
    getClm = ret
End Function


'�\���{������ʂɃt�B�b�g
Sub fit()
    '�����Ă����͈͂��擾
    Dim visibleColumn As String

    '�����Ă����͈͂̂���������7�Ԗڂ̗���擾�i�ҏW�{�^�����u����Ă���1�`6����΂��j
    visibleColumn = Split(ActiveWindow.VisibleRange.Cells(7, 1).Address(True, False), "$")(0)
    '1�`�����̂P�s���܂ł�I��
    Range(visibleColumn & "1:" & visibleColumn & BOTTOM_OF_TABLE + 1).Select
    '��ʂɃt�B�b�g
    ActiveWindow.Zoom = True
    'A1�Z����I������
    Range("A1").Select
    '��ʂ���ԏ�܂ŃX�N���[��
    ActiveWindow.ScrollRow = 1

End Sub


'�Đ��{�^��
'�����I���̃C���^�[�o��
Sub RegularInterval3()
    Dim iend, i As Long
    Dim dajsht() As String
    Dim l As Long

    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "�p���]���C���V�[�g")
    iend = UBound(dajsht)
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes("playBtn").Visible = False
        End With
    Next

    l = ActiveCell.Column
    If l < TIME_COLUMN_LEFT Then
        ActiveSheet.Cells(BOTTOM_OF_TABLE, TIME_COLUMN_LEFT).Select
        '2�b����n�܂�悤�Ɍ����邽��1�b�ҋ@����
        Application.Wait Now() + TimeValue("00:00:01")
    End If

    'activesheet�ŃR�s�[��ɂ��Ή�����

    '�ϐ�ResTime�Ɍ��݂�1�b��̎������i�[
    ResTime = Now + TimeValue("00:00:01")

    'Application�I�u�W�F�N�g��OnTime���\�b�h���g�p
    'EarliestTime : ���s����(����������1�b��j
    'Procedure : ���s�v���V�[�W�����B�������g���w�肵�ČJ��Ԃ�����
    Application.OnTime EarliestTime:=ResTime, _
    Procedure:="RegularInterval3"

    '�uTestSample1�v�v���V�[�W���̌Ăяo��
    Call nextTimeSelect

End Sub


'�����I������
Sub nextTimeSelect()

    '�A�N�e�B�u�Z���̈�ԍ��̗񐔂��擾
    '�擾�����񐔂̎���(23�s�ځj���A�N�e�B�u�ɂ���
    Cells(TIME_ROW, Selection.Column).Select

    '��E�̃Z����I��
    ActiveCell.Offset(0, 1).Select

    '1�b���X�N���[��
    ActiveWindow.SmallScroll ToRight:=TIME_WIDTH

    '�������\������Ă��Ȃ����́A������ύX����K�v������
    If IsEmpty(ActiveCell.Value) Then
        'arrow�������Ă��鎞�A���Ȃ킿���̃V�[�g�����݂���B
        If getPageShapeState(ActiveSheet, "nextPage") Then
            Call nextPage_Click
        Else '�I�[
            Call Cancel3
        End If
    End If

End Sub


'��~�{�^��
Sub Cancel3()
    Dim iend, i As Long
    Dim dajsht() As String

    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "�p���]���C���V�[�g")
    iend = UBound(dajsht)

    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes("playBtn").Visible = True
        End With
    Next

'RegularInterval3�v���V�[�W���̎��s�𒆒f�����܂��̂ŁA
'�uSchedule�v�ɁuFalse�v���w�肵�܂��B
    Application.OnTime EarliestTime:=ResTime, _
    Procedure:="RegularInterval3", Schedule:=False

End Sub


'���b�Z�[�W�{�b�N�X�̕\��
'�߂�l�F���b�Z�[�W�{�b�N�X
Function YesorNo() As VbMsgBoxResult
    YesorNo = MsgBox("���̏ꏊ��" & ActiveWorkbook.Name & _
                        "�Ƃ������O�̃t�@�C�������ɂ���܂��B�u�������܂����H", _
                        vbInformation + vbYesNoCancel + vbDefaultButton2)
End Function


'�u�b�N�S�̂̕ۑ�
Sub Savebook()
    Dim dotPoint     As String
    Dim workbookName As String
    Dim fps          As Double

    '�t���[�����[�g���擾
    fps = getFps()

    If YesorNo() = vbYes Then

        Call stopUpdate
        Call Module2.fixSheetJisya(fps)

        dotPoint = InStrRev(ActiveWorkbook.Name, ".")
        workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)
        Call Module2.outputCaption(workbookName)

        ActiveWorkbook.Save

        Call restartUpdate
    End If
End Sub


'sheet�̍����牽�Ԃɑ����邩���肷��
'����1�F�V�[�g
'�߂�l�F�V�[�g�����Ԗڂɑ����Ă��邩
Function calcSheetNamePlace(ws As Worksheet)
    Dim shNameArray()   As String
    Dim i               As Long
    Dim iend            As Long
    Dim ret             As Long: ret = 0

    shNameArray() = call_GetSheetNameToArrayspecific(ThisWorkbook, "�p���]���C���V�[�g")
    iend = UBound(shNameArray)
    For i = 1 To iend
        If ws.Name = shNameArray(i) Then
            ret = i - 1
        End If
    Next
    calcSheetNamePlace = ret
End Function

'�u�b�N�����猟��
'����1�F���[�N�u�b�N
'����2�F���[�N�V�[�g���i�p���]���C���V�[�g�j
'�߂�l�F�p���]���C���V�[�g�̖��O���i�[�����z��
Function call_GetSheetNameToArrayspecific(wb As Workbook, str As String)

    Dim tmp()   As String
    Dim ws      As Worksheet
    Dim i       As Long: i = 0
    Dim istr    As Long
    Dim scnt    As Long
    scnt = wb.Worksheets.Count

    For istr = 1 To scnt
        Set ws = wb.Worksheets(istr)
        If InStr(ws.Name, str) > 0 Then
            i = i + 1
            ReDim Preserve tmp(1 To i)
            tmp(i) = ws.Name
        End If
    Next

    call_GetSheetNameToArrayspecific = tmp
End Function


'�ȈՓI�ȃV�[�g�ؑ֏��������˂����̐}�`
'����1�F�p���]���C���V�[�g
'����2�F�O�y�[�W�Ɉړ�����A�C�R�����\���ɂ��邩�ǂ����itrue or false�j
'����3�F���y�[�W�Ɉړ�����A�C�R�����\���ɂ��邩�ǂ����itrue or false�j
Private Sub addPageShape(ws As Worksheet, pPageState As Boolean, nPageState As Boolean)
    Const pPage As String = "prevPage"
    Const nPage As String = "nextPage"

    Call initCellPlace(ws)

    ws.Shapes(pPage).Visible = pPageState
    ws.Shapes(nPage).Visible = nPageState
End Sub


'�}�`��Visible���ǂ������肷��
'����1�F���[�N�V�[�g
'����2�F�}�`�̖��O
'�߂�l:Visible���ǂ����i0 or 1�j
Private Function getPageShapeState(ws As Worksheet, shapeName As String)
    getPageShapeState = ws.Shapes(shapeName).Visible
End Function


'���[�N�V�[�g���R�s�[���A�E�ɑ}��
Sub createSheet(endclm As Long)
    ThisWorkbook.ActiveSheet.Copy After:=ActiveSheet
End Sub


'�ЂƂO�̃V�[�g���A�N�e�B�u�ɂ��A�f�[�^�̍Ō���܂ōs��
Sub prevPage_Click()
    ThisWorkbook.ActiveSheet.Previous.Activate
    Call finCellPlace(ThisWorkbook.ActiveSheet)
End Sub


'�ЂƂ��̃V�[�g���A�N�e�B�u�ɂ��A�f�[�^�̍ŏ��ɍs��
Sub nextPage_Click()
    ThisWorkbook.ActiveSheet.Next.Activate
    Call initCellPlace(ThisWorkbook.ActiveSheet)
End Sub


'�Z���̏����ʒu
Private Sub initCellPlace(ws As Worksheet)
    ws.Cells(TIME_ROW, TIME_COLUMN_LEFT).Select
End Sub


'�Z���̍ŏI�ʒu
Private Sub finCellPlace(ws As Worksheet)
    ws.Cells(TIME_ROW, ws.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column).Select
End Sub


'�i�K�I�ɃT�C�Y�̕ύX����������ׂ̊֐�
'����1�F��ʂ̊g�嗦
'����2�F�T�C�Y��ύX�ł��邩�ǂ���
'�߂�l�FSmall = 1
'        Medium = 2
'        Large = 4
'        LL = 6
Private Function sizeNext(wSize As widthSize, nextChange As Boolean)
    Dim tmpsize As widthSize

    Select Case wSize
        Case widthSize.Small
            If nextChange Then
                tmpsize = widthSize.Medium
                Call changeBtnState(REDUCEBTN_NAME, True)
            Else
                tmpsize = widthSize.Small
                Call changeBtnState(EXPANDBTN_NAME, True)
'                �x�[�X�t�@�C���̕ۑ��������������p
                Call changeBtnState(REDUCEBTN_NAME, False)

            End If
        Case widthSize.Medium
            If nextChange Then
                tmpsize = widthSize.Large
            Else
                tmpsize = widthSize.Small
                Call changeBtnState(EXPANDBTN_NAME, True)
                Call changeBtnState(REDUCEBTN_NAME, False)
            End If
        Case widthSize.Large
            If nextChange Then
                tmpsize = widthSize.LL
                Call changeBtnState(EXPANDBTN_NAME, False)
                Call changeBtnState(REDUCEBTN_NAME, True)
            Else
                tmpsize = widthSize.Medium
            End If
        Case widthSize.LL
            '�O�ɂȂ�Ȃ��Ƃ�
            If Not nextChange Then
                tmpsize = widthSize.Large
                Call changeBtnState(EXPANDBTN_NAME, True)
            Else
                tmpsize = widthSize.LL
                Call changeBtnState(REDUCEBTN_NAME, True)
'                �x�[�X�t�@�C���̕ۑ��������������p
                Call changeBtnState(EXPANDBTN_NAME, False)
            End If
    End Select
    sizeNext = tmpsize
End Function

Sub doNothing_btn()
    '�Ȃɂ����Ȃ�
End Sub

'�������p�̃{�^���Ɏg���\��B���ۖ��O�������߂邱�Ƃ��ł���΂Ȃ�Ƃł��Ȃ�B

'����1�F�{�^���̖��O�iEXPANDBTN_NAME or REDUCEBTN_NAME�j
'����2�F�{�^���������邩�ǂ���
Private Sub changeBtnState(btnName As String, btnstate As Boolean)
    Dim iend, i As Long
    Dim dajsht() As String

    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "�p���]���C���V�[�g")
    iend = UBound(dajsht)
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes(btnName).Visible = btnstate
        End With
    Next
End Sub

'�V�[�g�����Z�b�g����
Sub resetSheet()
    Const pPage As String = "prevPage"
    Const nPage As String = "nextPage"
    Dim iend, i As Long
    Dim dajsht() As String
    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "�p���]���C���V�[�g")
    iend = UBound(dajsht)
    For i = 1 To iend
        With Worksheets(dajsht(i))
            '�S�ĉB��
            .Shapes(EXPANDBTN_NAME).Visible = True
            .Shapes(REDUCEBTN_NAME).Visible = True
            .Shapes(pPage).Visible = False
            .Shapes(nPage).Visible = False
        End With
        Worksheets(dajsht(i)).Range("G2:G22").Select
        Worksheets(dajsht(i)).Range(Selection, Selection.End(xlToRight)).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

        Worksheets(dajsht(i)).Range("FN2:FN22").Select
        Worksheets(dajsht(i)).Range(Selection, Selection.End(xlToRight)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

        Worksheets(dajsht(i)).Range("G22:XFD23").Select
        Selection.ClearContents
    Next

End Sub


'��\���̖��O�̒�`���ĕ\���@20230215�@����@�V�[�g�R�s�[���ɔ�������G���[�΍�
Public Sub ShowInvisibleNames()
    Dim oName As Object
    For Each oName In Names
        If oName.Visible = False Then
            oName.Visible = True
        End If
    Next
    MsgBox "��\���̖��O�̒�`��\�����܂����B", vbOKOnly
End Sub


Private Sub UserForm_Terminate()
    MsgBox "Excel�̉�ʂ�\�����܂�"
    Application.Visible = True
End Sub


' �I��͈͂��f�[�^�L����ƃ`�F�b�N���L�����̒l��Ԃ�
' �߂�l : True �� ��������i leftCol/rightCol ���Ԃ� �j
'          False �� �����Ȃ��i���b�Z�[�W�͌Ăяo�����Łj
Public Function CropSelectionToDataArea(ByRef leftCol As Long, ByRef rightCol As Long) As Boolean
    Const PAGE_FRAME_MAX    As Long = LIMIT_COLUMN '16200
    Dim shtPage             As Long
    Dim baseClm             As Long
    Dim selR                As Long '�I���
    Dim frmR                As Long '�I���t���[��
    Dim pageFrmR            As Long '�y�[�W�̗L���t���[��
    Dim totalFrm            As Long

    '�{�^������ꏏ�ɑI�񂾂疳��
    If Selection.Column > Columns.Count Then
        Exit Function
    End If

    selR = Selection.Column + Selection.Columns.Count - 1 '�I���̒���

    shtPage = calcSheetNamePlace(ActiveSheet)
    baseClm = LIMIT_COLUMN * shtPage

    With Worksheets("�|�C���g�v�Z�V�[�g")
        totalFrm = .Cells(1, 3).End(xlDown).row - 1
    End With

    '�� �� �t���[����
    frmR = selR - COLUMN_ZERO_NUM + baseClm
    pageFrmR = WorksheetFunction.min(baseClm + PAGE_FRAME_MAX, totalFrm)
    frmR = WorksheetFunction.min(frmR, pageFrmR)    '�E�ӂɂ����ăy�[�W���̗L���t���[�����𒴂��Ȃ��悤�ɂ���

    '�p���f�_�C���V�[�g�Ŏn�܂�̗񂩂ǂ������`�F�b�N���A�Œ�l�ȉ����I������Ă����ꍇ��COLUMN_ZERO+1
    leftCol = WorksheetFunction.Max(Selection.Column, COLUMN_ZERO_NUM + 1)
    rightCol = frmR - baseClm + COLUMN_ZERO_NUM

    If leftCol > rightCol Then
        CropSelectionToDataArea = False   '�d�Ȃ�Ȃ�
    Else
        '�t���[�� �� ��ԍ��֖߂�
        CropSelectionToDataArea = True
    End If
End Function

'fps�̒l���擾����
'�߂�l�Ffps�̒l
Function getFps() As Double
    getFps = ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 199).Value
End Function


'�ŏI�s���擾����
'�߂�l�F�ŏI�s
Function getLastRow() As Long
    getLastRow = ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(1, 3).End(xlDown).row
End Function