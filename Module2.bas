Option Explicit '�ϐ��̐錾������


'======================================================================================
'�����ݒ�V�[�g�̊e�f�[�^�̍s�ԍ��A��ԍ����` (����T�v�̒萔�������Œ�`�j
'======================================================================================
Const KOBUSHIAGE_MISSING_DOWNLIM_TIME           As Double = 1     '�i�b�j ���㌇���m�C�Y����Ɏg��
Const TEKUBI_SPEED_UPLIM_PREDICT            As Double = 10    '�ikm/h�j���z�ʒu�̕ω��ʏ���@�Օ����m�Ɏg��
Const MEAGERE_TIME_MACROUPDATEDATA          As Boolean = True 'True�̂Ƃ�MacroUpdateData�̏������Ԃ𑪒肷��
Const KOBUSHIAGE_TIME_HOSEI_COEF_WORK           As Double = 5 / 355 '���㎞�ԕ␳�W�� �ΏۍH���̒��ōł���Ǝ��Ԃ������H���́@�m�F�K�v�Ȍ�����Ԑ�/��Ǝ���
Const KOBUSHIAGE_MISSING_DILATION_SIZE          As Double = 0.33   '�i�b�j���㌇���̖c�������Ɏg�����̑傫���i�Б��j
Const KOBUSHIAGE_MISSING_EROSION_SIZE           As Double = 0.33   '�i�b�j���㌇���̎��k�����Ɏg�����̑傫���i�Б��j
Const KOBUSHIAGE_TIME_HOSEI_COEF_MISSING        As Double = 0.2     '���㎞�ԕ␳�W�� �m�F�K�v�Ȍ�����Ԑ��P������

'makeGraph�AoutputCaption�AfixGraphDataAndSheet���W���[���̒��ɏ����ݒ�V�[�g�̃Z��������l��ǂݏo����������

'======================================================================================
'�|�C���g�v�Z�V�[�g��̊e�f�[�^�̍s�ԍ��A��ԍ����`
'======================================================================================
Const COLUMN_POSE_NAME                      As Long = 1
Const COLUMN_POSE_KEEP_TIME                 As Long = 2
Const COLUMN_HIZA_R_ANGLE                   As Long = 6
Const COLUMN_HIZA_L_ANGLE                   As Long = 7
Const COLUMN_KOSHI_ANGLE                    As Long = 8
Const COLUMN_SHOOTING_DIRECTION             As Long = 9

Const COLUMN_POS_KOSHI_Z                    As Long = 13

Const COLUMN_POS_AHIKUBI_R_Z                As Long = 25
Const COLUMN_POS_AHIKUBI_L_Z                As Long = 37

Const COLUMN_POS_KATA_R_Z                   As Long = 57
Const COLUMN_POS_KATA_L_Z                   As Long = 69

Const COLUMN_POS_HIJI_R_Z                   As Long = 61
Const COLUMN_POS_HIJI_L_Z                   As Long = 73

Const COLUMN_POS_TEKUBI_R_Z                 As Long = 65
Const COLUMN_POS_TEKUBI_L_Z                 As Long = 77

Const COLUMN_ROUGH_TIME                     As Long = 201
Const COLUMN_CAPTION_WORK_NAME              As Long = 202
Const COLUMN_DATA_RESULT_ORIGIN             As Long = 203
Const COLUMN_DATA_MEASURE_SECTION           As Long = 204
Const COLUMN_DATA_PREDICT_SECTION           As Long = 205
Const COLUMN_DATA_REMOVE_SECTION            As Long = 206
Const COLUMN_DATA_FORCED_SECTION            As Long = 207
Const COLUMN_DATA_RESULT_FIX                As Long = 208
Const COLUMN_DATA_RESULT_GREEN              As Long = 209
Const COLUMN_DATA_RESULT_YELLOW             As Long = 210
Const COLUMN_DATA_RESULT_RED                As Long = 211
Const COLUMN_CAPTION_TRACK1                 As Long = 212 '�f�o�b�O�p�i���i�͎g��Ȃ��j

Const COLUMN_DATA_MISSING_SECTION           As Long = 219

Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_ORG           As Long = 221
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_ORG           As Long = 222
Const COLUMN_KOBUSHIAGE_FORCED_SECTION                     As Long = 223 '����A���Ȃ��A�G�Ȃ��̋����A����t���O�A�t���O�̋L��
Const COLUMN_KOBUSHIAGE_RESULT                             As Long = 245
Const COLUMN_DATA_KOSHIMAGE_MEASURE_SECTION            As Long = 225
Const COLUMN_DATA_KOSHIMAGE_PREDICT_SECTION            As Long = 226
Const COLUMN_DATA_KOSHIMAGE_MISSING_SECTION            As Long = 227
Const COLUMN_KOSHIMAGE_FORCED_SECTION                  As Long = 228
Const COLUMN_KOSHIMAGE_RESULT                          As Long = 247
Const COLUMN_DATA_HIZAMAGE_MEASURE_SECTION             As Long = 230
Const COLUMN_DATA_HIZAMAGE_PREDICT_SECTION             As Long = 231
Const COLUMN_DATA_HIZAMAGE_MISSING_SECTION             As Long = 232
Const COLUMN_HIZAMAGE_FORCED_SECTION                   As Long = 233
Const COLUMN_HIZAMAGE_RESULT                           As Long = 249
Const COLUMN_CAPTION_TRACK2                            As Long = 235 '�f�o�b�O�p�i���i�͎g��Ȃ��j

Const COLUMN_TEKUBI_RZ_SPEED                           As Long = 237 '�E���y�ʒu�̍�
Const COLUMN_TEKUBI_LZ_SPEED                           As Long = 238 '�����y�ʒu�̍�
Const COLUMN_TEKUBI_Z_SPEED_OVER                       As Long = 239 '���y�ʒu�̍� �������l�����t���O
Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_DST           As Long = 240 '���㑪����
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST           As Long = 241 '���㌇�����
Const COLUMN_MEAGERE_TIME_MACROUPDATEDATA              As Long = 242 'MacroUpdateData�̏������Ԃ𑪒茋�ʂ��i�[����

Const COLUMN_DATA_RESULT_GH_KOBUSHIAGE      As Long = 245
Const COLUMN_DATA_RESULT_GH_KOSHIMAGE       As Long = 247
Const COLUMN_DATA_RESULT_GH_HIZAMAGE        As Long = 249
Const COLUMN_DATA_RESULT_GH_SONKYO          As Long = 251

Const COLUMN_GH_HIZA_L                      As Long = 252
Const COLUMN_GH_HIZA_R                      As Long = 253

Const COLUMN_MAX_NUMBER                                As Long = 256 '���ݎg�p����Ă����ԍ��̍ő�l


'======================================================================================
'�p���d�ʓ_�����[�V�[�g�̊e�f�[�^�̍s�ԍ��A��ԍ����`
'======================================================================================
Const SHIJUTEN_SHEET_ROW_KOUTEI_NAME                            As Long = 3
Const SHIJUTEN_SHEET_ROW_POSESTART_INDEX                        As Long = 9
Const SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK                    As Long = 29

Const SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD                      As String = "���̑��̎��ԁi�莞�ғ�����7.5H-�����׎��ԁj"


Const SHIJUTEN_SHEET_COLUMN_WORK_NUMBER                         As Long = 2
Const SHIJUTEN_SHEET_COLUMN_WORK_NAME                           As Long = 3
Const SHIJUTEN_SHEET_COLUMN_KOUTEI_NAME                         As Long = 4
Const SHIJUTEN_SHEET_COLUMN_WORK_TIME                           As Long = 9
Const SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX                    As Long = 10

Const SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME                      As Long = 36
Const SHIJUTEN_SHEET_COLUMN_WORKEND_TIME                        As Long = 38


Const SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION                As Long = 46
Const SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION                As Long = 47

Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME                         As Long = 49 '���㎞��
Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME                      As Long = 51 '���Ȃ�����
Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME                       As Long = 53 '�G�Ȃ�����


Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_MISSING_TIME                 As Long = 55 '���㌇�����

Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME              As Long = 57 '���Ȃ��������
Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME              As Long = 58 '���Ȃ�������

Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME               As Long = 60 '�G�Ȃ��������
Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME               As Long = 61 '�G�Ȃ�������


'======================================================================================
'�H���]���V�[�g�̊e�f�[�^�̍s�ԍ��A��ԍ����`
'======================================================================================

Const GH_HYOUKA_SHEET_ROW_YOUSO_HANTEI_LIMIT           As Long = 2 '�v�f���肵�����l
Const GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_LIMIT           As Long = 3 '�v�f���肵�����l
Const GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_LIMIT            As Long = 4 '�v�f���肵�����l
Const GH_HYOUKA_SHEET_ROW_KOUTEI_NAME                  As Long = 5 '�H����
Const GH_HYOUKA_SHEET_ROW_DATE                         As Long = 6 '������
Const GH_HYOUKA_SHEET_ROW_POSESTART                    As Long = 15
Const GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK          As Long = 115

Const GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_CAPTION         As Long = 13 '�v�f���肵�����l�̃L���v�V�����L���Z���s
Const GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_CAPTION          As Long = 13 '�v�f���肵�����l�̃L���v�V�����L���Z���s

Const GH_HYOUKA_SHEET_EXPAND_NUM_CHECK_WORD            As String = "���v"
Const GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_NG             As String = "�~"
Const GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_OK             As String = "��"

Const GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER               As Long = 2 '���No.
Const GH_HYOUKA_SHEET_COLUMN_WORK_NAME                 As Long = 3 '�v�f��Ɩ�
Const GH_HYOUKA_SHEET_COLUMN_KOUTEI_NAME                 As Long = 4 '�H�����ƒ�����
Const GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT       As Long = 12 '�v�f���茋��
Const GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME            As Long = 13 '��ƊJ�n����
Const GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME              As Long = 16 '��ƏI������
Const GH_HYOUKA_SHEET_COLUMN_WORK_TIME                 As Long = 19 '��Ǝ���
Const GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME           As Long = 20 '���グ����
Const GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME            As Long = 21 '���Ȃ�����
Const GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME             As Long = 22 '�G�Ȃ�����
Const GH_HYOUKA_SHEET_COLUMN_NG_TIME_A                 As Long = 26 '��J�x�]���iA)��NG��Ǝ���
Const GH_HYOUKA_SHEET_COLUMN_NG_TIME_B                 As Long = 29 '�C�����x�]���iB)��NG��Ǝ���
Const GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT              As Long = 36 '�v�f���肵�����l

Const GH_HYOUKA_SHEET_COLUMN_KOSHI_HANTEI_CAPTION      As Long = 21 '�v�f���肵�����l�̃L���v�V�����L���Z����
Const GH_HYOUKA_SHEET_COLUMN_HIZA_HANTEI_CAPTION       As Long = 22 '�v�f���肵�����l�̃L���v�V�����L���Z����



'======================================================================================
'�O�̗p�@�p������̂������l���`
'======================================================================================

'Const GH_ANGLE_KOSHIMAGE_MIN As Double = 30
Const GH_ANGLE_KOSHIMAGE_MAX As Double = 180
'Const GH_ANGLE_HIZAMAGE_MIN  As Double = 60
Const GH_ANGLE_HIZAMAGE_MAX As Double = 180

'======================================================================================
'DataAdjustingSheet�p
'======================================================================================
'debug
'Const LIMIT_COLUMN           As Long = 800
Const LIMIT_COLUMN           As Long = 16200

'======================================================================================
'�������̒�`
'======================================================================================
Const CAPTION_TRACK2_FILE_NAME_SOEJI           As String = "2" '�����g���b�N�Q�p�̃t�@�C���������ɂ���Y��
Const CAPTION_CHUKAN_FILE_NAME_SOEJI           As String = "tmp" '���ԃt�@�C���ɂ���Y��
'�e�펚���̃t�H���g�T�C�Y�W��
'����̒l�̂��߁A�l���������قǕ����͑傫��
'���悪�c�̎�
Const TRACK1_TATE_UPPER_COEF                   As Long = 22 '�g���b�N1�p�F��i
Const TRACK1_TATE_LOWER_COEF                   As Long = 11 '�g���b�N1�p�F���i
Const TRACK2_TATE_1ST_COEF                     As Long = 22 '�g���b�N2�p�F�P�i��
Const TRACK2_TATE_2ND_COEF                     As Long = 22 '�g���b�N2�p�F�Q�i��
Const TRACK2_TATE_3RD_COEF                     As Long = 13 '�g���b�N2�p�F�R�i��

'���悪���̎�
Const TRACK1_YOKO_UPPER_COEF                   As Long = 30 '�g���b�N1�p�F��i
Const TRACK1_YOKO_LOWER_COEF                   As Long = 15 '�g���b�N1�p�F���i
Const TRACK2_YOKO_1ST_COEF                     As Long = 30 '�g���b�N2�p�F�P�i��
Const TRACK2_YOKO_2ND_COEF                     As Long = 30 '�g���b�N2�p�F�Q�i��
Const TRACK2_YOKO_3RD_COEF                     As Long = 18 '�g���b�N2�p�F�R�i��

'�e�펚���̐F
Const COLOR_DATA_REMOVE_SECTION                As String = "#bfbfbf" '�O���[
Const COLOR_DATA_FORCED_SECTION                As String = "#0033cc" '�F
Const COLOR_DATA_MISSING_SECTION               As String = "#ff7c80" '��F
Const COLOR_DATA_PREDICT_SECTION               As String = "#fcf600" '���F
Const COLOR_DATA_MEASURE_SECTION               As String = "#00b0f0" '���F
Const COLOR_DATA_RESULT_GREEN                  As String = "#00b050" '�ΐF
Const COLOR_DATA_RESULT_YELLOW                 As String = "#ffc000" '���F
Const COLOR_DATA_RESULT_RED                    As String = "#c00000" '�ԐF
Const COLOR_DATA_RESULT_GLAY                   As String = "#bfbfbf" '�O���[

'�уO���t�̃f�[�^�i�M���x�j����������������i�����g���b�N1�p ��i�E���ɕ\���j
Const CAPTION_DATA_MEASURE_SECTION             As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_PREDICT_SECTION             As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_REMOVE_SECTION              As String = "�y�f�[�^���O��ԁz"
Const CAPTION_DATA_FORCED_SECTION              As String = "�y�f�[�^������ԁz"
Const CAPTION_DATA_MISSING_SECTION             As String = "�y�f�[�^������ԁz"

'�уO���t�̃f�[�^�i�M���x�j����������������i�����g���b�N2�p 2�i�ڂɕ\���j
Const CAPTION_DATA_TRACK2_MEASURE_SECTION      As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_TRACK2_PREDICT_SECTION      As String = "�y�f�[�^�����ԁz"
Const CAPTION_DATA_TRACK2_REMOVE_SECTION       As String = "�y�f�[�^���O��ԁz"
Const CAPTION_DATA_TRACK2_FORCED_SECTION       As String = "�y�f�[�^������ԁz"
Const CAPTION_DATA_TRACK2_MISSING_SECTION      As String = "�y�f�[�^������ԁz"

'�O�̗p�̎���������i�����g���b�N2�p 3�i�ڂɕ\���j
Const CAPTION_A_RESULT_NAME1  As String = "�@�@�@�@����"
Const CAPTION_B_RESULT_NAME1  As String = "  �@�@���Ȃ��@ �@"
Const CAPTION_C_RESULT_NAME1  As String = "�G�Ȃ�"

'�O�̗p�̏�������������i�����g���b�N2�p 4�i�ڂɕ\���j
Const CAPTION_A_RESULT_NAME2  As String = "��񂪌�����"
Const CAPTION_B_RESULT_NAME2  As String = "45���ȏ�"
Const CAPTION_C_RESULT_NAME2  As String = "60���ȏ�"

'�L���v�V�����m�C�Y������臒l
Const CAPTION_REMOVE_NOISE_SECOND              As Double = 0.1 '�L���v�V�����m�C�Y���������钷��(�b) �i�`�����Ȃ珜���j

'�p���f�_�̒l�ɂ���āA�΁^���^�Ԃ𕪂���ۂ̋��E����
Const DATA_SEPARATION_GREEN_BOTTOM             As Long = 1
Const DATA_SEPARATION_GREEN_TOP                As Long = 2
Const DATA_SEPARATION_YELLOW_BOTTOM            As Long = 3
Const DATA_SEPARATION_YELLOW_TOP               As Long = 5
Const DATA_SEPARATION_RED_BOTTOM               As Long = 6
Const DATA_SEPARATION_RED_TOP                  As Long = 10



'======================================================================================
'Sheet1�V�[�g�p
'======================================================================================
'Const COLUMN_CYCLE           As Long = 1
'Const COLUMN_POSE_NUM        As Long = 2
'Const COLUMN_POSE_NAME       As Long = 3
'Const COLUMN_TYPE            As Long = 4
'Const COLUMN_COMPANY_TYPE    As Long = 5
'Const COLUMN_POSE_START_TIME As Long = 6
'Const COLUMN_POSE_KEEP_TIME  As Long = 7
'Const COLUMN_MOVE            As Long = 8
'Const COLUMN_forced          As Long = 9
'Const COLUMN_COMPARTINO      As Long = 10




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



'������s�̍�������i�������폜����֐�
' ����1 �F������
' ����2 �F�폜������
' �߂�l�F�폜��̕�����
Function cutLeftString(s, i As Long) As String
    Dim iLen As Long '// ������

    '// ������ł͂Ȃ��ꍇ
    If VarType(s) <> vbString Then
        Exit Function
    End If

    iLen = Len(s)

    '// �����񒷂��w�蕶�������傫���ꍇ
    If iLen < i Then
        Exit Function
    End If

    '// �w�蕶�������폜���ĕԂ�
    cutLeftString = Right(s, iLen - i)
End Function


'�c������
' ����1 �F�����O�̔z��
' ����2 �F�z��̐�
' ����3 �F���̑傫��
' �߂�l�F������̔z��
Function dilation(array_src() As Long, max_array_num As Long, window_size As Long)
        Dim i As Long
        Dim j As Long
        Dim array_dst() As Long

        '���T�C�Y���̒[�̃t���O��������̂�h�~
        array_dst = array_src

        For i = 2 + window_size To max_array_num - window_size
            If array_src(i, 0) = 1 Then
                For j = i - window_size To i + window_size
                    array_dst(j, 0) = 1
                Next
            End If
        Next

        dilation = array_dst()

End Function


'���k����
' ����1 �F�����O�̔z��
' ����2 �F�z��̐�
' ����3 �F���̑傫��
' �߂�l�F������̔z��
Function erosion(array_src() As Long, max_array_num As Long, window_size As Long)
        Dim i As Long
        Dim j As Long
        Dim array_dst() As Long

        '���T�C�Y���̒[�̃t���O��������̂�h�~
        array_dst = array_src

        For i = 2 + window_size To max_array_num - window_size
            If array_src(i, 0) = 0 Then
                For j = i - window_size To i + window_size
                    array_dst(j, 0) = 0
                Next
            End If
        Next

        erosion = array_dst()

End Function


'����̃t���O����
' ����1 �F�Ȃ�
' �߂�l�F�Ȃ�

Sub kobusiage_hantei()

    '�\���E�X�V���I�t�ɂ���
    Call stopUpdate

    Dim KataPositionRz()       As Double
    Dim KataPositionLz()       As Double
    Dim TekubiPositionRz()     As Double
    Dim TekubiPositionLz()     As Double
    Dim TekubiSpeedRz()        As Double
    Dim TekubiSpeedLz()        As Double

    Dim kobushiage_missing_array()     As Long '���㌇���t���O���i�[����z��
    Dim tekubi_zspeed_over_array()     As Long '���ʒu�̍��̂������l����
    Dim kobushiage_measure_array()     As Long '���㑪����
    Dim kobushiage_array()             As Long '���㎞��

    Dim kobushiage_missing_limit       As Long '���㌇���t���O�̃m�C�Y���肵�����l
    Dim kobushiage_missing_count       As Long '���㌇���t���O�̘A���񐔂��J�E���g
    Dim kobushiage_missing_section_num As Long '���㌇����Ԃ��������l�𒴂��鐔���J�E���g�B���㎞�Ԃ̕␳�Ɏg��

    Dim window_size_dilation           As Long '�c���Ɏg�����̑傫��
    Dim window_size_erosion            As Long '���k�Ɏg�����̑傫��
    Dim work_time                      As Double  '��Ǝ���
    Dim kobushiage_hosei_time          As Double  '����␳����
    Dim kobushiage_hosei_frame_num     As Long '����␳�t���[����

    Dim fps                            As Double '�t���[�����[�g

    Dim max_row_num                    As Long '�s�̖���
    Dim max_array_num                  As Long '�z��̖���

    Dim i                              As Long
    Dim j                              As Long

    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        '��������s�����擾�i3��ڂ̍ŏI�Z���j
        max_row_num = .Cells(1, 3).End(xlDown).row
        'MsgBox ("max_row_num=" & max_row_num)

        max_array_num = max_row_num - 1 - 1 '2�s�ڂ���Z���ɒl�����邽��-1�A�z���0����g������-1

        ReDim KataPositionRz(max_array_num, 0)
        ReDim KataPositionLz(max_array_num, 0)
        ReDim TekubiPositionRz(max_array_num, 0)
        ReDim TekubiPositionLz(max_array_num, 0)
        ReDim TekubiSpeedRz(max_array_num, 0)
        ReDim TekubiSpeedLz(max_array_num, 0)

        ReDim kobushiage_missing_array(max_array_num, 0)
        ReDim tekubi_zspeed_over_array(max_array_num, 0)
        ReDim kobushiage_measure_array(max_array_num, 0)
        ReDim kobushiage_array(max_array_num, 0)


        '-------------------------------------------------------------------------------------------
        '�������猝�㔻��
        '-------------------------------------------------------------------------------------------
        fps = .Cells(2, 199).Value '�t���[�����[�g���擾

        '���ƌ��̍�����z��Ɋi�[
        For i = 0 To max_array_num
            '��,���̍���(Z���W)��ǂݏo���@�z���0����n�܂邽��+1�A�Z���̒l��2�s�ڂ���n�܂邽��+1
            TekubiPositionLz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_TEKUBI_L_Z).Value
            TekubiPositionRz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_TEKUBI_R_Z).Value
            KataPositionLz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_KATA_L_Z).Value
            KataPositionRz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_KATA_R_Z).Value
            kobushiage_missing_array(i, 0) = .Cells(i + 1 + 1, COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_ORG).Value
        Next

        '�f�o�b�N
'        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1)).Value = TekubiPositionLz
'        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE)).Value = TekubiPositionRz

        '���㔻������āA���ʂ�z��Ɋi�[
        For i = 0 To max_array_num
            If _
                KataPositionLz(i, 0) < TekubiPositionRz(i, 0) Or _
                KataPositionLz(i, 0) < TekubiPositionLz(i, 0) Or _
                KataPositionRz(i, 0) < TekubiPositionLz(i, 0) Or _
                KataPositionRz(i, 0) < TekubiPositionRz(i, 0) Then
                kobushiage_array(i, 0) = 1
            Else
                kobushiage_array(i, 0) = 0
            End If
        Next

        '�f�o�b�N
'        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1)).Value = kobushiage_array
'        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE)).Value = kobushiage_array


        '���㔻��Z���֔z��̒l����������
'        .Range(.Cells(2, COLUMN_KOBUSHIAGE_RESULT), .Cells(max_row_num, COLUMN_KOBUSHIAGE_RESULT)).Value = kobushiage_array�f�f�o�b�N�p


'��������������������������������A�m�C�Y�ɂȂ邽�߃R�����g�A�E�g������������������������
'        '-------------------------------------------------------------------------------------------
'        '����������z�����̑��x�v�Z
'        '-------------------------------------------------------------------------------------------
'        '�O�t���[�����������ߔz��̐擪��0������
'        TekubiSpeedRz(0, 0) = 0
'        TekubiSpeedLz(0, 0) = 0
'
'        '���Z�����̑��x���v�Z
'        For i = 1 To max_array_num '�A�z���0����n�܂邪�A�O�t���[���������ƌv�Z�o���Ȃ�����1����v�Z���n�߂�
'            '��񑬓x���v�Z���ĒP�ʂ�ύX�imm/sec��km/h�j
'            TekubiSpeedRz(i, 0) = CDbl(Abs(TekubiPositionRz(i, 0) - TekubiPositionRz(i - 1, 0)) * fps * 60 * 60 / 1000000)
'            TekubiSpeedLz(i, 0) = CDbl(Abs(TekubiPositionLz(i, 0) - TekubiPositionLz(i - 1, 0)) * fps * 60 * 60 / 1000000)
'        Next
'
'        '���y�����̑��x���L�^����
''        .Range(.Cells(2, COLUMN_TEKUBI_RZ_SPEED), .Cells(max_row_num, COLUMN_TEKUBI_RZ_SPEED)).Value = TekubiSpeedRz'�f�o�b�N�p
''        .Range(.Cells(2, COLUMN_TEKUBI_LZ_SPEED), .Cells(max_row_num, COLUMN_TEKUBI_LZ_SPEED)).Value = TekubiSpeedLz'�f�o�b�N�p
'
'        '���y�����̑��x��l�����̃t���O�������āA�z��Ɋi�[
'        For i = 0 To max_array_num
'            If TEKUBI_SPEED_UPLIM_PREDICT <= TekubiSpeedRz(i, 0) Or _
'                TEKUBI_SPEED_UPLIM_PREDICT <= TekubiSpeedLz(i, 0) Then
'                tekubi_zspeed_over_array(i, 0) = 1
'            Else
'                tekubi_zspeed_over_array(i, 0) = 0
'            End If
'        Next
'
'
'        '-------------------------------------------------------------------------------------------
'        '�������猝�㌇���m�C�Y��������
'        '-------------------------------------------------------------------------------------------
'
'        '�A�����J�E���g�ϐ����`
'        kobushiage_missing_limit = CLng(fps * KOBUSHIAGE_MISSING_DOWNLIM_TIME)
'        kobushiage_missing_count = 0
'        kobushiage_missing_section_num = 0
'
'        'kobushiage_missing_array�̒���1����������A�������J�E���g����
'        For i = 0 To max_array_num
'            '�z��1�̂Ƃ�
'            If kobushiage_missing_array(i, 0) > 0 Then
'                kobushiage_missing_count = kobushiage_missing_count + 1
'            Else '�z��0�̂Ƃ�
'                'kobushiage_missing_count����������΃m�C�Y�Ƃ��ď���
'                If kobushiage_missing_count < kobushiage_missing_limit Then
'                    For j = (i - kobushiage_missing_count) To i
'                        kobushiage_missing_array(j, 0) = 0
'                    Next
'                End If
'
'                '�J�E���g���Z�b�g
'                kobushiage_missing_count = 0
'
'            End If
'        Next
'
'        '�z��̒l���Z���ɑ��
'        '.Range(.Cells(2, 247), .Cells(max_row_num, 247)).Value = kobushiage_missing_array '�f�o�b�N�p
'
'
'        '-------------------------------------------------------------------------------------------
'        '�������猝�㌇����ԏC��
'        '-------------------------------------------------------------------------------------------
'
'        '���x��z���̃t���O�@�Ɓ@�m�C�Y�����������㌇���t���O�𑫂����킹��
'        For i = 0 To max_array_num
'             If tekubi_zspeed_over_array(i, 0) > 0 Or kobushiage_missing_array(i, 0) > 0 Then
'               kobushiage_missing_array(i, 0) = 1
'            End If
'        Next
'
'        '�c���E���k�����̑��̑傫���i�Б��j�����߂�
'        window_size_dilation = CLng(KOBUSHIAGE_MISSING_DILATION_SIZE * fps)
'        window_size_erosion = CLng(KOBUSHIAGE_MISSING_EROSION_SIZE * fps)
'
''        .Range(.Cells(2, 251), .Cells(max_row_num, 251)).Value = kobushiage_missing_array '�f�o�b�N�p
'
'        '�����ߏ���
'        '�c��
'        kobushiage_missing_array() = dilation(kobushiage_missing_array(), max_array_num, window_size_dilation)
''        .Range(.Cells(2, 252), .Cells(max_row_num, 252)).Value = kobushiage_missing_array '�f�o�b�N�p
'
'        '���k
'        kobushiage_missing_array() = erosion(kobushiage_missing_array(), max_array_num, window_size_erosion)
''        .Range(.Cells(2, 253), .Cells(max_row_num, 253)).Value = kobushiage_missing_array '�f�o�b�N�p
'
'        '���グ����t���O�̐���
'        For i = 0 To max_array_num
'            If kobushiage_missing_array(i, 0) = 0 Then
'                kobushiage_measure_array(i, 0) = 1
'            Else
'                kobushiage_measure_array(i, 0) = 0
'            End If
'        Next
'
'        '���㑪��A�����Z���֔z��̒l����������
'        .Range(.Cells(2, COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_DST), .Cells(max_row_num, COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_DST)).Value = kobushiage_measure_array
'        .Range(.Cells(2, COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST), .Cells(max_row_num, COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST)).Value = kobushiage_missing_array
'
'
'        '-------------------------------------------------------------------------------------------
'        '�������猝��t���O�C��
'        '-------------------------------------------------------------------------------------------
'
'        '�J�E���g���Z�b�g
'        kobushiage_missing_count = 0
'        kobushiage_missing_section_num = 0
'
'        'kobushiage_missing_array����1�̘A�������J�E���g����
'        For i = 0 To max_array_num
'            '�z��1�̂Ƃ�
'            If kobushiage_missing_array(i, 0) > 0 Then
'                kobushiage_missing_count = kobushiage_missing_count + 1
'            Else '�z��0�̂Ƃ�
'                '���㌇����Ԃ��������l��蒷�������܂�̐����J�E���g����
'                If kobushiage_missing_count > kobushiage_missing_limit Then
'                    kobushiage_missing_section_num = kobushiage_missing_section_num + 1
'                End If
'
'                '�J�E���g���Z�b�g
'                kobushiage_missing_count = 0
'
'            End If
'        Next
'
'        '��Ǝ���(�b)���v�Z
'        work_time = CDbl(max_row_num / fps)
'
'        '���㎞�Ԃ̕␳���Ԃ��v�Z����B�i��Ǝ��Ԃƌ������Ԃ�臒l�����񐔂���j
'        kobushiage_hosei_time = (work_time * KOBUSHIAGE_TIME_HOSEI_COEF_WORK) + (KOBUSHIAGE_TIME_HOSEI_COEF_MISSING * kobushiage_missing_section_num)
'
'        '���㌇����Ԃ��������l��蒷�������܂�̐��������o��
''        ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(2, 51).Value = kobushiage_hosei_time '�f�o�b�N�p
'
'        '����␳���Ԃ���t���[�������v�Z
'        kobushiage_hosei_frame_num = WorksheetFunction.RoundUp(kobushiage_hosei_time * fps, 0)
'
'
'        '����␳���Ԃ���v�Z�����t���[����
'        '����z�ʒu�ω��ʂ��傫���t���[���̌���t���O�ɑO���珇�Ɋ���t���Ă���
'        For i = 0 To max_array_num
'            '���㎞�Ԃ��傫���Ȃ肷����h�~���邽�߂ɁA0�ɂȂ����珈������߂�
'            If kobushiage_hosei_frame_num = 0 Then
'                Exit For
'            End If
'
'            If kobushiage_array(i, 0) = 0 And tekubi_zspeed_over_array(i, 0) = 1 Then
'                kobushiage_array(i, 0) = 1
'                '����␳���Ԃ���v�Z�����t���[�������炷
'                kobushiage_hosei_frame_num = kobushiage_hosei_frame_num - 1
'            End If
'        Next
'�����������������������������܂ŁA�m�C�Y�ɂȂ邽�߃R�����g�A�E�g������������������������

        '�␳��̌��㎞�Ԕz����Z���ɓ\��t��
        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1)).Value = kobushiage_array
        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE)).Value = kobushiage_array
    End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

    '�\���E�X�V���I���ɖ߂�
    Call restartUpdate
End Sub


'����A���Ȃ��A�G�Ȃ��̔���
' ����1 �F�Ȃ�
' �߂�l�F�Ȃ�
Sub makeGraphJisya()

    '�\���E�X�V���I�t�ɂ���
    Call stopUpdate

    '����̔���
    '�R�[�h���������ߕʂ̊֐��ɂ���
    Call kobusiage_hantei

    '�����ݒ�V�[�g����ǂݍ��ރp�����[�^
    Dim AngleKoshiMin  As Double
    Dim AngleKoshiMax  As Double

    Dim AngleHizaMin   As Double
    Dim AngleHizaMax   As Double

    '�֐ߊp�x�̃f�[�^��ǂݍ��ޕϐ�
    Dim ValAngleKoshi   As Double
    Dim ValAngleHizaR   As Double
    Dim ValAngleHizaL   As Double

    '���茋�ʂ��i�[����z��
    Dim KoshimageArray() As Double
    Dim HizamageArray()  As Double

    '�O�̗p�̕G�p�x
    '�Г��ƒ�`���قȂ邽�ߒ���
    Dim HizaAngleLArray() As Double
    Dim HizaAngleRArray() As Double


    '���̑��ϐ�
    Dim start_frame            As Long
    Dim end_frame              As Long
    Dim fps                    As Double '�t���[�����[�g
    Dim correctPose            As Boolean
    Dim mSeconds               As String
    Dim totalSecond            As Long
    Dim tempSecond             As Long
    Dim hour, min, sec         As Long
    Dim t                      As Date
    Dim ds                     As String

    Dim max_row_num            As Long '�s�̖���
    Dim max_array_num          As Long '�z��̖���

    Dim i                      As Long
    Dim j                      As Long
    Dim data_no                As Long

    Dim PointCalcSheetArray As Variant




    '����̂������l����
    AngleKoshiMin = ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_LIMIT, GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT).Value
    AngleKoshiMax = GH_ANGLE_KOSHIMAGE_MAX

    AngleHizaMin = ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_LIMIT, GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT).Value
    AngleHizaMax = GH_ANGLE_HIZAMAGE_MAX

    '����̂������l�ɉ����ĕ\���̕\��������i�L���v�V�����j����������
    ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_CAPTION, GH_HYOUKA_SHEET_COLUMN_KOSHI_HANTEI_CAPTION).Value = "���p�x" & vbLf & CStr(AngleKoshiMin) & "��" & vbLf & "�ȏ�"
    ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_CAPTION, GH_HYOUKA_SHEET_COLUMN_HIZA_HANTEI_CAPTION).Value = "�G�p�x" & vbLf & CStr(AngleHizaMin) & "��" & vbLf & "�ȏ�"


    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        '��������s�����擾�i3��ڂ̍ŏI�Z���j
        max_row_num = .Cells(1, 3).End(xlDown).row

        '�f�o�b�O�p
        'MsgBox ("max_row_num=" & max_row_num)

        '�|�C���g�v�Z�V�[�g�̒��g��z��ɓǍ�
        PointCalcSheetArray = .Range(.Cells(1, 1), .Cells(max_row_num, COLUMN_MAX_NUMBER))

        max_array_num = max_row_num - 1 - 1 '2�s�ڂ���Z���ɒl�����邽��-1�A�z���0����g������-1

        ReDim HizaAngleLArray(max_array_num, 0)
        ReDim HizaAngleRArray(max_array_num, 0)

        '-------------------------------------------------------------------------------------------
        '��������G�p�x�v�Z
        '-------------------------------------------------------------------------------------------
        For i = 0 To max_array_num
            HizaAngleLArray(i, 0) = 180 - .Cells(i + 2, COLUMN_HIZA_L_ANGLE).Value
            HizaAngleRArray(i, 0) = 180 - .Cells(i + 2, COLUMN_HIZA_R_ANGLE).Value
        Next

        '-------------------------------------------------------------------------------------------
        '��������p������
        '-------------------------------------------------------------------------------------------
        For i = 2 To max_row_num

            '�L���v�V������������
            mSeconds = Right(Format(WorksheetFunction.RoundDown(PointCalcSheetArray(i, 2), 3), "0.000"), 3) '�����_�ȉ��̂ݎ擾
            totalSecond = Application.WorksheetFunction.RoundDown(.Cells(i, 2), 0)
            hour = WorksheetFunction.RoundDown(totalSecond / 3600, 0)
            min = WorksheetFunction.RoundDown((totalSecond Mod 3600) / 60, 0)
            sec = totalSecond Mod 60
            t = TimeSerial(hour, min, sec)
            ds = Format(t, "hh:mm:ss")

            '�L���v�V���������̑��
            PointCalcSheetArray(i, COLUMN_ROUGH_TIME) = Format(t, "hh:mm:ss") & "," & mSeconds

            '�֐ߊp�x�̓ǂݏo��
            ValAngleKoshi = CDbl(PointCalcSheetArray(i, COLUMN_KOSHI_ANGLE))
            ValAngleHizaL = CDbl(PointCalcSheetArray(i, COLUMN_HIZA_L_ANGLE))
            ValAngleHizaR = CDbl(PointCalcSheetArray(i, COLUMN_HIZA_R_ANGLE))

            '���Ȃ��̔���
            If ( _
                AngleKoshiMax >= ValAngleKoshi And _
                AngleKoshiMin < ValAngleKoshi _
            ) Then
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE) = 1
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE - 1) = 1

            Else
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE) = 0
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_KOSHIMAGE - 1) = 0
            End If

            '�G�Ȃ��̔���
            If _
                (AngleHizaMax >= (180 - ValAngleHizaL)) And (AngleHizaMin < (180 - ValAngleHizaL)) Or _
                (AngleHizaMax >= (180 - ValAngleHizaR)) And (AngleHizaMin < (180 - ValAngleHizaR)) _
            Then
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE) = 1
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE - 1) = 1
            Else
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE) = 0
                PointCalcSheetArray(i, COLUMN_DATA_RESULT_GH_HIZAMAGE - 1) = 0
            End If

            '�L���v�V���������̃Z�����
            PointCalcSheetArray(i, COLUMN_ROUGH_TIME) = Format(t, "hh:mm:ss") & "," & mSeconds

        Next

        '-------------------------------------------------------------------------------------------
        '��������z��̒��g���|�C���g�v�Z�V�[�g�ɏ���
        '-------------------------------------------------------------------------------------------
        .Range(.Cells(1, 1), .Cells(max_row_num, COLUMN_MAX_NUMBER)) = PointCalcSheetArray

        '�O�̗p�G�p�x
        .Range(.Cells(2, COLUMN_GH_HIZA_L), .Cells(max_row_num, COLUMN_GH_HIZA_L)).Value = HizaAngleLArray
        .Range(.Cells(2, COLUMN_GH_HIZA_R), .Cells(max_row_num, COLUMN_GH_HIZA_R)).Value = HizaAngleRArray

    End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

    '�\���E�X�V���I���ɖ߂�
    Call restartUpdate

End Sub


'�p���_�̔���
' ����1 �F�Ȃ�
' �߂�l�F�Ȃ�
Sub makeGraphZensya()

    '�\���E�X�V���I�t�ɂ���
    Call stopUpdate

    Dim KoshiMax(10)           As Double '�����g���b�N1�p �p���f�_1�`10�̊e���̏��臒l
    Dim KoshiMin(10)           As Double '�����g���b�N1�p �p���f�_1�`10�̊e���̉���臒l
    Dim HizaMax(10)            As Double '�����g���b�N1�p �p���f�_1�`10�̊e�G�̏��臒l
    Dim HizaMin(10)            As Double '�����g���b�N1�p �p���f�_1�`10�̊e�G�̉���臒l

    Dim CaptionName2(10)       As String '�p���f�_1�`10�̎���������

    Dim work_time              As Double  '��Ǝ���

    Dim KoshiAngle             As Double

    Dim HizaAngleR             As Double
    Dim HizaAngleL             As Double

    Dim start_frame            As Long
    Dim end_frame              As Long
    Dim fps                    As Double '�t���[�����[�g
    Dim correctPose            As Boolean
    Dim mSeconds               As String
    Dim totalSecond            As Long
    Dim tempSecond             As Long
    Dim hour, min, sec         As Long
    Dim t                      As Date
    Dim ds                     As String

    Dim max_row_num            As Long '�s�̖���
    Dim max_array_num          As Long '�z��̖���

    Dim i                      As Long
    Dim j                      As Long
    Dim data_no                As Long

    Dim CaptionName3Kobushiage      As String '�����g���b�N2�p �R�i�� ����̎���������
    Dim CaptionName3Koshimage       As String '�����g���b�N2�p �R�i�� ���Ȃ��̎���������
    Dim CaptionName3Hizamage        As String '�����g���b�N2�p �R�i�� �G�Ȃ��̎���������
    Dim Koshimage              As Double '�����g���b�N2�p ���Ȃ�����p��臒l
    Dim Hizamage               As Double '�����g���b�N2�p �G�Ȃ�����p��臒l


    '�e�p���̖��O�Ə����̓ǂݏo��
    'Min��Max�������I�łȂ��̂Œ���
    With ThisWorkbook.Worksheets("�����ݒ�V�[�g")
        CaptionName2(10) = .Cells(6, 2) '10�_�̎p���������l
        KoshiMax(10) = .Cells(8, 7) 'x�ȏ�
        KoshiMin(10) = .Cells(9, 7) 'x����
        HizaMax(10) = .Cells(11, 7)
        HizaMin(10) = .Cells(12, 7)

        CaptionName2(9) = .Cells(24, 2) '9�_�̎p���������l
        KoshiMax(9) = .Cells(26, 7)
        KoshiMin(9) = .Cells(27, 7)
        HizaMax(9) = .Cells(29, 7)
        HizaMin(9) = .Cells(30, 7)

        CaptionName2(8) = .Cells(42, 2) '8�_�̎p���������l
        KoshiMax(8) = .Cells(44, 7)
        KoshiMin(8) = .Cells(45, 7)
        HizaMax(8) = .Cells(47, 7)
        HizaMin(8) = .Cells(48, 7)

        CaptionName2(7) = .Cells(60, 2) '7�_�̎p���������l
        KoshiMax(7) = .Cells(62, 7)
        KoshiMin(7) = .Cells(63, 7)
        HizaMax(7) = .Cells(65, 7)
        HizaMin(7) = .Cells(66, 7)

        CaptionName2(6) = .Cells(78, 2) '6�_�̎p���������l
        KoshiMax(6) = .Cells(80, 7)
        KoshiMin(6) = .Cells(81, 7)
        HizaMax(6) = .Cells(83, 7)
        HizaMin(6) = .Cells(84, 7)

        CaptionName2(5) = .Cells(96, 2) '5�_�̎p���������l
        KoshiMax(5) = .Cells(98, 7)
        KoshiMin(5) = .Cells(99, 7)
        HizaMax(5) = .Cells(101, 7)
        HizaMin(5) = .Cells(102, 7)

        CaptionName2(4) = .Cells(114, 2) '4�_�̎p���������l
        KoshiMax(4) = .Cells(116, 7)
        KoshiMin(4) = .Cells(117, 7)
        HizaMax(4) = .Cells(119, 7)
        HizaMin(4) = .Cells(120, 7)

        CaptionName2(3) = .Cells(132, 2) '3�_�̎p���������l
        KoshiMax(3) = .Cells(134, 7)
        KoshiMin(3) = .Cells(135, 7)
        HizaMax(3) = .Cells(137, 7)
        HizaMin(3) = .Cells(138, 7)

        CaptionName2(2) = .Cells(150, 2) '2�_�̎p���������l
        KoshiMax(2) = .Cells(152, 7)
        KoshiMin(2) = .Cells(153, 7)
        HizaMax(2) = .Cells(155, 7)
        HizaMin(2) = .Cells(156, 7)

        CaptionName2(1) = .Cells(168, 2) '1�_�̎p���������l
        KoshiMax(1) = .Cells(170, 7)
        KoshiMin(1) = .Cells(171, 7)
        HizaMax(1) = .Cells(173, 7)
        HizaMin(1) = .Cells(174, 7)

        '����A���Ȃ��A�G�Ȃ��p
        '������A�������l���擾
        CaptionName3Kobushiage = .Cells(192, 2)
        CaptionName3Koshimage = .Cells(210, 2)
        Koshimage = .Cells(212, 7)
        CaptionName3Hizamage = .Cells(228, 2)
        Hizamage = .Cells(230, 7)

    End With 'With ThisWorkbook.Worksheets("�����ݒ�V�[�g")


    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        '��������s�����擾�i3��ڂ̍ŏI�Z���j
        max_row_num = .Cells(1, 3).End(xlDown).row
        'MsgBox ("max_row_num=" & max_row_num)

        max_array_num = max_row_num - 1 - 1 '2�s�ڂ���Z���ɒl�����邽��-1�A�z���0����g������-1



        '�p�����肵�ăZ���ɑ��
        For i = 2 To max_row_num

            '-------------------------------------------------------------------------------------------
            '�p������(1�`10)�𔻒肷��
            '-------------------------------------------------------------------------------------------

            For j = 2 To 10
                mSeconds = Right(Format(WorksheetFunction.RoundDown(.Cells(i, 2), 3), "0.000"), 3) '�����_�ȉ��̂ݎ擾
                totalSecond = Application.WorksheetFunction.RoundDown(.Cells(i, 2), 0)
                hour = WorksheetFunction.RoundDown(totalSecond / 3600, 0)
                min = WorksheetFunction.RoundDown((totalSecond Mod 3600) / 60, 0)
                sec = totalSecond Mod 60
                t = TimeSerial(hour, min, sec)
                ds = Format(t, "hh:mm:ss")
                correctPose = False

                '�֐ߊp�x�̓ǂݏo��
                KoshiAngle = CDbl(.Cells(i, COLUMN_KOSHI_ANGLE).Value)
                HizaAngleL = CDbl(.Cells(i, COLUMN_HIZA_L_ANGLE).Value)
                HizaAngleR = CDbl(.Cells(i, COLUMN_HIZA_R_ANGLE).Value)

                If ( _
                KoshiMin(j) >= KoshiAngle And _
                KoshiMax(j) < KoshiAngle) And (( _
                HizaMin(j) >= HizaAngleL And _
                HizaAngleL > HizaMax(j)) Or _
                HizaMin(j) >= HizaAngleR And ( _
                HizaAngleR > HizaMax(j))) Then
                    correctPose = True

                    '�L���v�V���������̃Z�����
                    .Cells(i, COLUMN_ROUGH_TIME).Value = Format(t, "hh:mm:ss") & "," & mSeconds

                    '���茋�ʂ����f�[�^�p�Z���ɓ���
                    .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value = j

                    '���茋�ʂ��������W�v�p�Z���ɓ���
                    .Cells(i, COLUMN_DATA_RESULT_FIX).Value = j

                    '�p�����ނ�����������For(j)�𔲂���
                    Exit For
                End If
            Next 'For j = 2 To 10

            '�Y���p�����Ȃ������ꍇ��1�_�̎p���ɂ���
            If correctPose = False Then
                '�L���v�V���������̃Z�����
                .Cells(i, COLUMN_ROUGH_TIME).Value = Format(t, "hh:mm:ss") & "," & mSeconds
                '���茋�ʂ����f�[�^�p�Z���ɓ���
                .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value = 1
                '���茋�ʂ��������W�v�p�Z���ɓ���
                .Cells(i, COLUMN_DATA_RESULT_FIX).Value = 1
            End If
        Next ' For i = 2 To max_row_num

        '�O���t�`��̐F�����̂��߂̃t���O����
        For i = 2 To max_row_num

            '�����s�̎p���f�_��data_no�ֈꎞ�L������
            data_no = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value

            '�p���f�_�̗΁A���A�Ԃ̐F�����`��p�f�[�^���o��
            If _
            data_no >= DATA_SEPARATION_GREEN_BOTTOM And _
            data_no <= DATA_SEPARATION_GREEN_TOP Then
                .Cells(i, COLUMN_DATA_RESULT_GREEN).Value = data_no
                .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_RED).Value = 0
            ElseIf _
            data_no >= DATA_SEPARATION_YELLOW_BOTTOM And _
            data_no <= DATA_SEPARATION_YELLOW_TOP Then
                .Cells(i, COLUMN_DATA_RESULT_GREEN).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value = data_no
                .Cells(i, COLUMN_DATA_RESULT_RED).Value = 0
            ElseIf _
            data_no >= DATA_SEPARATION_RED_BOTTOM And _
            data_no <= DATA_SEPARATION_RED_TOP Then
                .Cells(i, COLUMN_DATA_RESULT_GREEN).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value = 0
                .Cells(i, COLUMN_DATA_RESULT_RED).Value = data_no
            End If
        Next ' i = 2 To max_row_num
    End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

    '�\���E�X�V���I���ɖ߂�
    Call restartUpdate

End Sub






'�p���f�_�̎����A�t���O�̃m�C�Y����������
' ����1 �F�t���[�����[�g
' �߂�l�F�Ȃ�
Function removeCaptionNoise(fps As Double)

    Dim max_row_num   As Long
    Dim max_array_num As Long

    Dim i             As Long
    Dim j             As Long
    Dim k             As Long
    Dim tmp           As Long

    Dim i_max         As Long
    Dim j_max         As Long
    Dim k_max         As Long

    Dim currentValue  As String
    Dim targetValue   As String
    Dim compareValue  As String

    Dim sameValueNum  As Long
    Dim noise_num     As Long

    noise_num = CAPTION_REMOVE_NOISE_SECOND * fps
    If noise_num < 2 Then
        noise_num = 2
    End If

    '�\���E�X�V���I�t�ɂ���
    Call stopUpdate

    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        '��������s�����擾�i3��ڂ̍ŏI�Z���j
        max_row_num = .Cells(1, 3).End(xlDown).row
        max_array_num = max_row_num - 1 - 1 '2�s�ڂ���Z���ɒl�����邽��-1�A�z���0����g������-1
        'MsgBox ("max_row_num=" & max_row_num)

        '�������֒T������ۂ̋N�_(i), �I�_(i_max)
        i_max = max_row_num - noise_num - 1

        '�L���v�V�����̃m�C�Y����
        For i = 2 To i_max

            currentValue = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value
            targetValue = .Cells(i + 1, COLUMN_DATA_RESULT_ORIGIN).Value

            '���茋�ʂ��ς�����Ƃ�
            If currentValue <> targetValue Then

                '�m�C�Y���ǂ����T������ �N�_(j), �I�_(j_max)
                j_max = i + 1 + noise_num - 1
                sameValueNum = 1
                For j = i + 2 To j_max
                    compareValue = .Cells(j, COLUMN_DATA_RESULT_ORIGIN).Value
                    '���茋�ʂ��ς�����烋�[�v�𔲂���
                    If targetValue = compareValue Then
                        sameValueNum = sameValueNum + 1
                    Else
                        Exit For
                    End If
                Next

                '�m�C�Y�����������Ƃ��̏���
                If sameValueNum < noise_num Then
                    For k = i + 1 To j
                        If Not IsEmpty(.Cells(i, COLUMN_DATA_RESULT_ORIGIN)) Then
                            For tmp = 0 To 14
                                .Cells(k, COLUMN_DATA_RESULT_ORIGIN + tmp) = .Cells(i, COLUMN_DATA_RESULT_ORIGIN + tmp)
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

    '�\���E�X�V���I���ɖ߂�
    Call restartUpdate
End Function




' �P��ڂ͕����Ȃ��̃f�[�^����
' �X�V�{�^���������ꂽ�ۂ́A��ƊJ�n���Ԃ��g���ĕ���
' ����1 �F�t���[�����[�g
' �߂�l�F�Ȃ�
Sub fixSheetJisya()

    '�\���E�X�V���I�t�ɂ���
    Call stopUpdate

    Dim fps As Double

    Dim separate_work_time        As Double 't��t0�̍����擾����
    Dim t0                        As Double '1�O��t���ꎞ�ۑ�����
    Dim t                         As Double '��Ǝ���

    Dim i                         As Long
    Dim j                         As Long

    Dim max_row_num               As Long

    Dim expand_no                 As Long '�ǉ����ꂽ�s���𒲂ׂ邽�߂Ɏg��

    Dim Kobushiage_flag           As Long
    Dim koshimage_flag             As Long
    Dim hizamage_flag             As Long

    Dim start_frame               As Long
    Dim end_frame                 As Long

    Dim data_no                   As Long

    Dim removeFrames              As Long
    Dim workFrames                As Long

    Dim top_jogai_end             As Long
    Dim bottom_jogai_start        As Long

    Dim worktime_sum              As Double

    Dim seconds          As Double
    Dim hours            As String
    Dim minutes          As String
    Dim remainingSeconds As String
    Dim milliseconds     As String
    Dim format_time      As String

    Dim youso_hantei_limit As Double
    Dim NG_time_A As Double
    Dim NG_time_B As Double

    '�t���[�����[�g���擾
    fps = ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 199)

    '�|�C���g�v�Z�V�[�g�̍ŏI�s���擾
    max_row_num = ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(1, 2).End(xlDown).row

    '�e�평����
    removeFrames = 0

    '�v�f��Ɣ���̂������l��Ǎ�
    youso_hantei_limit = ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_YOUSO_HANTEI_LIMIT, GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT).Value

    '��������ǉ��s�����擾����
    '"�v�f��"�̃Z���ʒu�̈ړ��ʂ𒲂ׂ�  ���ő�999�s(<1050)�ɂ���
    expand_no = 0
    Do While ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK + expand_no, GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER) <> _
    GH_HYOUKA_SHEET_EXPAND_NUM_CHECK_WORD And expand_no < 1050
        expand_no = expand_no + 1
    Loop


    '�H���]���V�[�g�ɒl�����
    With ThisWorkbook.Sheets("�H���]���V�[�g")
        '�������珉�񕪐͂̂��߂̏���
        '��ƊJ�n���Ԃ���̏ꍇ�́A0.0�����
        If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) = True Then
            .Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value = 0
        End If

        '��ƏI�����Ԃ���̏ꍇ�́A�|�C���g�v�Z�V�[�g�ŏI�s����v�Z���ē���
        If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME)) = True Then
            seconds = max_row_num / fps '�����ɕϊ��������b������͂��Ă�������

            format_time = timeConvert(seconds)

            .Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value = format_time

        End If

        '�������璠�[�X�V�̂��߂̏���
        '����̐擪�ɏ��O������ꍇ�A���O�̖���������̃Z������P�ڂ̍�ƊJ�n���Ԃ��v�Z����
        With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
            '���O�t���O�̐擪��0�̎�
            If .Cells(2, COLUMN_DATA_REMOVE_SECTION) = 0 Then
                '0�b�ɂ���
                ThisWorkbook.Sheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value = 0

            '���O�t���O�̐擪���P�̎�
            ElseIf .Cells(2, COLUMN_DATA_REMOVE_SECTION) = 1 Then
                '���Z�b�g
                top_jogai_end = 0
                '���O�̖����𒲂ׂ�
                '���O�t���O��1�łȂ��Ȃ�܂Ń��[�v
                Do While .Cells(2 + top_jogai_end, COLUMN_DATA_REMOVE_SECTION) = 1
                    top_jogai_end = top_jogai_end + 1
                Loop

                '���O�̏I�����Ԃ��v�Z���ĊJ�n���Ԃ̂P�s�ڂɓ���
                seconds = top_jogai_end / fps '�����ɕϊ��������b������͂��Ă�������

                format_time = timeConvert(seconds)

                ThisWorkbook.Sheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value = format_time
            End If
        End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        '���������ƕ����Ɋւ��鏈��
        For i = 0 To GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK - GH_HYOUKA_SHEET_ROW_POSESTART - 1 + expand_no
            '��ƊJ�n���Ԃ���Ȃ�
            If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) Then
                '��Ɩ��A��ƏI�����ԁA��Ǝ��ԁA����A���Ȃ��A�G�Ȃ�����ɂ���
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NAME).MergeArea.ClearContents '�Z�����������邽��
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).MergeArea.ClearContents '�Z�����������邽��
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).MergeArea.ClearContents '�Z�����������邽��
                'NG���Ԃ���ɂ���
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME).ClearContents
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME).ClearContents
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME).ClearContents
                '�v�f��Ƃ̔������ɂ���
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT).ClearContents

            '��ƊJ�n���Ԃ����͂���Ă���Ȃ�
            Else
                '���������Ɩ��̓���
                '��Ɩ�����Ȃ���͂���
                If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NAME)) Then
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NAME) = "���" & i + 1
                End If

                '���������ƏI�����Ԃ̓���
                '�P��̍s�̍�ƊJ�n���Ԃ���̎�
                If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i + 1, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) Then
                    '����̖����ɏ��O���Ȃ��ꍇ�A�|�C���g�v�Z�V�[�g�ŏI�s�����ƏI�����Ԃ��v�Z���ē��͂���
                    If ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row_num, COLUMN_DATA_REMOVE_SECTION).Value <> 1 Then
                        'max_row_num�ŎZ�o�����ꍇ�A���o����index 0 ���l������Ȃ��ׁA�ŏI�Z���̒l�𒼐ڎQ��
                        seconds = ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row_num, COLUMN_POSE_KEEP_TIME).Value

                        Debug.Print "seconds:", seconds

                        format_time = timeConvert(seconds)

                        .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value = format_time

                    '����̖����ɏ��O������ꍇ�A���O�̐擪�����̃Z������I�����Ԃ��v�Z����
                    ElseIf ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row_num, COLUMN_DATA_REMOVE_SECTION).Value = 1 Then
                        '�J�E���g���Z�b�g
                        bottom_jogai_start = 0
                        'max_row_num�s�ڂ������オ���āA���O�̐擪�ʒu��T��
                        Do While ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row_num - bottom_jogai_start, COLUMN_DATA_REMOVE_SECTION) = 1
                            bottom_jogai_start = bottom_jogai_start + 1
                        Loop

                        '���斖���ɂ��鏜�O�̊J�n���Ԃ��v�Z���ē���
                        '�|�C���g�v�Z�V�[�g�̌��o��1�s�����l���Ɋ܂�-1
                        seconds = (max_row_num - bottom_jogai_start - 1) / fps

                        format_time = timeConvert(seconds)

                        .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value = format_time
                    End If


                '�P��̍s�̍�ƊJ�n���Ԃɒl�����鎞�A���̒l������
                Else
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value _
                        = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i + 1, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value
                End If

                '�s���]���V�[�g�Ōv�Z�������͂��ꂽ�Z���̒l���X�V����
                Call restartUpdate
                Call stopUpdate

                '��ƏI�����Ԃƍ�ƊJ�n���Ԃ����Ǝ��Ԃ��v�Z���ăZ���ɓ���
                '�Z�����������邽��+2���邱�Ƃŕb���Z�����Q��
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).Value = _
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME + 2).Value _
                    - .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME + 2).Value
            End If
        Next

    End With 'With ThisWorkbook.Sheets("�H���]���V�[�g")

    '���Ԃ������l�ɐݒ�
    separate_work_time = 0
    t0 = 0
    '����擪�����O�����Ƃ��ɕ]���̃X�^�[�g��0.0�b�ł͂Ȃ��Ȃ邽�ߕύX
    t = ThisWorkbook.Sheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME + 2).Value

    '�|�C���g�v�Z�V�[�g�̃t���O���J�E���g���āA�e��Ǝp���̎��Ԃ��v�Z����
    For i = 0 To GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK - GH_HYOUKA_SHEET_ROW_POSESTART - 1 + expand_no

        '��ƊJ�n���Ԃ���Ȃ番�������͂��Ȃ�
        If IsEmpty(ThisWorkbook.Sheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) Then

        '��ƊJ�n���Ԃ����͂���Ă���Ȃ番������������
        Else
            separate_work_time = ThisWorkbook.Sheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME + 2).Value
            t0 = t
            t = separate_work_time '��Ǝ��Ԃ�P��œ��͂���ꍇ
            '�b������t���[�����֕ϊ�
            start_frame = t0 * fps
            end_frame = t * fps - 1

            '��������|�C���g�v�Z�V�[�g�̃t���O���J�E���g
            With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

                '�J�E���^�[�����Z�b�g
                Kobushiage_flag = 0
                koshimage_flag = 0
                hizamage_flag = 0

                'start_frame�t���[��(t0�b) ���� end_frame�t���[��(t�b) �܂ł̏���
                If start_frame < end_frame Then

                    Debug.Print "start_end:", start_frame, ":", end_frame
                    '�������ɏ��O�t���[���J�E���g��������
                    removeFrames = 0

                    For j = start_frame To end_frame

                        '����t���O���J�E���g
                        data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_KOBUSHIAGE).Value
                        If data_no = 1 Then
                            Kobushiage_flag = Kobushiage_flag + 1
                        End If

                        '���Ȃ��t���O���J�E���g
                        data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_KOSHIMAGE).Value
                        If data_no = 1 Then
                            koshimage_flag = koshimage_flag + 1
                        End If

                        '�G�Ȃ��t���O���J�E���g
                        data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_HIZAMAGE).Value
                        If data_no = 1 Then
                            hizamage_flag = hizamage_flag + 1
                        End If

                        '���O��Ԃ��J�E���g
                        data_no = .Cells(2 + j, COLUMN_DATA_REMOVE_SECTION).Value
                        If data_no = 1 Then
                            removeFrames = removeFrames + 1
                        End If
                    Next

                    '��Ǝ��ԍ��v�l���Z�o
                    workFrames = (end_frame + 1 - start_frame) - removeFrames
                    ThisWorkbook.Sheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).Value = workFrames / fps

                Else
                    'start_frame��end_frame���傫���ꍇ�́A��Ǝ��Ԃ�0�ɂ���
                    ThisWorkbook.Sheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).Value = 0
                End If
            End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

            '��������J�E���g�����t���O�����Ԃɕϊ����āA�H���]���V�[�g�ɓ���
            With ThisWorkbook.Sheets("�H���]���V�[�g")

                '����ɑ΂���ʏ���
                If Kobushiage_flag = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME).Value = Kobushiage_flag / fps
                End If

                '���Ȃ��ɑ΂���ʏ���
                If koshimage_flag = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME).Value = koshimage_flag / fps
                End If

                '�G�Ȃ��ɑ΂���ʏ���
                If hizamage_flag = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME).Value = hizamage_flag / fps
                End If
            End With 'With ThisWorkbook.Sheets("�H���]���V�[�g")

            '�v�f��Ɣ���
            With ThisWorkbook.Sheets("�H���]���V�[�g")
                '�Z���v�Z�̒l���Q�Ƃ��邽�߁A�X�g�b�v����
                Call restartUpdate
                Call stopUpdate

                '��J�x�]���A�C�����x�]����NG��Ǝ��Ԃ�Ǎ�
                NG_time_A = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_NG_TIME_A).Value
                NG_time_B = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_NG_TIME_B).Value

                '���肵�Č��ʂ���������
                '�~�̏ꍇ
                If NG_time_A >= youso_hantei_limit Or NG_time_B >= youso_hantei_limit Then
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT).Value = GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_NG
                '���̏ꍇ
                Else
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT).Value = GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_OK
                End If
            End With 'With ThisWorkbook.Sheets("�H���]���V�[�g")
        End If
    Next

    '�\���E�X�V���I���ɖ߂�
    Call restartUpdate

End Sub


'�p���d�ʓ_�����[�Ŏw�肳�ꂽ�]�����O�A�]���������|�C���g�v�Z�V�[�g�ɔ��f������
'�|�C���g�v�Z�V�[�g�̃t���O���玞�Ԃ��v�Z���āA�p���d�ʓ_�����[�ɓ]�L����
'�P��ڂ�Python�v���O��������l�����炤
'�X�V�{�^���������ꂽ�Ƃ��̓|�C���g�v�Z�V�[�g����l��ǂݎ��
' ����1 �F�t���[�����[�g
' �߂�l�F�Ȃ�
Sub fixSheetZensya()

    '�\���E�X�V���I�t�ɂ���
    Call stopUpdate
    Dim fps As Double

    Dim separate_work_time        As Double 't��t0�̍����擾����
    Dim t0                        As Double '1�O��t���ꎞ�ۑ�����
    Dim t                         As Double '��Ǝ���

    Dim i                         As Long
    Dim j                         As Long
    Dim k                         As Long

    Dim max_row_num               As Long

    Dim expand_no                 As Long '�����s���g���p
    Dim data_flag                 As Long '�p���f�_�� �f�[�^���O�i0�j �܂��� �f�[�^�����i1�`10�j�t���O�L���p ���L�ɊY�����Ȃ��ꍇ��-1�����Ďg��

    Dim top_jogai_end             As Long
    Dim bottom_jogai_start        As Long

    Dim Kobushiage_flag           As Long '����� �f�[�^���O�i0�j�܂��� �f�[�^�����i1�j �t���O�L���p ���L�ɊY�����Ȃ��ꍇ��-1�����Ďg��
    Dim koshimage_flag            As Long '���Ȃ��� �f�[�^���O�i0�j�܂��� �f�[�^�����i1�j �t���O�L���p ���L�ɊY�����Ȃ��ꍇ��-1�����Ďg��
    Dim hizamage_flag             As Long '�G�Ȃ��� �f�[�^���O�i0�j�܂��� �f�[�^�����i1�j �t���O�L���p ���L�ɊY�����Ȃ��ꍇ��-1�����Ďg��

    Dim start_frame               As Long
    Dim end_frame                 As Long
    Dim start_array_num           As Long
    Dim end_array_num             As Long


    Dim data_array(15)            As Long '�p���d�ʓ_�P�`�P�O�_�A������ԁA�����ԁA����A���Ȃ��A�G�Ȃ��̎��Ԃ����v���邽�߂Ɏg�p
    Dim data_no                   As Long  'data_array�̔z��ԍ��B1�`10:�p���d�ʓ_ 11:������� 12:������ 13:���� 14:���Ȃ� 15:�G�Ȃ�

    Dim separate_KOBUSHIAGE_missing   As Double '��ƕ�����@���㌇�����
    Dim separate_koshimage_missing    As Double '��ƕ�����@���Ȃ��������
    Dim separate_koshimage_predict    As Double '��ƕ�����@���Ȃ�������
    Dim separate_hizamage_missing     As Double '��ƕ�����@�G�Ȃ��������
    Dim separate_hizamage_predict     As Double '��ƕ�����@�G�Ȃ�������

    Dim seconds          As Double
    Dim hours            As String
    Dim minutes          As String
    Dim remainingSeconds As String
    Dim milliseconds     As String
    Dim format_time      As String

    '�t���[�����[�g���擾
    fps = ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 199)

    '�|�C���g�v�Z�V�[�g�̍ŏI�s���擾
    max_row_num = ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(1, 2).End(xlDown).row

    separate_work_time = 0
    t0 = 0
    t = 0



    '��������ǉ��s�����擾����
    '���̑��i���Ԍv7.5H�j�̃Z���ʒu�̈ړ��ʂ𒲂ׂ�  ���ő�999�s(<979)�ɂ���
    expand_no = 0
    Do While ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(29 + expand_no, 3) <> SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD And expand_no < 979
        expand_no = expand_no + 1
    Loop
    'MsgBox (expand_no)

    '�p���d�ʓ_�����[�ɒl�����
    With ThisWorkbook.Sheets("�p���d�ʓ_�����[")
        '�������珉�񕪐͂̂��߂̏���
        '��ƊJ�n���Ԃ���̏ꍇ�́A0.0�����
        If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) = True Then
            .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = 0
        End If

        '��ƏI�����Ԃ���̏ꍇ�́A�|�C���g�v�Z�V�[�g�ŏI�s����v�Z���ē���
        If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME)) = True Then
            seconds = max_row_num / fps '�����ɕϊ��������b������͂��Ă�������

            hours = seconds \ 3600
            minutes = (seconds Mod 3600) \ 60
            remainingSeconds = seconds Mod 60
            milliseconds = (seconds - Int(seconds)) * 10

            format_time = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00") & "." & Format(milliseconds, "0")

            .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = format_time
        End If

        '�������璠�[�X�V�̂��߂̏���
        '����̐擪�ɏ��O������ꍇ�A���O�̖���������̃Z������P�ڂ̍�ƊJ�n���Ԃ��v�Z����
        With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
            '���O�t���O�̐擪��0�̎�
            If .Cells(2, COLUMN_DATA_REMOVE_SECTION) = 0 Then
                '0�b�ɂ���
                ThisWorkbook.Sheets("�p���d�ʓ_�����[").Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = 0

            '���O�t���O�̐擪���P�̎�
            ElseIf .Cells(2, COLUMN_DATA_REMOVE_SECTION) = 1 Then
                '���Z�b�g
                top_jogai_end = 0
                '���O�̖����𒲂ׂ�
                '���O�t���O��1�łȂ��Ȃ�܂Ń��[�v
                Do While .Cells(2 + top_jogai_end, COLUMN_DATA_REMOVE_SECTION) = 1
                    top_jogai_end = top_jogai_end + 1
                Loop

                '���O�̏I�����Ԃ��v�Z���ĊJ�n���Ԃ̂P�s�ڂɓ���
                seconds = top_jogai_end / fps '�����ɕϊ��������b������͂��Ă�������

                hours = seconds \ 3600
                minutes = (seconds Mod 3600) \ 60
                remainingSeconds = seconds Mod 60
                milliseconds = (seconds - Int(seconds)) * 10

                format_time = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00") & "." & Format(milliseconds, "0")

                ThisWorkbook.Sheets("�p���d�ʓ_�����[").Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value = format_time
            End If
        End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        '���������ƕ����Ɋւ��鏈��
        For i = 0 To SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK - SHIJUTEN_SHEET_ROW_POSESTART_INDEX - 1 + expand_no
            '��ƊJ�n���Ԃ���Ȃ�
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) Then
                '��Ɩ��A��ƏI�����ԁA��Ǝ��ԁA����A���Ȃ��A�G�Ȃ�����ɂ���
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).MergeArea.ClearContents '�Z�����������邽��
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).ClearContents
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).ClearContents

                '�p���f�_�ƂЂ˂����ɂ���
                For j = 0 To 10
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).ClearContents
                Next

                'NG���Ԃ���ɂ���
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME).ClearContents
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME).ClearContents
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME).ClearContents

            '��ƊJ�n���Ԃ����͂���Ă���Ȃ�
            Else
                '���������Ɩ��̓���
                '��Ɩ�����Ȃ���͂���
                If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME)) Then
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME) = "���" & i + 1
                End If

                '���������ƏI�����Ԃ̓���
                '�P��̍s�̍�ƊJ�n���Ԃ���̎�
                If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i + 1, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME)) Then
                    '����̖����ɏ��O���Ȃ��ꍇ�A�|�C���g�v�Z�V�[�g�ŏI�s�����ƏI�����Ԃ��v�Z���ē��͂���
                    If ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row_num, COLUMN_DATA_REMOVE_SECTION).Value <> 1 Then
                        seconds = max_row_num / fps '�����ɕϊ��������b������͂��Ă�������

                        hours = seconds \ 3600
                        minutes = (seconds Mod 3600) \ 60
                        remainingSeconds = seconds Mod 60
                        milliseconds = (seconds - Int(seconds)) * 10

                        format_time = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00") & "." & Format(milliseconds, "0")

                        .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = format_time

                    '����̖����ɏ��O������ꍇ�A���O�̐擪�����̃Z������I�����Ԃ��v�Z����
                    ElseIf ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row_num, COLUMN_DATA_REMOVE_SECTION).Value = 1 Then
                        '�J�E���g���Z�b�g
                        bottom_jogai_start = 0
                        'max_row_num�s�ڂ������オ���āA���O�̐擪�ʒu��T��
                        Do While ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row_num - bottom_jogai_start, COLUMN_DATA_REMOVE_SECTION) = 1
                            bottom_jogai_start = bottom_jogai_start + 1
                        Loop

                        '���斖���ɂ��鏜�O�̊J�n���Ԃ��v�Z���ē���
                        seconds = (max_row_num - bottom_jogai_start) / fps '�����ɕϊ��������b������͂��Ă�������

                        hours = seconds \ 3600
                        minutes = (seconds Mod 3600) \ 60
                        remainingSeconds = seconds Mod 60
                        milliseconds = (seconds - Int(seconds)) * 10

                        format_time = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00") & "." & Format(milliseconds, "0")

                        .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value = format_time
                    End If


                '�P��̍s�̍�ƊJ�n���Ԃɒl�����鎞�A���̒l������
                Else
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME).Value _
                        = .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i + 1, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME).Value
                End If

                '�s���]���V�[�g�Ōv�Z�������͂��ꂽ�Z���̒l���X�V����
                Call restartUpdate
                Call stopUpdate

                '��ƏI�����Ԃƍ�ƊJ�n���Ԃ����Ǝ��Ԃ��v�Z���ăZ���ɓ���
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value = _
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKEND_TIME + 1).Value _
                    - .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME + 1).Value
            End If
        Next

    End With 'With ThisWorkbook.Sheets("�p���d�ʓ_�����[")

    '���No.�̑��
    For i = 0 To 19 + expand_no
        ThisWorkbook.Sheets("�p���d�ʓ_�����[").Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NUMBER).Value = i + 1
    Next

    For i = 0 To 19 + expand_no
        '���F�Z���̎��Ԃ�ǂݎ��
        '�it0�b�`t�b�܂ł̎p�������߂�j
        With ThisWorkbook.Sheets("�p���d�ʓ_�����[")
            separate_work_time = .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME).Value
            t0 = t
            t = t + separate_work_time '��Ǝ��Ԃ�P��œ��͂���ꍇ


        End With 'With ThisWorkbook.Sheets("�p���d�ʓ_�����[")

        '�b������t���[�����֕ϊ�
        start_frame = t0 * fps
        end_frame = t * fps

        '2�Z�b�g�ڈȍ~�ŁA�O���end_frame�ƍ����start_frame���d�Ȃ�̂�h�~����
        If start_frame > 0 Then
            start_frame = start_frame + 1
        End If

        '�f�o�b�O�p
        'MsgBox ("i=" & i & " / start_frame = " & start_frame & "(" & t0 & ")" & ",  end_frame = " & end_frame & "(" & t & ")")

        '�p���v�f���Ԃ�����ϐ��̏�����
        For j = 1 To 15
            data_array(j) = 0
        Next

        '���������Ԃ��J�E���g����ϐ��̏�����
        separate_KOBUSHIAGE_missing = 0
        separate_koshimage_missing = 0
        separate_koshimage_predict = 0
        separate_hizamage_missing = 0
        separate_hizamage_predict = 0

        With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

            'start_frame�t���[��(t0�b) ���� end_frame�t���[��(t�b) �܂ł̏���
            If start_frame < end_frame Then
                For j = start_frame To end_frame

                    '�p���f�_�̒l���J�E���g
                    data_no = .Cells(2 + j, COLUMN_DATA_RESULT_FIX).Value
                    data_array(data_no) = data_array(data_no) + 1

                    '�f�[�^������Ԃ��J�E���g
                    data_no = .Cells(2 + j, COLUMN_DATA_MISSING_SECTION).Value
                    If data_no = 1 Then
                        data_array(11) = data_array(11) + 1
                    End If

                    '�f�[�^�����Ԃ��J�E���g
                    data_no = .Cells(2 + j, COLUMN_DATA_PREDICT_SECTION).Value
                    If data_no >= 1 And data_no <= 10 Then
                        data_array(12) = data_array(12) + 1
                    End If

                    '����t���O���J�E���g
                    data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_KOBUSHIAGE).Value
                    If data_no = 1 Then
                        data_array(13) = data_array(13) + 1
                    End If

                    '���Ȃ��t���O���J�E���g
                    data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_KOSHIMAGE).Value
                    If data_no = 1 Then
                        data_array(14) = data_array(14) + 1
                    End If

                    '�G�Ȃ��t���O���J�E���g
                    data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_HIZAMAGE).Value
                    If data_no = 1 Then
                        data_array(15) = data_array(15) + 1
                    End If

                    '���㌇�����J�E���g
                    If .Cells(2 + j, COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST).Value = 1 Then
                        separate_KOBUSHIAGE_missing = separate_KOBUSHIAGE_missing + 1
                    End If

                    '���Ȃ��������J�E���g
                    If .Cells(2 + j, COLUMN_DATA_KOSHIMAGE_MISSING_SECTION).Value = 1 Then
                        separate_koshimage_missing = separate_koshimage_missing + 1
                    End If

                    '���Ȃ�������J�E���g
                    If .Cells(2 + j, COLUMN_DATA_KOSHIMAGE_PREDICT_SECTION).Value = 1 Then
                        separate_koshimage_predict = separate_koshimage_predict + 1
                    End If

                    '�G�Ȃ��������J�E���g
                    If .Cells(2 + j, COLUMN_DATA_HIZAMAGE_MISSING_SECTION).Value = 1 Then
                        separate_hizamage_missing = separate_hizamage_missing + 1
                    End If

                    '�G�Ȃ�������J�E���g
                    If .Cells(2 + j, COLUMN_DATA_HIZAMAGE_PREDICT_SECTION).Value = 1 Then
                        separate_hizamage_predict = separate_hizamage_predict + 1
                    End If

                    '�|�C���g�v�Z�V�[�g�̃L���v�V������ցA�p���d�ʓ_�����[�̍��No,�ƍ�Ɩ���ǂݎ��A
                    '"���No.xxx_��Ɩ� "�Ƃ��ē���Ă���
                    .Cells(2 + j, COLUMN_CAPTION_WORK_NAME).Value = _
                    "���No." & _
                    ThisWorkbook.Sheets("�p���d�ʓ_�����[").Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NUMBER).Value & _
                    " " & _
                    ThisWorkbook.Sheets("�p���d�ʓ_�����[").Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_NAME).Value & _
                    " "

                Next
            End If
        End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        With ThisWorkbook.Sheets("�p���d�ʓ_�����[")

            '�p���v�f10�`1�ɑ΂���ʏ���
            For j = 0 To 9
                If data_array(10 - j) = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX + j).Value = data_array(SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX - j) / fps
                End If
            Next

            '���i���o���s�ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION).Value = ""
            Else
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION).Value = data_array(11) / fps
            End If

            '�p���f�_�i�����ԁj�ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION).Value = ""
            Else
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION).Value = data_array(12) / fps
            End If

            '����ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME).Value = ""
            Else
                If data_array(13) = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    '���㎞�Ԃƕ␳�l�����v���ăZ���ɓ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME).Value = data_array(13) / fps
                End If
            End If

            '���Ȃ��ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME).Value = ""
            Else

                If data_array(14) = 0 Then
                '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME).Value = ""
                Else
                '�p���v�f���ԁi�t���[�����j������Α������
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME).Value = data_array(14) / fps
                End If
            End If

            '�G�Ȃ��ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME).Value = ""
            Else

                If data_array(15) = 0 Then
                '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME).Value = ""
                Else
                '�p���v�f���ԁi�t���[�����j������Α������
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME).Value = data_array(15) / fps
                End If
            End If

            '���㌇���ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_MISSING_TIME).Value = ""
            Else
                If separate_KOBUSHIAGE_missing = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_MISSING_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_MISSING_TIME).Value = separate_KOBUSHIAGE_missing / fps
                End If
            End If

            '���Ȃ������ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME).Value = ""
            Else
                If separate_koshimage_missing = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME).Value = separate_koshimage_missing / fps
                End If
            End If

            '���Ȃ�����ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME).Value = ""
            Else
                If separate_koshimage_predict = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME).Value = separate_koshimage_predict / fps
                End If
            End If

            '�G�Ȃ������ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME).Value = ""
            Else
                If separate_hizamage_missing = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME).Value = separate_hizamage_missing / fps
                End If
            End If

            '�G�Ȃ�����ɑ΂���ʏ���
            If IsEmpty(.Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_WORK_TIME)) Then
                .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME).Value = ""
            Else
                If separate_hizamage_predict = 0 Then
                    '�p���v�f���ԁi�t���[�����j��0�̂Ƃ��́A�󔒃Z���ɂ���
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME).Value = ""
                Else
                    '�p���v�f���ԁi�t���[�����j������Α������
                    .Cells(SHIJUTEN_SHEET_ROW_POSESTART_INDEX + i, SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME).Value = separate_hizamage_predict / fps
                End If
            End If
        End With 'With ThisWorkbook.Sheets("�p���d�ʓ_�����[")
    Next

    '�\���E�X�V���I���ɖ߂�
    Call restartUpdate

End Sub


'�����t�@�C���o��
'����1 �F���於
'�߂�l�F�Ȃ�
Function outputCaption(movieName As String)
    Dim i                           As Long
    Dim max_row_num                 As Long

    '����̏c�����r���ĕ����T�C�Y�������邽�߁A���E�����ǂ�����g�p����
    Dim video_width                 As Long '���͓���̕� ��3D�|�[�Y���������ꂽ���ł͂Ȃ����ߒ���
    Dim video_height                As Long '���͓���̍���

    '��coef��coefficient�i�W���A���j�̗��L
    Dim track1_coef_font_size1      As Long '�����g���b�N1�p  ��i�̃T�C�Y�����p�W��
    Dim track1_coef_font_size2      As Long '�����g���b�N1�p  ���i�̃T�C�Y�����p�W��
    Dim track1_font_size1           As Long '�����g���b�N1�p  ��i�̃T�C�Y
    Dim track1_font_size2           As Long '�����g���b�N1�p  ���i�̃T�C�Y

    Dim track2_coef_font_size1      As Long '�����g���b�N2�p �P�i�ڂ̃T�C�Y�����p�W��
    Dim track2_coef_font_size2      As Long '�����g���b�N2�p �Q�i�ڂ̃T�C�Y�����p�W��
    Dim track2_coef_font_size3      As Long '�����g���b�N2�p �R�i�ڂ̃T�C�Y�����p�W��
    Dim track2_font_size1           As Long '�����g���b�N2�p �P�i�ڂ̃T�C�Y
    Dim track2_font_size2           As Long '�����g���b�N2�p �Q�i�ڂ̃T�C�Y
    Dim track2_font_size3           As Long '�����g���b�N2�p �R�i�ڂ̃T�C�Y

    Dim CaptionName0                As String  '�����g���b�N1�p ��i���� ��Ɩ�          �̎���������
    Dim CaptionName1                As String  '�����g���b�N1�p ��i�E�� �уO���t�̃f�[�^�i�M���x�j�̎���������
    Dim CaptionName2(10)            As String  '�����g���b�N1�p ���i �]�����O(�Y����0)+�p���f�_1�`10(�Y����1�`10)�̎���������
    Dim CaptionNo2                  As Long 'CaptionName2(10)�ɃA�N�Z�X����ۂ̓Y�����i�[�p�ϐ�

    Dim CaptionName2Kobushiage      As String '�����g���b�N2�p �Q�i�� �����Ԃ̎���������
    Dim CaptionName2Koshimage       As String '�����g���b�N2�p �Q�i�� ���Ȃ��f�[�^��Ԃ̎���������
    Dim CaptionName2Hizamage        As String '�����g���b�N2�p �Q�i�� �G�Ȃ��f�[�^��Ԃ̎���������

    Dim CaptionName3Kobushiage      As String '�����g���b�N2�p �R�i�� ����̎���������
    Dim CaptionName3Koshimage       As String '�����g���b�N2�p �R�i�� ���Ȃ��̎���������
    Dim CaptionName3Hizamage        As String '�����g���b�N2�p �R�i�� �G�Ȃ��̎���������

    Dim ColorName1                  As String '�����g���b�N1�p ��i�E���i�M���x �j�̐F
    Dim ColorName2                  As String '�����g���b�N1�p ���i  �i�p���f�_�j�̐F
    Dim ColorName2Kobushiage        As String '�����g���b�N2�p �Q�i�� �i������ �j�̐F
    Dim ColorName2Koshimage         As String '�����g���b�N2�p �Q�i�� �i���Ȃ��f�[�^��� �j�̐F
    Dim ColorName2Hizamage          As String '�����g���b�N2�p �Q�i�� �i�G�Ȃ��f�[�^��� �j�̐F
    Dim ColorName3Kobushiage        As String '�����g���b�N2�p �R�i�� �i���� �j�̐F
    Dim ColorName3Koshimage         As String '�����g���b�N2�p �R�i�� �i���Ȃ� �j�̐F
    Dim ColorName3Hizamage          As String '�����g���b�N2�p �R�i�� �i�G�Ȃ� �j�̐F

    Dim Track1OutputString1         As String '�����g���b�N1�p�F��i������
    Dim Track1OutputString2         As String '�����g���b�N1�p�F���i������

    Dim Track2OutputString1         As String '�����g���b�N2�p�F1�i�ڕ�����
    Dim Track2OutputString2         As String '�����g���b�N2�p�F2�i�ڕ�����
    Dim Track2OutputString3         As String '�����g���b�N2�p�F3�i�ڕ�����

    Dim Track1FileName              As String '�����g���b�N1�p�̃t�@�C����
    Dim Track2FileName              As String '�����g���b�N2�p�̃t�@�C����


    '�\���E�X�V���I�t�ɂ���
    Call stopUpdate

    '����̏c���T�C�Y���擾
    video_width = ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 198)
    video_height = ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 197) '����̏c������̂��߂ɍ������擾

    '����̏c���ɂ���ČW����ύX����
    '���悪�c�̎�
    If video_width < video_height Then
        track1_coef_font_size1 = TRACK1_TATE_UPPER_COEF  '���悪�c�̂Ƃ��̃g���b�N1�p�F��i
        track1_coef_font_size2 = TRACK1_TATE_LOWER_COEF
        track2_coef_font_size1 = TRACK2_TATE_1ST_COEF    '�g���b�N2�p�F�P�i��
        track2_coef_font_size2 = TRACK2_TATE_2ND_COEF    '�g���b�N2�p�F�Q�i��
        track2_coef_font_size3 = TRACK2_TATE_3RD_COEF    '�g���b�N2�p�F�R�i��
    '���悪���̎�
    Else
        track1_coef_font_size1 = TRACK1_YOKO_UPPER_COEF  '���悪�c�̂Ƃ��̃g���b�N1�p�F��i
        track1_coef_font_size2 = TRACK1_YOKO_LOWER_COEF
        track2_coef_font_size1 = TRACK2_YOKO_1ST_COEF    '�g���b�N2�p�F�P�i��
        track2_coef_font_size2 = TRACK2_YOKO_2ND_COEF    '�g���b�N2�p�F�Q�i��
        track2_coef_font_size3 = TRACK2_YOKO_3RD_COEF    '�g���b�N2�p�F�R�i��
    End If

    '�t�H���g�T�C�Y��ݒ�
    track1_font_size1 = video_width / track1_coef_font_size1 '����̏cor���ɂ���ĕ����ύX���邱�ƂŁA�����T�C�Y���ς��
    track1_font_size2 = video_width / track1_coef_font_size2
    track2_font_size1 = video_width / track2_coef_font_size1
    track2_font_size2 = video_width / track2_coef_font_size2
    track2_font_size3 = video_width / track2_coef_font_size3

    '�e�p���̖��O�Ə����̓ǂݏo��
    'Min��Max�������I�łȂ��̂Œ���
    With ThisWorkbook.Worksheets("�����ݒ�V�[�g")
        CaptionName2(10) = .Cells(6, 2)
        CaptionName2(9) = .Cells(24, 2)
        CaptionName2(8) = .Cells(42, 2)
        CaptionName2(7) = .Cells(60, 2)
        CaptionName2(6) = .Cells(78, 2)
        CaptionName2(5) = .Cells(96, 2)
        CaptionName2(4) = .Cells(114, 2)
        CaptionName2(3) = .Cells(132, 2)
        CaptionName2(2) = .Cells(150, 2)
        CaptionName2(1) = .Cells(168, 2)
        CaptionName3Kobushiage = .Cells(192, 2)
        CaptionName2Koshimage = .Cells(210, 2)
        CaptionName2Hizamage = .Cells(228, 2)
    End With 'With ThisWorkbook.Worksheets("�����ݒ�V�[�g")

    '�]�����O�p
    CaptionName2(0) = "0-�p���]���Ȃ�" '���i�̃L���v�V��������\�����Ȃ�

    '�e�t�@�C������ݒ�
'    '�����g���b�N����ւ��Ĥ���̑��������g���b�N1�ɂȂ鎞
'    If ThisWorkbook.Worksheets("�p���d�ʓ_�����[").CheckBox1 = True Then
'        Track1FileName = ActiveWorkbook.Path & "\" & movieName & CAPTION_TRACK2_FILE_NAME_SOEJI & ".srt"
'        Track2FileName = ActiveWorkbook.Path & "\" & movieName & ".srt"
'    '�ʏ폈�� �p���f�_�������g���b�N1�ɂȂ鎞
'    Else
'        Track1FileName = ActiveWorkbook.Path & "\" & movieName & ".srt"
'        Track2FileName = ActiveWorkbook.Path & "\" & movieName & CAPTION_TRACK2_FILE_NAME_SOEJI & ".srt"
'    End If
    Track2FileName = ActiveWorkbook.Path & "\" & movieName & ".srt"


    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

        '�t�@�C�����J��
        'Open Track1FileName For Output As #2
        Open Track2FileName For Output As #2

        '��������s�����擾�i3��ڂ̍ŏI�Z���j
        max_row_num = .Cells(1, 3).End(xlDown).row

        '�t�@�C���o��
        For i = 2 To max_row_num

            '�|�C���g�v�Z�V�[�g�̃L���v�V��������A�p���d�ʓ_�����[�̍�Ɩ����ɓǂݎ���Ă���
            CaptionName0 = .Cells(i, COLUMN_CAPTION_WORK_NAME).Value

            '�J�����̃V�X�e���͕��������d�l�̂��߁A�G���[�h�~�̂��߁A�b��Ŏ����Q�͋�̃t�@�C�����쐬����
            '���ԗp�̃V�X�e�����쐬����ۂɂ͎����t�@�C�����P�ɂ���

'            '////////////////////////////////////////
'            '// �����g���b�N1�p�̏��� ��������
'            '//
'
'            '�f�[�^��Ԃ̕`��F�A�L���v�V��������ݒ肷��
'            '���͂��߂ɕ]�����O�A�f�[�^������ԁA�f�[�^�s�ǋ�Ԃ̏��ɔ��肷��i�d���r�b�gON���A�����\���̗D��x���������j
'            ' ���̂Ƃ���Ō��fillData�ő���r�b�g�͓����ɗ��d�l�B
'            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
'                CaptionName1 = CAPTION_DATA_REMOVE_SECTION
'                ColorName1 = COLOR_DATA_REMOVE_SECTION
'            ElseIf .Cells(i, COLUMN_DATA_FORCED_SECTION).Value > 0 Then
'                CaptionName1 = CAPTION_DATA_FORCED_SECTION
'                ColorName1 = COLOR_DATA_FORCED_SECTION
'            ElseIf .Cells(i, COLUMN_DATA_MISSING_SECTION).Value > 0 Then
'                CaptionName1 = CAPTION_DATA_MISSING_SECTION
'                ColorName1 = COLOR_DATA_MISSING_SECTION
'            ElseIf .Cells(i, COLUMN_DATA_MEASURE_SECTION).Value > 0 Then
'                CaptionName1 = CAPTION_DATA_MEASURE_SECTION
'                ColorName1 = COLOR_DATA_MEASURE_SECTION
'            ElseIf .Cells(i, COLUMN_DATA_PREDICT_SECTION).Value > 0 Then
'                CaptionName1 = CAPTION_DATA_PREDICT_SECTION
'                ColorName1 = COLOR_DATA_PREDICT_SECTION
'            End If
'
'            '�p���f�_�̕`��F�A�L���v�V��������ݒ肷��
'            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
'                '�]�����O�̂Ƃ�
'                CaptionNo2 = 0
'                ColorName2 = COLOR_DATA_REMOVE_SECTION
'            Else
'                '�ʏ펞
'                CaptionNo2 = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value
'                If .Cells(i, COLUMN_DATA_RESULT_GREEN).Value > 0 Then
'                    ColorName2 = COLOR_DATA_RESULT_GREEN
'                ElseIf .Cells(i, COLUMN_DATA_RESULT_YELLOW).Value > 0 Then
'                    ColorName2 = COLOR_DATA_RESULT_YELLOW
'                ElseIf .Cells(i, COLUMN_DATA_RESULT_RED).Value > 0 Then
'                    ColorName2 = COLOR_DATA_RESULT_RED
'                End If
'            End If
'
'            '����������𐶐�
'            Track1OutputString1 = _
'                "<font size=""" & track1_font_size1 & """ color =" & "#ffffff" & ">" & CaptionName0 & "</font>" & _
'                "<font size=""" & track1_font_size1 & """ color =" & ColorName1 & ">" & CaptionName1 & "</font>"
'            Track1OutputString2 = _
'                "<font size=""" & track1_font_size2 & """ color =" & ColorName2 & ">" & CaptionName2(CaptionNo2) & _
'                "</font>"
'
'            '������������|�C���g�v�Z�V�[�g�ɏo��
'            '�f�o�b�O�p�i���i�͎g��Ȃ��j
'            '.Cells(i, COLUMN_CAPTION_TRACK1).Value = Track1OutputString1 & Track1OutputString2
'
'            '������������e�L�X�g�t�@�C���ɏ����o������
'            Print #1, " " & i - 1 '�����̗����ɔ��p�X�y�[�X������B�����g���b�N2�Ƌ�ʂ��邽��
'            Print #1, .Cells(i, COLUMN_ROUGH_TIME).Value&; " --> " & .Cells(i + 1, COLUMN_ROUGH_TIME).Value '�������o��
'
'            Print #1, Replace(Track1OutputString1, vbLf, vbCrLf) '���s�R�[�h��u�������A�L���v�V�����o��
'            Print #1, Replace(Track1OutputString2, vbLf, vbCrLf) '���s�R�[�h��u�������A�L���v�V�����o��
'
'            Print #1, ""
'            Print #1, ""
'
'            '//
'            '// �����g���b�N1�p�̏��� �����܂�
'            '////////////////////////////////////////



            '////////////////////////////////////////
            '// �����g���b�N2�p�̏��� ��������
            '//

            '�f�[�^��Ԃ̕`��F�A�L���v�V��������ݒ肷��
            '���͂��߂ɕ]�����O�A�f�[�^������ԁA�f�[�^�s�ǋ�Ԃ̏��ɔ��肷��i�d���r�b�gON���A�����\���̗D��x���������j
            ' ���̂Ƃ���Ō��fillData�ő���r�b�g�͓����ɗ��d�l�B

           '1�i�ڂ̕`��F�A�L���v�V�������͎����g���b�N1�p��i�p�̍��No.�ƍ�Ɩ��𗬗p���邽�߁A�����ł͏����Ȃ�
           '2�i�ڂ̕`��F�A�L���v�V��������ݒ肷��i�f�[�^�̐M�����j
            '�������猝��
            '���O
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_REMOVE_SECTION
                ColorName2Kobushiage = COLOR_DATA_REMOVE_SECTION
            '����
            ElseIf .Cells(i, COLUMN_KOBUSHIAGE_FORCED_SECTION).Value > 0 Then
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_FORCED_SECTION
                ColorName2Kobushiage = COLOR_DATA_FORCED_SECTION
            '����
            ElseIf .Cells(i, COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST).Value > 0 Then
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_MISSING_SECTION
                ColorName2Kobushiage = COLOR_DATA_MISSING_SECTION
            '����
            Else
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_MEASURE_SECTION
                ColorName2Kobushiage = COLOR_DATA_MEASURE_SECTION
            End If

            '�������獘�Ȃ�
            '���O
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_REMOVE_SECTION
                ColorName2Koshimage = COLOR_DATA_REMOVE_SECTION
            '����
            ElseIf .Cells(i, COLUMN_KOSHIMAGE_FORCED_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_FORCED_SECTION
                ColorName2Koshimage = COLOR_DATA_FORCED_SECTION
            '����
            ElseIf .Cells(i, COLUMN_DATA_KOSHIMAGE_MISSING_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_MISSING_SECTION
                ColorName2Koshimage = COLOR_DATA_MISSING_SECTION
            '����
            ElseIf .Cells(i, COLUMN_DATA_KOSHIMAGE_MEASURE_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_MEASURE_SECTION
                ColorName2Koshimage = COLOR_DATA_MEASURE_SECTION
            '����
            ElseIf .Cells(i, COLUMN_DATA_KOSHIMAGE_PREDICT_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_PREDICT_SECTION
                ColorName2Koshimage = COLOR_DATA_PREDICT_SECTION
            End If

            '��������G�Ȃ�
            '���O
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_REMOVE_SECTION
                ColorName2Hizamage = COLOR_DATA_REMOVE_SECTION
            '����
            ElseIf .Cells(i, COLUMN_HIZAMAGE_FORCED_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_FORCED_SECTION
                ColorName2Hizamage = COLOR_DATA_FORCED_SECTION
            '����
            ElseIf .Cells(i, COLUMN_DATA_HIZAMAGE_MISSING_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_MISSING_SECTION
                ColorName2Hizamage = COLOR_DATA_MISSING_SECTION
            '����
            ElseIf .Cells(i, COLUMN_DATA_HIZAMAGE_MEASURE_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_MEASURE_SECTION
                ColorName2Hizamage = COLOR_DATA_MEASURE_SECTION
            '����
            ElseIf .Cells(i, COLUMN_DATA_HIZAMAGE_PREDICT_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_PREDICT_SECTION
                ColorName2Hizamage = COLOR_DATA_PREDICT_SECTION
            End If



            '�R�i�ڂ̕`��F�A�L���v�V��������ݒ肷��
            '����
            If .Cells(i, COLUMN_KOBUSHIAGE_RESULT).Value > 0 Then
                CaptionName3Kobushiage = "<b>" & CAPTION_A_RESULT_NAME1 & "</b>"
                ColorName3Kobushiage = COLOR_DATA_RESULT_RED
            Else
                CaptionName3Kobushiage = "<b>" & CAPTION_A_RESULT_NAME1 & "</b>"
                ColorName3Kobushiage = COLOR_DATA_RESULT_GLAY
            End If

            '���Ȃ�
            If .Cells(i, COLUMN_KOSHIMAGE_RESULT).Value > 0 Then
                CaptionName3Koshimage = "<b>" & CAPTION_B_RESULT_NAME1 & "</b>"
                ColorName3Koshimage = COLOR_DATA_RESULT_RED
            Else
                CaptionName3Koshimage = "<b>" & CAPTION_B_RESULT_NAME1 & "</b>"
                ColorName3Koshimage = COLOR_DATA_RESULT_GLAY
            End If

            '�G�Ȃ�
            If .Cells(i, COLUMN_HIZAMAGE_RESULT).Value > 0 Then
                CaptionName3Hizamage = "<b>" & CAPTION_C_RESULT_NAME1 & "</b>"
                ColorName3Hizamage = COLOR_DATA_RESULT_RED
            Else
                CaptionName3Hizamage = "<b>" & CAPTION_C_RESULT_NAME1 & "</b>"
                ColorName3Hizamage = COLOR_DATA_RESULT_GLAY
            End If

            '���������������
            Track2OutputString1 = _
                "<font size=""" & track2_font_size1 & """ color =" & "#ffffff" & ">" & CaptionName0 & "</font>"

            Track2OutputString2 = _
                "<font size=""" & track2_font_size2 & """ color =" & ColorName2Kobushiage & ">" & CaptionName2Kobushiage & "</font>" & _
                "<font size=""" & track2_font_size2 & """ color =" & ColorName2Kobushiage & ">" & "          " & "</font>" & _
                "<font size=""" & track2_font_size2 & """ color =" & ColorName2Koshimage & ">" & CaptionName2Koshimage & "</font>" & _
                "<font size=""" & track2_font_size2 & """ color =" & ColorName2Kobushiage & ">" & "          " & "</font>" & _
                "<font size=""" & track2_font_size2 & """ color =" & ColorName2Hizamage & ">" & CaptionName2Hizamage & "</font>"

            Track2OutputString3 = _
                "<font size=""" & track2_font_size3 & """ color =" & ColorName3Kobushiage & ">" & CaptionName3Kobushiage & "</font>" & _
                "<font size=""" & track2_font_size3 & """ color =" & ColorName3Kobushiage & ">" & "    " & "</font>" & _
                "<font size=""" & track2_font_size3 & """ color =" & ColorName3Koshimage & ">" & CaptionName3Koshimage & "</font>" & _
                "<font size=""" & track2_font_size3 & """ color =" & ColorName3Kobushiage & ">" & "    " & "</font>" & _
                "<font size=""" & track2_font_size3 & """ color =" & ColorName3Hizamage & ">" & CaptionName3Hizamage & "</font>"

            '������������|�C���g�v�Z�V�[�g�ɏo��
            '�f�o�b�O�p�i���i�͎g��Ȃ��j
            '.Cells(i, COLUMN_CAPTION_TRACK2).Value = Track1OutputString1 & Track1OutputString2

            '�e�L�X�g�t�@�C���ɂ��̑�����������������o��
            Print #2, "  " & i - 1 '�����̑O�ɔ��p�X�y�[�X��2�����B�����g���b�N1�Ƌ�ʂ��邽��
            Print #2, .Cells(i, COLUMN_ROUGH_TIME).Value&; " --> " & .Cells(i + 1, COLUMN_ROUGH_TIME).Value '�������o��

            Print #2, Replace(Track2OutputString1, vbLf, vbCrLf) '���s�R�[�h��u�������A�L���v�V�����o��
            Print #2, Replace(Track2OutputString2, vbLf, vbCrLf) '���s�R�[�h��u�������A�L���v�V�����o��
            Print #2, Replace(Track2OutputString3, vbLf, vbCrLf) '���s�R�[�h��u�������A�L���v�V�����o��

            Print #2, ""
            Print #2, ""

            '//
            '// �����g���b�N2�p�̏��� �����܂�
            '////////////////////////////////////////

            '�|�C���g�v�Z�V�[�g�̎��������� ���No. - ��Ɩ����N���A
            .Cells(i, COLUMN_CAPTION_WORK_NAME).clear


            '�f�o�b�O���A���肳��Ȃ�������������悤�ɐF�������Z�b�g���Ă���
            ColorName1 = "#ffffff"
            ColorName2 = "#ffffff"

        Next

        '�t�@�C�������
        Close #1
        Close #2


    End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

    '�\���E�X�V���I���ɖ߂�
    Call restartUpdate

End Function


'���[�X�V�{�^���������ꂽ���̏���
' ����  �F�Ȃ�
' �߂�l�F�Ȃ�
Function ClickUpdateDataCore()
    Dim tstart_click As Double
    Dim dotPoint     As String
    Dim workbookName As String
    Dim fps          As Double

    tstart_click = Timer
    fps = ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 199)

    '�m�C�Y����
    Call removeCaptionNoise(fps)

    '��ƕ����A���ԑ���
    Call fixSheetJisya
    'Call fixSheetZensya


    dotPoint = InStrRev(ActiveWorkbook.Name, ".")
    workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)

    Call outputCaption(workbookName)
    Debug.Print " �X�V����" & Format$(Timer - tstart_click, "0.00") & " sec."

End Function


'���[�X�V�{�^���������ꂽ���̏���
' �����P�F�Ȃ�
' �߂�l�F�Ȃ�
Sub ClickUpdateData()
    Call ClickUpdateDataCore
End Sub

Sub ClickJisyaLimitChangeUpdateData()

    Dim response As VbMsgBoxResult

    ' ���b�Z�[�W�{�b�N�X�Ń��[�U�[�Ɋm�F
    response = MsgBox("���E�G�p�x�̂������l��ύX���čĕ]�����܂��B�p���]���C���V�[�g�̕ҏW���e�����Z�b�g����܂�����낵���ł����H", vbOKCancel + vbQuestion, "�m�F")

    ' ���[�U�[�̑I���ɉ����ď���
    If response = vbOK Then
        ' OK�{�^���������ꂽ�ꍇ�A���b�Z�[�W��\��
        Dim tstart_click As Double
        Dim dotPoint     As String
        Dim workbookName As String


        tstart_click = Timer

        dotPoint = InStrRev(ActiveWorkbook.Name, ".")
        workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)


        '�p������
        Call makeGraphJisya
        '��ƕ����A���ԑ���
        Call fixSheetJisya
        '�C���V�[�g�̍X�V
        Call Module3.paintAll
        '�����o��
        Call outputCaption(workbookName)
        '�X�V���ԏo��
        Debug.Print " �X�V����" & Format$(Timer - tstart_click, "0.00") & " sec."
        Sheets("�H���]���V�[�g").Activate
    Else
        ' �L�����Z���{�^���������ꂽ�ꍇ�A�������Ȃ�
    End If
End Sub




' �T�v : �֐ߊp�x�A3d�f�[�^��csv���R�s�[�\��t������
' �Ăь��̃V�[�g : �}�N���e�X�g
' �⑫ : �{�t�@�C���Ɠ����f�B���N�g����csv�t�@�C����u���Ă���
' ����1 �F�t���[�����[�g
' ����2 �F���扡���̒l
' ����3 �Fcsv�t�@�C����
' ����4 �F����c�̒l ����̌����ɂ���Ď��������T�C�Y�𒲐����邽�߂Ɏg�p
' �߂�l�F�Ȃ�
Sub MacroInput3dData(fps As Double, video_width As Long, csv_file_name As String, video_height As Double)

    Dim wb     As Variant
    Dim ws     As Variant
    Dim MaxRow As Long
    Dim i      As Long

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With 'With Application

    Sheets("�|�C���g�v�Z�V�[�g").Select
    Range("D2").Select

    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & csv_file_name)

    With wb
        Set ws = .Sheets(1)

        Range("B2").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy

        '���̃u�b�N�́u�\�t��v�V�[�g�֒l�\��t��
        ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Range("D2").PasteSpecial _
            xlPasteValuesAndNumberFormats

        '�R�s�[��Ԃ�����
        Application.CutCopyMode = False

        '�ۑ������I��
        .Close False
    End With 'With wb

    ' A ���� C �̎��Ԃ�\���Z�������̉�������
    ' angle.csv�𒣂�t�������Ƃ̍ŉ��s�ԍ����擾����
    MaxRow = Range("D2").End(xlDown).row
    For i = 0 To MaxRow - 2
        Range("A" & i + 2).Value = i
        'Range("B" & i + 2).Value = (8 * 60 + 42) * i / 15686
        Range("B" & i + 2).Value = i * (1 / fps)
        Range("C" & i + 2).FormulaR1C1 = "=LEFT(TEXT(RC[-1]/(24*60*60), ""hh:mm:ss.000""), 8)"
    Next

    'fps�l�̕ۑ�
    ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 199) = fps
    'video_width�l�̕ۑ�
    ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 198) = video_width
    'video_height�l�̕ۑ� ����̌����ɂ���Ď��������T�C�Y�𒲐����邽�߂Ɏg�p
    ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 197) = video_height

    ThisWorkbook.Save

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With 'With Application
End Sub

' ����1 �F�Ȃ�
' ����2 �F�Ȃ�
' �߂�l�F�Ȃ�
Sub VeryHiddenSheet()
    Sheets("�|�C���g�v�Z�V�[�g").Visible = xlVeryHidden
    Sheets("�����ݒ�V�[�g").Visible = xlVeryHidden
End Sub



'Python����Ăяo�������
' ����1 �F���於
' ����2 �F�t���[�����[�g
' �߂�l�F�Ȃ�
Sub MacroUpdateData(movieName As String, fps As Double)

    Dim tstart_first As Double

    If MEAGERE_TIME_MACROUPDATEDATA = True Then 'MacroUpdateData�̏������Ԃ𑪒肷��
        tstart_first = Timer
    End If

    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
        Dim max_row_num As Long
        Dim i As Long

        '��������s�����擾�i3��ڂ̍ŏI�Z���j
        max_row_num = .Cells(1, 3).End(xlDown).row

        '�������{�����́A�����I��Python�R�[�h���ōs���\�聚����
        '�t���O�����͂����Z���ɓ��͂���Ă���X�y�[�X���������ď�������
        '���C�������̎p���f�_�̐F���S�ė΂ɂȂ�s��̎b��΍�
        '�Z���͈͂��L�����ă������s���ɂȂ邽�߁Afor���ŏ������ו���
        For i = 4 To 253
            .Range(.Cells(2, i), .Cells(max_row_num, i)).Replace " ", ""
        Next

        'fps�l�̕ۑ�
        fps = .Cells(2, 199)

    End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")

    '�p������
    Call makeGraphJisya
'    Call makeGraphZensya


    '�m�C�Y����
    Call removeCaptionNoise(fps)

    '��ƕ����A���ԑ���
    Call fixSheetJisya
'    Call fixSheetZensya


    '�C���V�[�g�̍X�V
    'Call Module1.paintAll
    Call Module3.paintAll

    '��������
    Call outputCaption(movieName)

    '�V�[�g���B��
    Call VeryHiddenSheet

    'MacroUpdateData�̏������Ԃ𑪒肷��
    If MEAGERE_TIME_MACROUPDATEDATA = True Then
        ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, COLUMN_MEAGERE_TIME_MACROUPDATEDATA) = Format$(Timer - tstart_first, "0.00")
    End If

    '���񕪐͍ς݂̃t���O�𗧂Ă�
    ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g").Cells(2, 196) = 1

End Sub



'�p���d�ʓ_�����[�̑I���ƕۑ�
' ����1 �F���於
' �߂�l�F�Ȃ�
Sub MacroSaveData(movieName As String)

    '�H���]���V�[�g��\��
    ThisWorkbook.Worksheets("�H���]���V�[�g").Select

    '�H���]���V�[�g�̍�Ɩ��͎���͂ɕύX�������߃R�����g�A�E�g
    '�H���]���V�[�g�̍�Ɩ����L������
    'ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_KOUTEI_NAME, GH_HYOUKA_SHEET_COLUMN_KOUTEI_NAME).Value = cutLeftString(movieName, 16)

    '�H���]���V�[�g�̒��������L������
    ThisWorkbook.Worksheets("�H���]���V�[�g").Cells(GH_HYOUKA_SHEET_ROW_DATE, GH_HYOUKA_SHEET_COLUMN_KOUTEI_NAME).Value = Date

    '�p���d�ʓ_�����[���A�N�e�B�u�ɂ��ĕۑ�����
    Sheets("�H���]���V�[�g").Activate
    ThisWorkbook.Save



End Sub


Sub OutputOtrs()

    Dim max_row_num    As Long
    Dim i              As Long

    Dim targetRowCount As Long
    Dim writePoseNum   As Long
    Dim lastPoseNum    As Long
    Dim currentTime    As Double
    Dim lastTime       As Double
    Dim ret            As Long
    Dim destFilePath   As String
    Dim sourceFilePath As String

    Dim ReturnBook     As Workbook, targetWorkbook As Workbook
    Dim strYYYYMMDD    As String
    Dim PosExt         As Long
    Dim StrFileName    As String

    StrFileName = ThisWorkbook.Name
    PosExt = InStrRev(StrFileName, ".")

    '--- �g���q���������p�X�i�t�@�C�����j���i�[����ϐ� ---'
    Dim strFileExExt As String

    If (0 < PosExt) Then
        StrFileName = Left(StrFileName, PosExt - 1)
    End If


  'Now�֐��Ŏ擾�������ݓ��t��Format�Ő��`���ĕϐ��Ɋi�[
  strYYYYMMDD = Format(Now, "yyyymmdd_HHMMSS_")

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With 'With Application
    Set ReturnBook = ActiveWorkbook
    destFilePath = ActiveWorkbook.Path & "\" & StrFileName & "_otrs.xlsx"

    '����otrs�p�t�@�C��������΁A��x�폜���Ă���
    If Dir(destFilePath) <> "" Then
        Kill destFilePath
    End If

    '��Ɨp�̃��[�N�u�b�N�̃C���X�^���X�����

    If Dir(destFilePath) = "" Then
        '�V�����t�@�C�����쐬
        Set targetWorkbook = Workbooks.Add
        '�V�����t�@�C����VBA�����s�����t�@�C���Ɠ����t�H���_�ۑ�
        targetWorkbook.SaveAs destFilePath
    Else
        Set targetWorkbook = Workbooks.Open(destFilePath)
    End If
'    targetRowCount = 1
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_CYCLE).Value = "�T�C�N��"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NUM).Value = "No."
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = "�v�f��"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_TYPE).Value = "���1"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_COMPANY_TYPE).Value = "�g���Ǝ��"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_START_TIME).Value = "�X�^�[�g"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = "�v�f����"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_MOVE).Value = "����"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_forced).Value = "����"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_COMPARTINO).Value = "��r�l"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = "�v�f��"
'    targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = "�v�f����"

    ReturnBook.Activate
    lastPoseNum = -1
    lastTime = 0

    Dim CaptionName2(10) As String

    With ThisWorkbook.Worksheets("�����ݒ�V�[�g")
        CaptionName2(10) = .Cells(6, 2)
        CaptionName2(9) = .Cells(24, 2)
        CaptionName2(8) = .Cells(42, 2)
        CaptionName2(7) = .Cells(60, 2)
        CaptionName2(6) = .Cells(78, 2)
        CaptionName2(5) = .Cells(96, 2)
        CaptionName2(4) = .Cells(114, 2)
        CaptionName2(3) = .Cells(132, 2)
        CaptionName2(2) = .Cells(150, 2)
        CaptionName2(1) = .Cells(168, 2)
    End With 'With ThisWorkbook.Worksheets("�����ݒ�V�[�g")


'    For i = 2 To 10
'        posExt = InStrRev(CaptionName2(i), "-")
'        If (0 < posExt) Then
'            CaptionName2(i) = Mid(CaptionName2(i), posExt + 1, Len(CaptionName2(i)))
'        End If
'    Next


    CaptionName2(0) = "�f�[�^�Ȃ�"
        '�ȉ��̃p�^�[���ȊO�͂��̑��Ƃ���B
        '(10) �G���Ȃ��㔼�g�O��(30���`90��)
        '(9) �G���Ȃ��㔼�g�O��(15���`30��)
        '(8) �㔼�g�O��(45���`90��)
        '(7) �㔼�g�O��(30���`45��)
        '(6) �㔼�g�O��(90���`180��)
        '(4) �L���܂��͕ЕG���L��
        '(2) �㔼�g�O��(15���`30��)
        '(1) ��{�̗����p��
        '(0) ��"

    With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
        max_row_num = .Cells(1, 3).End(xlDown).row
        targetRowCount = 1
        Dim lastI As Long

        For i = 2 To max_row_num
            'COLUMN_DATA_RESULT_ORIGIN���󔒂̉\�������邽�߈�U���̑������Ă���
            writePoseNum = 0
            On Error Resume Next
            writePoseNum = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value '�L���v�V�����ԍ��̃Z�����
'            If writePoseNum = 3 Or _
'                writePoseNum = 5 Then
'                '�|�[�Y3��5�͂��̑��ɂ���B
'                writePoseNum = 0
'            End If
            '�ŏ��ɕʂ̃|�[�Y�ɕς���������~�����̂ň��ڂ͓���ɂ���B
            If i = 2 Then
                lastPoseNum = writePoseNum
                lastI = i - 2
            End If


            If lastPoseNum <> writePoseNum Then
                '����|�[�Y������Ă������Ԃ��K�v�i�؂�ւ������O�̎��ԁj
                currentTime = .Cells(i - 1, 2).Value
                '�������ݏ���
                targetWorkbook.Activate
                'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NUM).Value = lastPoseNum
                targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = CaptionName2(lastPoseNum)
                'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_START_TIME).Value = Round(lastTime, 5)
                targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = Round(currentTime - lastTime, 5)
                'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, 9).Value = lastI
                'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, 10).Value = i - 3
                lastI = i - 2
                targetRowCount = targetRowCount + 1

                lastTime = currentTime
                lastPoseNum = writePoseNum

                ReturnBook.Activate
            End If
        Next

        '���[�v�I����ɍŌ�Ɏ���Ă����p�����p�����Ă���Ȃ炻�����������
        If lastPoseNum = writePoseNum Then
            currentTime = .Cells(i - 1, 2).Value
            '�������ݏ���
            targetWorkbook.Activate
            'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NUM).Value = writePoseNum
            targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = CaptionName2(writePoseNum)
            'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_START_TIME).Value = Round(lastTime, 5)
            targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = Round(currentTime - lastTime, 5)

            'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, 9).Value = lastI
            'targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, 10).Value = i - 1 - 3
            ReturnBook.Activate
        End If
    End With 'With ThisWorkbook.Sheets("�|�C���g�v�Z�V�[�g")
    Sheet4.Activate
    ThisWorkbook.Save
    targetWorkbook.Close savechanges:=True
End Sub



Sub InputOtrs()

    Dim openFileName As String
    Dim lastTime     As Double

    Dim ex           As New Excel.Application    '// �����pExcel
    Dim wb           As Workbook                 '// ���[�N�u�b�N
    Dim r            As Range                    '// �擾�Ώۂ̃Z���͈�
    Dim sht          As Worksheet                '// �Q�ƃV�[�g

    Dim i            As Long
    Dim max_row_num  As Long
    Dim max_row      As Long

    'OTRS�G�N�X�|�[�g�t�@�C�����J��
    openFileName = Application.GetOpenFilename("OTRS�G�N�X�|�[�g�t�@�C��,*.xlsx?")

    '�������t�@�C�����J���ꂽ�ꍇ�̏���
    If openFileName <> "False" Then

        '// �ǂݎ���p�ŊJ��
        Set wb = ex.Workbooks.Open(FileName:=openFileName, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)

        '��������s�����擾�i1��ڂ̍ŏI�Z���j
        max_row_num = wb.Worksheets(1).Cells(1, 1).End(xlDown).row

        For i = 2 To max_row_num
            '�v�f���̃R�s�[�A�Z���w�i�F�͔��ɂ���
            ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 3).Value = wb.Worksheets(1).Cells(i, 1).Value
            ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 3).Interior.Color = RGB(255, 255, 255)

            '��ƏI�����Ԃ̃R�s�[�A�Z���w�i�F�͔��ɂ���
            If i = 2 Then
                '��������s�����擾�i3��ڂ̍ŏI�Z���j
                max_row = ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(1, 2).End(xlDown).row
                '�b����؂�グ�đ��
                lastTime = Application.WorksheetFunction.RoundUp(ThisWorkbook.Worksheets("�|�C���g�v�Z�V�[�g").Cells(max_row, 2), 0)

            End If

            If i <> max_row_num Then
                ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 36).Value = "�|"
                ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 37).Value = wb.Worksheets(1).Cells(i + 1, 2).Value
                ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 37).Interior.Color = RGB(255, 255, 255)
            Else
                ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 36).Value = "�|"
                ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 37).Value = lastTime
                ThisWorkbook.Worksheets("�p���d�ʓ_�����[").Cells(9 + i - 2, 37).Interior.Color = RGB(255, 255, 255)
            End If
        Next

        '// �u�b�N�����
        Call wb.Close

        '// Excel�A�v���P�[�V���������
        Call ex.Application.Quit

        '�f�[�^�X�V
        ClickUpdateData

    End If
End Sub

'�b��hh:mm:ss:ms�ɕϊ�����
Function timeConvert(seconds As Double) As String

    Dim milliseconds        As Long
    Dim remainingSeconds    As Long
    Dim minutes             As Long
    Dim hours               As Long

    '����h�~�̂��߂ɏ����_�ȉ���؂�̂ă~���b�E�b�����ɏo��
    milliseconds = (seconds - Int(seconds)) * 1000
    seconds = Int(seconds)

    remainingSeconds = seconds Mod 60
    minutes = (seconds Mod 3600) \ 60
    hours = seconds \ 3600

    timeConvert = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00") & "." & Format(milliseconds, "000")
End Function

