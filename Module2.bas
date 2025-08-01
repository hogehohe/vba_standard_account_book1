Option Explicit '変数の宣言を強制

'======================================================================================
'条件設定シートの各データの行番号、列番号を定義 (拳上概要の定数もここで定義）
'======================================================================================
Const KOBUSHIAGE_MISSING_DOWNLIM_TIME           As Double = 1       '（秒） 拳上欠損ノイズ判定に使う
Const TEKUBI_SPEED_UPLIM_PREDICT                As Double = 10      '（km/h）手首z位置の変化量上限　遮蔽検知に使う
Const MEAGERE_TIME_MACROUPDATEDATA              As Boolean = True   'TrueのときMacroUpdateDataの処理時間を測定する
Const KOBUSHIAGE_TIME_HOSEI_COEF_WORK           As Double = 5 / 355 '拳上時間補正係数 対象工程の中で最も作業時間が長い工程の　確認必要な欠損区間数/作業時間
Const KOBUSHIAGE_MISSING_DILATION_SIZE          As Double = 0.33    '（秒）拳上欠損の膨張処理に使う窓の大きさ（片側）
Const KOBUSHIAGE_MISSING_EROSION_SIZE           As Double = 0.33    '（秒）拳上欠損の収縮処理に使う窓の大きさ（片側）
Const KOBUSHIAGE_TIME_HOSEI_COEF_MISSING        As Double = 0.2     '拳上時間補正係数 確認必要な欠損区間数１個あたり

'makeGraph、outputCaption、fixGraphDataAndSheetモジュールの中に条件設定シートのセル内から値を読み出す部分あり

'======================================================================================
'ポイント計算シート上の各データの行番号、列番号を定義
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

Const COLUMN_DATA_MISSING_SECTION           As Long = 219

Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_ORG       As Long = 221
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_ORG       As Long = 222
Const COLUMN_KOBUSHIAGE_FORCED_SECTION                 As Long = 223 '拳上、腰曲げ、膝曲げの強制、判定フラグ、フラグの記憶
Const COLUMN_KOBUSHIAGE_RESULT                         As Long = 245
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

Const COLUMN_TEKUBI_RZ_SPEED                           As Long = 237 '右手首Z位置の差
Const COLUMN_TEKUBI_LZ_SPEED                           As Long = 238 '左手首Z位置の差
Const COLUMN_TEKUBI_Z_SPEED_OVER                       As Long = 239 '手首Z位置の差 しきい値超えフラグ
Const COLUMN_DATA_KOBUSHIAGE_MEASURE_SECTION_DST       As Long = 240 '拳上測定区間
Const COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST       As Long = 241 '拳上欠損区間
Const COLUMN_MEAGERE_TIME_MACROUPDATEDATA              As Long = 242 'MacroUpdateDataの処理時間を測定結果を格納する

Const COLUMN_DATA_RESULT_GH_KOBUSHIAGE      As Long = 245
Const COLUMN_DATA_RESULT_GH_KOSHIMAGE       As Long = 247
Const COLUMN_DATA_RESULT_GH_HIZAMAGE        As Long = 249
Const COLUMN_DATA_RESULT_GH_SONKYO          As Long = 251

Const COLUMN_GH_HIZA_L                      As Long = 252
Const COLUMN_GH_HIZA_R                      As Long = 253

Const COLUMN_MAX_NUMBER                     As Long = 256 '現在使用されている列番号の最大値


'======================================================================================
'姿勢重量点調査票シートの各データの行番号、列番号を定義
'======================================================================================
Const SHIJUTEN_SHEET_ROW_KOUTEI_NAME                            As Long = 3
Const SHIJUTEN_SHEET_ROW_POSESTART_INDEX                        As Long = 9
Const SHIJUTEN_SHEET_ROW_EXPAND_NUMBER_CHECK                    As Long = 29

Const SHIJUTEN_SHEET_EXPAND_NUM_CHECK_WORD                      As String = "その他の時間（定時稼働時間7.5H-Σ延べ時間）"

Const SHIJUTEN_SHEET_COLUMN_WORK_NUMBER                         As Long = 2
Const SHIJUTEN_SHEET_COLUMN_WORK_NAME                           As Long = 3
Const SHIJUTEN_SHEET_COLUMN_KOUTEI_NAME                         As Long = 4
Const SHIJUTEN_SHEET_COLUMN_WORK_TIME                           As Long = 9
Const SHIJUTEN_SHEET_COLUMN_POSE_START_INDEX                    As Long = 10

Const SHIJUTEN_SHEET_COLUMN_WORKSTART_TIME                      As Long = 36
Const SHIJUTEN_SHEET_COLUMN_WORKEND_TIME                        As Long = 38

Const SHIJUTEN_SHEET_COLUMN_DATA_MISSING_SECTION                As Long = 46
Const SHIJUTEN_SHEET_COLUMN_DATA_PREDICT_SECTION                As Long = 47

Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_TIME                     As Long = 49 '拳上時間
Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_TIME                      As Long = 51 '腰曲げ時間
Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_TIME                       As Long = 53 '膝曲げ時間

Const SHIJUTEN_SHEET_COLUMN_KOBUSHIAGE_MISSING_TIME             As Long = 55 '拳上欠損区間

Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_MISSING_TIME              As Long = 57 '腰曲げ欠損区間
Const SHIJUTEN_SHEET_COLUMN_KOSHIMAGE_PREDICT_TIME              As Long = 58 '腰曲げ推定区間

Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_MISSING_TIME               As Long = 60 '膝曲げ欠損区間
Const SHIJUTEN_SHEET_COLUMN_HIZAMAGE_PREDICT_TIME               As Long = 61 '膝曲げ推定区間


'======================================================================================
'工程評価シートの各データの行番号、列番号を定義
'======================================================================================

Const GH_HYOUKA_SHEET_ROW_YOUSO_HANTEI_LIMIT           As Long = 2 '要素判定しきい値
Const GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_LIMIT           As Long = 3 '要素判定しきい値
Const GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_LIMIT            As Long = 4 '要素判定しきい値
Const GH_HYOUKA_SHEET_ROW_KOUTEI_NAME                  As Long = 5 '工程名
Const GH_HYOUKA_SHEET_ROW_DATE                         As Long = 6 '調査日
Const GH_HYOUKA_SHEET_ROW_POSESTART                    As Long = 15
Const GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK          As Long = 115

Const GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_CAPTION         As Long = 13 '要素判定しきい値のキャプション記入セル行
Const GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_CAPTION          As Long = 13 '要素判定しきい値のキャプション記入セル行

Const GH_HYOUKA_SHEET_EXPAND_NUM_CHECK_WORD            As String = "合計"
Const GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_NG             As String = "×"
Const GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_OK             As String = "○"

Const GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER               As Long = 2 '作業No.
Const GH_HYOUKA_SHEET_COLUMN_WORK_NAME                 As Long = 3 '要素作業名
Const GH_HYOUKA_SHEET_COLUMN_KOUTEI_NAME               As Long = 4 '工程名と調査日
Const GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT       As Long = 12 '要素判定結果
Const GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME            As Long = 13 '作業開始時間
Const GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME              As Long = 16 '作業終了時間
Const GH_HYOUKA_SHEET_COLUMN_WORK_TIME                 As Long = 19 '作業時間
Const GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME           As Long = 20 '拳上げ時間
Const GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME            As Long = 21 '腰曲げ時間
Const GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME             As Long = 22 '膝曲げ時間
Const GH_HYOUKA_SHEET_COLUMN_NG_TIME_A                 As Long = 26 '疲労度評価（A)のNG作業時間
Const GH_HYOUKA_SHEET_COLUMN_NG_TIME_B                 As Long = 29 '気遣い度評価（B)のNG作業時間
Const GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT              As Long = 36 '要素判定しきい値

Const GH_HYOUKA_SHEET_COLUMN_KOSHI_HANTEI_CAPTION      As Long = 21 '要素判定しきい値のキャプション記入セル列
Const GH_HYOUKA_SHEET_COLUMN_HIZA_HANTEI_CAPTION       As Long = 22 '要素判定しきい値のキャプション記入セル列

'======================================================================================
'外販用　姿勢判定のしきい値を定義
'======================================================================================

Const GH_ANGLE_KOSHIMAGE_MAX As Double = 180
Const GH_ANGLE_HIZAMAGE_MAX As Double = 180

'======================================================================================
'DataAdjustingSheet用
'======================================================================================
Const LIMIT_COLUMN           As Long = 16200

'======================================================================================
'字幕情報の定義
'======================================================================================
Const CAPTION_TRACK2_FILE_NAME_SOEJI           As String = "2" '字幕トラック２用のファイル名末尾につける添字
Const CAPTION_CHUKAN_FILE_NAME_SOEJI           As String = "tmp" '中間ファイルにつける添字
'各種字幕のフォントサイズ係数
'分母の値のため、値が小さいほど文字は大きい
'動画が縦の時
Const TRACK1_TATE_UPPER_COEF                   As Long = 22 'トラック1用：上段
Const TRACK1_TATE_LOWER_COEF                   As Long = 11 'トラック1用：下段
Const TRACK2_TATE_1ST_COEF                     As Long = 22 'トラック2用：1段目
Const TRACK2_TATE_2ND_COEF                     As Long = 22 'トラック2用：2段目
Const TRACK2_TATE_3RD_COEF                     As Long = 13 'トラック2用：3段目

'動画が横の時
Const TRACK1_YOKO_UPPER_COEF                   As Long = 30 'トラック1用：上段
Const TRACK1_YOKO_LOWER_COEF                   As Long = 15 'トラック1用：下段
Const TRACK2_YOKO_1ST_COEF                     As Long = 30 'トラック2用：1段目
Const TRACK2_YOKO_2ND_COEF                     As Long = 30 'トラック2用：2段目
Const TRACK2_YOKO_3RD_COEF                     As Long = 18 'トラック2用：3段目

'各種字幕の色
Const COLOR_DATA_REMOVE_SECTION                As String = "#bfbfbf" 'グレー
Const COLOR_DATA_FORCED_SECTION                As String = "#0033cc" '青色
Const COLOR_DATA_MISSING_SECTION               As String = "#ff7c80" '朱色
Const COLOR_DATA_PREDICT_SECTION               As String = "#fcf600" '黄色
Const COLOR_DATA_MEASURE_SECTION               As String = "#00b0f0" '水色
Const COLOR_DATA_RESULT_GREEN                  As String = "#00b050" '緑色
Const COLOR_DATA_RESULT_YELLOW                 As String = "#ffc000" '黄色
Const COLOR_DATA_RESULT_RED                    As String = "#c00000" '赤色
Const COLOR_DATA_RESULT_GLAY                   As String = "#bfbfbf" 'グレー

'帯グラフのデータ（信頼度）を示す字幕文字列（字幕トラック1用 上段右側に表示）
Const CAPTION_DATA_MEASURE_SECTION             As String = "【データ測定区間】"
Const CAPTION_DATA_PREDICT_SECTION             As String = "【データ推定区間】"
Const CAPTION_DATA_REMOVE_SECTION              As String = "【データ除外区間】"
Const CAPTION_DATA_FORCED_SECTION              As String = "【データ強制区間】"
Const CAPTION_DATA_MISSING_SECTION             As String = "【データ欠損区間】"

'帯グラフのデータ（信頼度）を示す字幕文字列（字幕トラック2用 2段目に表示）
Const CAPTION_DATA_TRACK2_MEASURE_SECTION      As String = "【データ測定区間】"
Const CAPTION_DATA_TRACK2_PREDICT_SECTION      As String = "【データ推定区間】"
Const CAPTION_DATA_TRACK2_REMOVE_SECTION       As String = "【データ除外区間】"
Const CAPTION_DATA_TRACK2_FORCED_SECTION       As String = "【データ強制区間】"
Const CAPTION_DATA_TRACK2_MISSING_SECTION      As String = "【データ欠損区間】"

'外販用の字幕文字列（字幕トラック2用 3段目に表示）
Const CAPTION_A_RESULT_NAME1  As String = "　　　　拳上"
Const CAPTION_B_RESULT_NAME1  As String = "  　　腰曲げ　 　"
Const CAPTION_C_RESULT_NAME1  As String = "膝曲げ"

'外販用の条件字幕文字列（字幕トラック2用 4段目に表示）
Const CAPTION_A_RESULT_NAME2  As String = "手首が肩より上"
Const CAPTION_B_RESULT_NAME2  As String = "45°以上"
Const CAPTION_C_RESULT_NAME2  As String = "60°以上"

'キャプションノイズ除去の閾値
Const CAPTION_REMOVE_NOISE_SECOND              As Double = 0.1 'キャプションノイズを除去する長さ(秒) （〜未満なら除去）

'姿勢素点の値によって、緑／黄／赤を分ける際の境界条件
Const DATA_SEPARATION_GREEN_BOTTOM             As Long = 1
Const DATA_SEPARATION_GREEN_TOP                As Long = 2
Const DATA_SEPARATION_YELLOW_BOTTOM            As Long = 3
Const DATA_SEPARATION_YELLOW_TOP               As Long = 5
Const DATA_SEPARATION_RED_BOTTOM               As Long = 6
Const DATA_SEPARATION_RED_TOP                  As Long = 10


'処理時間短縮のため、更新をストップ
' 引数1 ：なし
' 戻り値：なし
Function stopUpdate()
    '表示・更新をオフにする
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Function


'処理時間短縮のため、更新をリスタート
' 引数1 ：なし
' 戻り値：なし
Function restartUpdate()
    '表示・更新をオンに戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function


'文字列sの左側からi文字を削除する関数
' 引数1 ：文字列
' 引数2 ：削除文字数
' 戻り値：削除後の文字列
Function cutLeftString(s, i As Long) As String
    Dim iLen As Long '// 文字列長

    '// 文字列ではない場合
    If VarType(s) <> vbString Then
        Exit Function
    End If

    iLen = Len(s)

    '// 文字列長より指定文字数が大きい場合
    If iLen < i Then
        Exit Function
    End If

    '// 指定文字数を削除して返す
    cutLeftString = Right(s, iLen - i)
End Function


'膨張処理
' 引数1 ：処理前の配列
' 引数2 ：配列の数
' 引数3 ：窓の大きさ
' 戻り値：処理後の配列
Function dilation(array_src() As Long, max_array_num As Long, window_size As Long)
        Dim i As Long
        Dim j As Long
        Dim array_dst() As Long

        '窓サイズ分の端のフラグが消えるのを防止
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


'収縮処理
' 引数1 ：処理前の配列
' 引数2 ：配列の数
' 引数3 ：窓の大きさ
' 戻り値：処理後の配列
Function erosion(array_src() As Long, max_array_num As Long, window_size As Long)
        Dim i As Long
        Dim j As Long
        Dim array_dst() As Long

        '窓サイズ分の端のフラグが消えるのを防止
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


'拳上のフラグ生成
' 引数1 ：なし
' 戻り値：なし

Sub kobusiage_hantei()

    '表示・更新をオフにする
    Call stopUpdate

    Dim KataPositionRz()       As Double
    Dim KataPositionLz()       As Double
    Dim TekubiPositionRz()     As Double
    Dim TekubiPositionLz()     As Double
    Dim TekubiSpeedRz()        As Double
    Dim TekubiSpeedLz()        As Double

    Dim kobushiage_missing_array()     As Long '拳上欠損フラグを格納する配列
    Dim tekubi_zspeed_over_array()     As Long '手首位置の差のしきい値超え
    Dim kobushiage_measure_array()     As Long '拳上測定区間
    Dim kobushiage_array()             As Long '拳上時間

    Dim kobushiage_missing_limit       As Long '拳上欠損フラグのノイズ判定しきい値
    Dim kobushiage_missing_count       As Long '拳上欠損フラグの連続回数をカウント
    Dim kobushiage_missing_section_num As Long '拳上欠損区間がしきい値を超える数をカウント。拳上時間の補正に使う

    Dim window_size_dilation           As Long '膨張に使う窓の大きさ
    Dim window_size_erosion            As Long '収縮に使う窓の大きさ
    Dim work_time                      As Double  '作業時間
    Dim kobushiage_hosei_time          As Double  '拳上補正時間
    Dim kobushiage_hosei_frame_num     As Long '拳上補正フレーム数

    Dim fps                            As Double 'フレームレート

    Dim max_row_num                    As Long '行の末尾
    Dim max_array_num                  As Long '配列の末尾

    Dim i                              As Long
    Dim j                              As Long

    With ThisWorkbook.Sheets("ポイント計算シート")

        '処理する行数を取得
        max_row_num = getLastRow()

        max_array_num = max_row_num - 1 - 1 '2行目からセルに値が入るため-1、配列は0から使うため-1

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
        'ここから拳上判定
        '-------------------------------------------------------------------------------------------
        fps = getFps() 'フレームレートを取得

        '手首と肩の高さを配列に格納
        For i = 0 To max_array_num
            '肩,手首の高さ(Z座標)を読み出し　配列は0から始まるため+1、セルの値は2行目から始まるため+1
            TekubiPositionLz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_TEKUBI_L_Z).Value
            TekubiPositionRz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_TEKUBI_R_Z).Value
            KataPositionLz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_KATA_L_Z).Value
            KataPositionRz(i, 0) = .Cells(i + 1 + 1, COLUMN_POS_KATA_R_Z).Value
            kobushiage_missing_array(i, 0) = .Cells(i + 1 + 1, COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_ORG).Value
        Next

        '拳上判定をして、結果を配列に格納
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

        '補正後の拳上時間配列をセルに貼り付け
        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE - 1)).Value = kobushiage_array
        .Range(.Cells(2, COLUMN_DATA_RESULT_GH_KOBUSHIAGE), .Cells(max_row_num, COLUMN_DATA_RESULT_GH_KOBUSHIAGE)).Value = kobushiage_array
    End With

    '表示・更新をオンに戻す
    Call restartUpdate
End Sub


'拳上、腰曲げ、膝曲げの判定
' 引数1 ：なし
' 戻り値：なし
Sub makeGraphJisya()

    '表示・更新をオフにする
    Call stopUpdate

    '拳上の判定
    'コードが長いため別の関数にする
    Call kobusiage_hantei

    '条件設定シートから読み込むパラメータ
    Dim AngleKoshiMin  As Double
    Dim AngleKoshiMax  As Double

    Dim AngleHizaMin   As Double
    Dim AngleHizaMax   As Double

    '関節角度のデータを読み込む変数
    Dim ValAngleKoshi   As Double
    Dim ValAngleHizaR   As Double
    Dim ValAngleHizaL   As Double

    '判定結果を格納する配列
    Dim KoshimageArray() As Double
    Dim HizamageArray()  As Double

    '外販用の膝角度
    '社内と定義が異なるため注意
    Dim HizaAngleLArray() As Double
    Dim HizaAngleRArray() As Double


    'その他変数
    Dim start_frame            As Long
    Dim end_frame              As Long
    Dim fps                    As Double 'フレームレート
    Dim correctPose            As Boolean
    Dim mSeconds               As String
    Dim totalSecond            As Long
    Dim tempSecond             As Long
    Dim hour, min, sec         As Long
    Dim t                      As Date
    Dim ds                     As String

    Dim max_row_num            As Long '行の末尾
    Dim max_array_num          As Long '配列の末尾

    Dim i                      As Long
    Dim j                      As Long
    Dim data_no                As Long

    Dim PointCalcSheetArray    As Variant

    '判定のしきい値を代入
    AngleKoshiMin = ThisWorkbook.Worksheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_LIMIT, GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT).Value
    AngleKoshiMax = GH_ANGLE_KOSHIMAGE_MAX

    AngleHizaMin = ThisWorkbook.Worksheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_LIMIT, GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT).Value
    AngleHizaMax = GH_ANGLE_HIZAMAGE_MAX

    '判定のしきい値に応じて表内の表示文字列（キャプション）を書き換え
    ThisWorkbook.Worksheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_KOSHI_HANTEI_CAPTION, GH_HYOUKA_SHEET_COLUMN_KOSHI_HANTEI_CAPTION).Value = "腰角度" & vbLf & CStr(AngleKoshiMin) & "°" & vbLf & "以上"
    ThisWorkbook.Worksheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_HIZA_HANTEI_CAPTION, GH_HYOUKA_SHEET_COLUMN_HIZA_HANTEI_CAPTION).Value = "膝角度" & vbLf & CStr(AngleHizaMin) & "°" & vbLf & "以上"


    With ThisWorkbook.Sheets("ポイント計算シート")

        '処理する行数を取得
        max_row_num = getLastRow()

        'ポイント計算シートの中身を配列に読込
        PointCalcSheetArray = .Range(.Cells(1, 1), .Cells(max_row_num, COLUMN_MAX_NUMBER))

        max_array_num = max_row_num - 1 - 1 '2行目からセルに値が入るため-1、配列は0から使うため-1

        ReDim HizaAngleLArray(max_array_num, 0)
        ReDim HizaAngleRArray(max_array_num, 0)

        '-------------------------------------------------------------------------------------------
        'ここから膝角度計算
        '-------------------------------------------------------------------------------------------
        For i = 0 To max_array_num
            HizaAngleLArray(i, 0) = 180 - .Cells(i + 2, COLUMN_HIZA_L_ANGLE).Value
            HizaAngleRArray(i, 0) = 180 - .Cells(i + 2, COLUMN_HIZA_R_ANGLE).Value
        Next

        '-------------------------------------------------------------------------------------------
        'ここから姿勢判定
        '-------------------------------------------------------------------------------------------
        For i = 2 To max_row_num

            'キャプション時刻生成
            mSeconds = Right(Format(WorksheetFunction.RoundDown(PointCalcSheetArray(i, 2), 3), "0.000"), 3) '小数点以下のみ取得
            totalSecond = Application.WorksheetFunction.RoundDown(.Cells(i, 2), 0)
            hour = WorksheetFunction.RoundDown(totalSecond / 3600, 0)
            min = WorksheetFunction.RoundDown((totalSecond Mod 3600) / 60, 0)
            sec = totalSecond Mod 60
            t = TimeSerial(hour, min, sec)
            ds = Format(t, "hh:mm:ss")

            'キャプション時刻の代入
            PointCalcSheetArray(i, COLUMN_ROUGH_TIME) = Format(t, "hh:mm:ss") & "," & mSeconds

            '関節角度の読み出し
            ValAngleKoshi = CDbl(PointCalcSheetArray(i, COLUMN_KOSHI_ANGLE))
            ValAngleHizaL = CDbl(PointCalcSheetArray(i, COLUMN_HIZA_L_ANGLE))
            ValAngleHizaR = CDbl(PointCalcSheetArray(i, COLUMN_HIZA_R_ANGLE))

            '腰曲げの判定
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

            '膝曲げの判定
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

            'キャプション時刻のセル代入
            PointCalcSheetArray(i, COLUMN_ROUGH_TIME) = Format(t, "hh:mm:ss") & "," & mSeconds

        Next

        '-------------------------------------------------------------------------------------------
        'ここから配列の中身をポイント計算シートに書込
        '-------------------------------------------------------------------------------------------
        .Range(.Cells(1, 1), .Cells(max_row_num, COLUMN_MAX_NUMBER)) = PointCalcSheetArray

        '外販用膝角度
        .Range(.Cells(2, COLUMN_GH_HIZA_L), .Cells(max_row_num, COLUMN_GH_HIZA_L)).Value = HizaAngleLArray
        .Range(.Cells(2, COLUMN_GH_HIZA_R), .Cells(max_row_num, COLUMN_GH_HIZA_R)).Value = HizaAngleRArray

    End With

    '表示・更新をオンに戻す
    Call restartUpdate

End Sub


'姿勢素点の字幕、フラグのノイズを消去する
' 引数1 ：フレームレート
' 戻り値：なし
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

    '表示・更新をオフにする
    Call stopUpdate

    With ThisWorkbook.Sheets("ポイント計算シート")

        '処理する行数を取得（3列目の最終セル）
        max_row_num = getLastRow()
        max_array_num = max_row_num - 1 - 1 '2行目からセルに値が入るため-1、配列は0から使うため-1

        '下方向へ探索する際の起点(i), 終点(i_max)
        i_max = max_row_num - noise_num - 1

        'キャプションのノイズ除去
        For i = 2 To i_max

            currentValue = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value
            targetValue = .Cells(i + 1, COLUMN_DATA_RESULT_ORIGIN).Value

            '判定結果が変わったとき
            If currentValue <> targetValue Then

                'ノイズかどうか探索する 起点(j), 終点(j_max)
                j_max = i + 1 + noise_num - 1
                sameValueNum = 1
                For j = i + 2 To j_max
                    compareValue = .Cells(j, COLUMN_DATA_RESULT_ORIGIN).Value
                    '判定結果が変わったらループを抜ける
                    If targetValue = compareValue Then
                        sameValueNum = sameValueNum + 1
                    Else
                        Exit For
                    End If
                Next

                'ノイズが見つかったときの処理
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
    End With

    '表示・更新をオンに戻す
    Call restartUpdate
End Function


' １回目は分割なしのデータ入力
' 更新ボタンが押された際は、作業開始時間を使って分割
' 引数1 ：フレームレート
' 戻り値：なし
Sub fixSheetJisya()

    '表示・更新をオフにする
    Call stopUpdate

    Dim fps As Double

    Dim separate_work_time    As Double 'tとt0の差を取得する
    Dim t0                    As Double '1つ前のtを一時保存する
    Dim t                     As Double '作業時間

    Dim i                     As Long
    Dim j                     As Long

    Dim max_row_num           As Long

    Dim expand_no             As Long '追加された行数を調べるために使う

    Dim Kobushiage_flag       As Long
    Dim koshimage_flag        As Long
    Dim hizamage_flag         As Long

    Dim start_frame           As Long
    Dim end_frame             As Long

    Dim data_no               As Long

    Dim removeFrames          As Long
    Dim workFrames            As Long

    Dim top_jogai_end         As Long
    Dim bottom_jogai_start    As Long

    Dim worktime_sum          As Double

    Dim seconds             As Double
    Dim hours               As String
    Dim minutes             As String
    Dim remainingSeconds    As String
    Dim milliseconds        As String
    Dim format_time         As String

    Dim youso_hantei_limit As Double
    Dim NG_time_A As Double
    Dim NG_time_B As Double

    'フレームレートを取得
    fps = getFps()

    'ポイント計算シートの最終行を取得
    max_row_num = getLastRow()

    '各種初期化
    removeFrames = 0

    '要素作業判定のしきい値を読込
    youso_hantei_limit = ThisWorkbook.Worksheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_YOUSO_HANTEI_LIMIT, GH_HYOUKA_SHEET_COLUMN_HANTEI_LIMIT).Value

    '処理する追加行数を取得する
    '"要素数"のセル位置の移動量を調べる  ※最大999行(<1050)にする
    expand_no = 0
    Do While ThisWorkbook.Worksheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK + expand_no, GH_HYOUKA_SHEET_COLUMN_WORK_NUMBER) <> _
    GH_HYOUKA_SHEET_EXPAND_NUM_CHECK_WORD And expand_no < 1050
        expand_no = expand_no + 1
    Loop


    '工程評価シートに値を入力
    With ThisWorkbook.Sheets("工程評価シート")
        'ここから初回分析のための処理
        '作業開始時間が空の場合は、0.0を入力
        If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) = True Then
            .Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value = 0
        End If

        '作業終了時間が空の場合は、ポイント計算シート最終行から計算して入力
        If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME)) = True Then
            seconds = max_row_num / fps 'ここに変換したい秒数を入力してください

            format_time = timeConvert(seconds)

            .Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value = format_time

        End If

        'ここから帳票更新のための処理
        '動画の先頭に除外がある場合、除外の末尾より一つ下のセルから１つ目の作業開始時間を計算する
        With ThisWorkbook.Sheets("ポイント計算シート")
            '除外フラグの先頭が0の時
            If .Cells(2, COLUMN_DATA_REMOVE_SECTION) = 0 Then
                '0秒にする
                ThisWorkbook.Sheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value = 0

            '除外フラグの先頭が１の時
            ElseIf .Cells(2, COLUMN_DATA_REMOVE_SECTION) = 1 Then
                'リセット
                top_jogai_end = 0
                '除外の末尾を調べる
                '除外フラグが1でなくなるまでループ
                Do While .Cells(2 + top_jogai_end, COLUMN_DATA_REMOVE_SECTION) = 1
                    top_jogai_end = top_jogai_end + 1
                Loop

                '除外の終了時間を計算して開始時間の１行目に入力
                seconds = top_jogai_end / fps 'ここに変換したい秒数を入力してください

                format_time = timeConvert(seconds)

                ThisWorkbook.Sheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value = format_time
            End If
        End With

        'ここから作業分割に関する処理
        For i = 0 To GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK - GH_HYOUKA_SHEET_ROW_POSESTART - 1 + expand_no
            '作業開始時間が空なら
            If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) Then
                '作業名、作業終了時間、作業時間、拳上、腰曲げ、膝曲げを空にする
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NAME).MergeArea.ClearContents 'セル結合があるため
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).MergeArea.ClearContents 'セル結合があるため
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).MergeArea.ClearContents 'セル結合があるため
                'NG時間を空にする
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME).ClearContents
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME).ClearContents
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME).ClearContents
                '要素作業の判定を空にする
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT).ClearContents

            '作業開始時間が入力されているなら
            Else
                'ここから作業名の入力
                '作業名が空なら入力する
                If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NAME)) Then
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_NAME) = "作業" & i + 1
                End If

                'ここから作業終了時間の入力
                '１つ先の行の作業開始時間が空の時
                If IsEmpty(.Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i + 1, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) Then
                    '動画の末尾に除外がない場合、ポイント計算シート最終行から作業終了時間を計算して入力する
                    If ThisWorkbook.Worksheets("ポイント計算シート").Cells(max_row_num, COLUMN_DATA_REMOVE_SECTION).Value <> 1 Then
                        'max_row_numで算出した場合、見出しとindex 0 が考慮されない為、最終セルの値を直接参照
                        seconds = ThisWorkbook.Worksheets("ポイント計算シート").Cells(max_row_num, COLUMN_POSE_KEEP_TIME).Value

                        Debug.Print "seconds:", seconds

                        format_time = timeConvert(seconds)

                        .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value = format_time

                    '動画の末尾に除外がある場合、除外の先頭より一つ上のセルから終了時間を計算する
                    ElseIf ThisWorkbook.Worksheets("ポイント計算シート").Cells(max_row_num, COLUMN_DATA_REMOVE_SECTION).Value = 1 Then
                        'カウントリセット
                        bottom_jogai_start = 0
                        'max_row_num行目から一つずつ上がって、除外の先頭位置を探す
                        Do While ThisWorkbook.Worksheets("ポイント計算シート").Cells(max_row_num - bottom_jogai_start, COLUMN_DATA_REMOVE_SECTION) = 1
                            bottom_jogai_start = bottom_jogai_start + 1
                        Loop

                        '動画末尾にある除外の開始時間を計算して入力
                        'ポイント計算シートの見出し1行分を考慮に含め-1
                        seconds = (max_row_num - bottom_jogai_start - 1) / fps

                        format_time = timeConvert(seconds)

                        .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value = format_time
                    End If


                '１つ先の行の作業開始時間に値がある時、その値を入れる
                Else
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME).Value _
                        = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i + 1, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME).Value
                End If

                '行程評価シートで計算式が入力されたセルの値を更新する
                Call restartUpdate
                Call stopUpdate

                '作業終了時間と作業開始時間から作業時間を計算してセルに入力
                'セル結合があるため+2することで秒数セルを参照
                .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).Value = _
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME + 2).Value _
                    - .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME + 2).Value
            End If
        Next

    End With

    '時間を初期値に設定
    separate_work_time = 0
    t0 = 0
    '動画先頭を除外したときに評価のスタートが0.0秒ではなくなるため変更
    t = ThisWorkbook.Sheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_POSESTART, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME + 2).Value

    'ポイント計算シートのフラグをカウントして、各作業姿勢の時間を計算する
    For i = 0 To GH_HYOUKA_SHEET_ROW_EXPAND_NUMBER_CHECK - GH_HYOUKA_SHEET_ROW_POSESTART - 1 + expand_no

        '作業開始時間が空なら分割処理はしない
        If IsEmpty(ThisWorkbook.Sheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKSTART_TIME)) Then

        '作業開始時間が入力されているなら分割処理をする
        Else
            separate_work_time = ThisWorkbook.Sheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORKEND_TIME + 2).Value
            t0 = t
            t = separate_work_time '作業時間を単一で入力する場合
            '秒数からフレーム数へ変換
            start_frame = t0 * fps
            end_frame = t * fps - 1

            'ここからポイント計算シートのフラグをカウント
            With ThisWorkbook.Sheets("ポイント計算シート")

                'カウンターをリセット
                Kobushiage_flag = 0
                koshimage_flag = 0
                hizamage_flag = 0

                'start_frameフレーム(t0秒) から end_frameフレーム(t秒) までの処理
                If start_frame < end_frame Then

                    Debug.Print "start_end:", start_frame, ":", end_frame
                    '分割時に除外フレームカウントを初期化
                    removeFrames = 0

                    For j = start_frame To end_frame

                        '拳上フラグをカウント
                        data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_KOBUSHIAGE).Value
                        If data_no = 1 Then
                            Kobushiage_flag = Kobushiage_flag + 1
                        End If

                        '腰曲げフラグをカウント
                        data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_KOSHIMAGE).Value
                        If data_no = 1 Then
                            koshimage_flag = koshimage_flag + 1
                        End If

                        '膝曲げフラグをカウント
                        data_no = .Cells(2 + j, COLUMN_DATA_RESULT_GH_HIZAMAGE).Value
                        If data_no = 1 Then
                            hizamage_flag = hizamage_flag + 1
                        End If

                        '除外区間をカウント
                        data_no = .Cells(2 + j, COLUMN_DATA_REMOVE_SECTION).Value
                        If data_no = 1 Then
                            removeFrames = removeFrames + 1
                        End If
                    Next

                    '作業時間合計値を算出
                    workFrames = (end_frame + 1 - start_frame) - removeFrames
                    ThisWorkbook.Sheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).Value = workFrames / fps

                Else
                    'start_frameがend_frameより大きい場合は、作業時間を0にする
                    ThisWorkbook.Sheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_WORK_TIME).Value = 0
                End If
            End With

            'ここからカウントしたフラグを時間に変換して、工程評価シートに入力
            With ThisWorkbook.Sheets("工程評価シート")

                '拳上に対する個別処理
                If Kobushiage_flag = 0 Then
                    '姿勢要素時間（フレーム数）が0のときは、空白セルにする
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME).Value = ""
                Else
                    '姿勢要素時間（フレーム数）があれば代入する
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOBUSHIAGE_TIME).Value = Kobushiage_flag / fps
                End If

                '腰曲げに対する個別処理
                If koshimage_flag = 0 Then
                    '姿勢要素時間（フレーム数）が0のときは、空白セルにする
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME).Value = ""
                Else
                    '姿勢要素時間（フレーム数）があれば代入する
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_KOSHIMAGE_TIME).Value = koshimage_flag / fps
                End If

                '膝曲げに対する個別処理
                If hizamage_flag = 0 Then
                    '姿勢要素時間（フレーム数）が0のときは、空白セルにする
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME).Value = ""
                Else
                    '姿勢要素時間（フレーム数）があれば代入する
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_HIZAMAGE_TIME).Value = hizamage_flag / fps
                End If
            End With 'With ThisWorkbook.Sheets("工程評価シート")

            '要素作業判定
            With ThisWorkbook.Sheets("工程評価シート")
                'セル計算の値を参照するため、ストップする
                Call restartUpdate
                Call stopUpdate

                '疲労度評価、気遣い度評価のNG作業時間を読込
                NG_time_A = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_NG_TIME_A).Value
                NG_time_B = .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_NG_TIME_B).Value

                '判定して結果を書き込み
                '×の場合
                If NG_time_A >= youso_hantei_limit Or NG_time_B >= youso_hantei_limit Then
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT).Value = GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_NG
                '○の場合
                Else
                    .Cells(GH_HYOUKA_SHEET_ROW_POSESTART + i, GH_HYOUKA_SHEET_COLUMN_YOUSO_HANTEI_RESULT).Value = GH_HYOUKA_SHEET_YOUSO_HANTEI_WORD_OK
                End If
            End With
        End If
    Next

    '表示・更新をオンに戻す
    Call restartUpdate

End Sub


'字幕ファイル出力
'引数1 ：動画名
'戻り値：なし
Function outputCaption(movieName As String)
    Dim i                           As Long
    Dim max_row_num                 As Long

    '動画の縦横を比較して文字サイズ調整するため、幅・高さどちらも使用する
    Dim video_width                 As Long '入力動画の幅 ※3Dポーズが結合された幅ではないため注意
    Dim video_height                As Long '入力動画の高さ

    '※coefはcoefficient（係数、率）の略記
    Dim track1_coef_font_size1      As Long '字幕トラック1用  上段のサイズ調整用係数
    Dim track1_coef_font_size2      As Long '字幕トラック1用  下段のサイズ調整用係数
    Dim track1_font_size1           As Long '字幕トラック1用  上段のサイズ
    Dim track1_font_size2           As Long '字幕トラック1用  下段のサイズ

    Dim track2_coef_font_size1      As Long '字幕トラック2用 1段目のサイズ調整用係数
    Dim track2_coef_font_size2      As Long '字幕トラック2用 2段目のサイズ調整用係数
    Dim track2_coef_font_size3      As Long '字幕トラック2用 3段目のサイズ調整用係数
    Dim track2_font_size1           As Long '字幕トラック2用 1段目のサイズ
    Dim track2_font_size2           As Long '字幕トラック2用 2段目のサイズ
    Dim track2_font_size3           As Long '字幕トラック2用 3段目のサイズ

    Dim CaptionName0                As String  '字幕トラック1用 上段左側 作業名          の字幕文字列
    Dim CaptionName1                As String  '字幕トラック1用 上段右側 帯グラフのデータ（信頼度）の字幕文字列
    Dim CaptionName2(10)            As String  '字幕トラック1用 下段 評価除外(添え字0)+姿勢素点1〜10(添え字1〜10)の字幕文字列
    Dim CaptionNo2                  As Long 'CaptionName2(10)にアクセスする際の添え字格納用変数

    Dim CaptionName2Kobushiage      As String '字幕トラック2用 2段目 拳上区間の字幕文字列
    Dim CaptionName2Koshimage       As String '字幕トラック2用 2段目 腰曲げデータ区間の字幕文字列
    Dim CaptionName2Hizamage        As String '字幕トラック2用 2段目 膝曲げデータ区間の字幕文字列

    Dim CaptionName3Kobushiage      As String '字幕トラック2用 ３段目 拳上の字幕文字列
    Dim CaptionName3Koshimage       As String '字幕トラック2用 ３段目 腰曲げの字幕文字列
    Dim CaptionName3Hizamage        As String '字幕トラック2用 ３段目 膝曲げの字幕文字列

    Dim ColorName1                  As String '字幕トラック1用 上段右側（信頼度 ）の色
    Dim ColorName2                  As String '字幕トラック1用 下段  （姿勢素点）の色
    Dim ColorName2Kobushiage        As String '字幕トラック2用 2段目 （拳上区間 ）の色
    Dim ColorName2Koshimage         As String '字幕トラック2用 2段目 （腰曲げデータ区間 ）の色
    Dim ColorName2Hizamage          As String '字幕トラック2用 2段目 （膝曲げデータ区間 ）の色
    Dim ColorName3Kobushiage        As String '字幕トラック2用 ３段目 （拳上 ）の色
    Dim ColorName3Koshimage         As String '字幕トラック2用 ３段目 （腰曲げ ）の色
    Dim ColorName3Hizamage          As String '字幕トラック2用 ３段目 （膝曲げ ）の色

    Dim Track1OutputString1         As String '字幕トラック1用：上段文字列
    Dim Track1OutputString2         As String '字幕トラック1用：下段文字列

    Dim Track2OutputString1         As String '字幕トラック2用：1段目文字列
    Dim Track2OutputString2         As String '字幕トラック2用：2段目文字列
    Dim Track2OutputString3         As String '字幕トラック2用：3段目文字列

    Dim Track1FileName              As String '字幕トラック1用のファイル名
    Dim Track2FileName              As String '字幕トラック2用のファイル名


    '表示・更新をオフにする
    Call stopUpdate

    '動画の縦横サイズを取得
    video_width = ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 198)
    video_height = ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 197) '動画の縦横判定のために高さも取得

    '動画の縦横によって係数を変更する
    '動画が縦の時
    If video_width < video_height Then
        track1_coef_font_size1 = TRACK1_TATE_UPPER_COEF  '動画が縦のときのトラック1用：上段
        track1_coef_font_size2 = TRACK1_TATE_LOWER_COEF
        track2_coef_font_size1 = TRACK2_TATE_1ST_COEF    'トラック2用：1段目
        track2_coef_font_size2 = TRACK2_TATE_2ND_COEF    'トラック2用：2段目
        track2_coef_font_size3 = TRACK2_TATE_3RD_COEF    'トラック2用：3段目
    '動画が横の時
    Else
        track1_coef_font_size1 = TRACK1_YOKO_UPPER_COEF  '動画が縦のときのトラック1用：上段
        track1_coef_font_size2 = TRACK1_YOKO_LOWER_COEF
        track2_coef_font_size1 = TRACK2_YOKO_1ST_COEF    'トラック2用：1段目
        track2_coef_font_size2 = TRACK2_YOKO_2ND_COEF    'トラック2用：2段目
        track2_coef_font_size3 = TRACK2_YOKO_3RD_COEF    'トラック2用：3段目
    End If

    'フォントサイズを設定
    track1_font_size1 = video_width / track1_coef_font_size1 '動画の縦or横によって分母を変更することで、文字サイズが変わる
    track1_font_size2 = video_width / track1_coef_font_size2
    track2_font_size1 = video_width / track2_coef_font_size1
    track2_font_size2 = video_width / track2_coef_font_size2
    track2_font_size3 = video_width / track2_coef_font_size3

    '各姿勢の名前と条件の読み出し
    'MinとMaxが直感的でないので注意
    With ThisWorkbook.Worksheets("条件設定シート")
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
    End With 'With ThisWorkbook.Worksheets("条件設定シート")

    '評価除外用
    CaptionName2(0) = "0-姿勢評価なし" '下段のキャプション名を表示しない

    Track2FileName = ActiveWorkbook.Path & "\" & movieName & ".srt"

    With ThisWorkbook.Sheets("ポイント計算シート")

        'ファイルを開く
        Open Track2FileName For Output As #2

        '処理する行数を取得
        max_row_num = getLastRow()

        'ファイル出力
        For i = 2 To max_row_num

            'ポイント計算シートのキャプション列より、姿勢重量点調査票の作業名を先に読み取っておく
            CaptionName0 = .Cells(i, COLUMN_CAPTION_WORK_NAME).Value

            '開発環境のシステムは複数字幕仕様のため、エラー防止のため、暫定で字幕２は空のファイルを作成する
            '自車用のシステムを作成する際には字幕ファイルを１つにする

            'ここから拳上
            '除外
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_REMOVE_SECTION
                ColorName2Kobushiage = COLOR_DATA_REMOVE_SECTION
            '強制
            ElseIf .Cells(i, COLUMN_KOBUSHIAGE_FORCED_SECTION).Value > 0 Then
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_FORCED_SECTION
                ColorName2Kobushiage = COLOR_DATA_FORCED_SECTION
            '欠損
            ElseIf .Cells(i, COLUMN_DATA_KOBUSHIAGE_MISSING_SECTION_DST).Value > 0 Then
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_MISSING_SECTION
                ColorName2Kobushiage = COLOR_DATA_MISSING_SECTION
            '測定
            Else
                CaptionName2Kobushiage = CAPTION_DATA_TRACK2_MEASURE_SECTION
                ColorName2Kobushiage = COLOR_DATA_MEASURE_SECTION
            End If

            'ここから腰曲げ
            '除外
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_REMOVE_SECTION
                ColorName2Koshimage = COLOR_DATA_REMOVE_SECTION
            '強制
            ElseIf .Cells(i, COLUMN_KOSHIMAGE_FORCED_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_FORCED_SECTION
                ColorName2Koshimage = COLOR_DATA_FORCED_SECTION
            '欠損
            ElseIf .Cells(i, COLUMN_DATA_KOSHIMAGE_MISSING_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_MISSING_SECTION
                ColorName2Koshimage = COLOR_DATA_MISSING_SECTION
            '測定
            ElseIf .Cells(i, COLUMN_DATA_KOSHIMAGE_MEASURE_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_MEASURE_SECTION
                ColorName2Koshimage = COLOR_DATA_MEASURE_SECTION
            '推定
            ElseIf .Cells(i, COLUMN_DATA_KOSHIMAGE_PREDICT_SECTION).Value > 0 Then
                CaptionName2Koshimage = CAPTION_DATA_TRACK2_PREDICT_SECTION
                ColorName2Koshimage = COLOR_DATA_PREDICT_SECTION
            End If

            'ここから膝曲げ
            '除外
            If .Cells(i, COLUMN_DATA_REMOVE_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_REMOVE_SECTION
                ColorName2Hizamage = COLOR_DATA_REMOVE_SECTION
            '強制
            ElseIf .Cells(i, COLUMN_HIZAMAGE_FORCED_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_FORCED_SECTION
                ColorName2Hizamage = COLOR_DATA_FORCED_SECTION
            '欠損
            ElseIf .Cells(i, COLUMN_DATA_HIZAMAGE_MISSING_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_MISSING_SECTION
                ColorName2Hizamage = COLOR_DATA_MISSING_SECTION
            '測定
            ElseIf .Cells(i, COLUMN_DATA_HIZAMAGE_MEASURE_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_MEASURE_SECTION
                ColorName2Hizamage = COLOR_DATA_MEASURE_SECTION
            '推定
            ElseIf .Cells(i, COLUMN_DATA_HIZAMAGE_PREDICT_SECTION).Value > 0 Then
                CaptionName2Hizamage = CAPTION_DATA_TRACK2_PREDICT_SECTION
                ColorName2Hizamage = COLOR_DATA_PREDICT_SECTION
            End If

            '3段目の描画色、キャプション名を設定する
            '拳上
            If .Cells(i, COLUMN_KOBUSHIAGE_RESULT).Value > 0 Then
                CaptionName3Kobushiage = "<b>" & CAPTION_A_RESULT_NAME1 & "</b>"
                ColorName3Kobushiage = COLOR_DATA_RESULT_RED
            Else
                CaptionName3Kobushiage = "<b>" & CAPTION_A_RESULT_NAME1 & "</b>"
                ColorName3Kobushiage = COLOR_DATA_RESULT_GLAY
            End If

            '腰曲げ
            If .Cells(i, COLUMN_KOSHIMAGE_RESULT).Value > 0 Then
                CaptionName3Koshimage = "<b>" & CAPTION_B_RESULT_NAME1 & "</b>"
                ColorName3Koshimage = COLOR_DATA_RESULT_RED
            Else
                CaptionName3Koshimage = "<b>" & CAPTION_B_RESULT_NAME1 & "</b>"
                ColorName3Koshimage = COLOR_DATA_RESULT_GLAY
            End If

            '膝曲げ
            If .Cells(i, COLUMN_HIZAMAGE_RESULT).Value > 0 Then
                CaptionName3Hizamage = "<b>" & CAPTION_C_RESULT_NAME1 & "</b>"
                ColorName3Hizamage = COLOR_DATA_RESULT_RED
            Else
                CaptionName3Hizamage = "<b>" & CAPTION_C_RESULT_NAME1 & "</b>"
                ColorName3Hizamage = COLOR_DATA_RESULT_GLAY
            End If

            '字幕文字列を決定
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

            '字幕文字列をポイント計算シートに出力
            'デバッグ用（普段は使わない）
            '.Cells(i, COLUMN_CAPTION_TRACK2).Value = Track1OutputString1 & Track1OutputString2

            'テキストファイルにその他字幕文字列を書き出し
            Print #2, "  " & i - 1 '数字の前に半角スペースを2個入れる。字幕トラック1と区別するため
            Print #2, .Cells(i, COLUMN_ROUGH_TIME).Value&; " --> " & .Cells(i + 1, COLUMN_ROUGH_TIME).Value '時刻を出力

            Print #2, Replace(Track2OutputString1, vbLf, vbCrLf) '改行コードを置き換え、キャプション出力
            Print #2, Replace(Track2OutputString2, vbLf, vbCrLf) '改行コードを置き換え、キャプション出力
            Print #2, Replace(Track2OutputString3, vbLf, vbCrLf) '改行コードを置き換え、キャプション出力

            Print #2, ""
            Print #2, ""

            '//
            '// 字幕トラック2用の処理 ここまで
            '////////////////////////////////////////

            'ポイント計算シートの字幕文字列 作業No. - 作業名をクリア
            .Cells(i, COLUMN_CAPTION_WORK_NAME).clear


            'デバッグ時、判定されない条件が分かるように色名をリセットしておく
            ColorName1 = "#ffffff"
            ColorName2 = "#ffffff"

        Next

        'ファイルを閉じる
        Close #1
        Close #2


    End With 'With ThisWorkbook.Sheets("ポイント計算シート")

    '表示・更新をオンに戻す
    Call restartUpdate

End Function


'帳票更新ボタンが押された時の処理
' 引数  ：なし
' 戻り値：なし
Function ClickUpdateDataCore()
    Dim tstart_click As Double
    Dim dotPoint     As String
    Dim workbookName As String
    Dim fps          As Double

    tstart_click = Timer
    fps = getFps()

    'ノイズ除去
    Call removeCaptionNoise(fps)

    '作業分割、時間測定
    Call fixSheetJisya

    dotPoint = InStrRev(ActiveWorkbook.Name, ".")
    workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)

    Call outputCaption(workbookName)
    Debug.Print " 更新時間" & Format$(Timer - tstart_click, "0.00") & " sec."

End Function


'帳票更新ボタンが押された時の処理
' 引数１：なし
' 戻り値：なし
Sub ClickUpdateData()
    Call ClickUpdateDataCore
End Sub

Sub ClickJisyaLimitChangeUpdateData()

    Dim response As VbMsgBoxResult

    ' メッセージボックスでユーザーに確認
    response = MsgBox("腰・膝角度のしきい値を変更して再評価します。姿勢評価修正シートの編集内容がリセットされますがよろしいですか？", vbOKCancel + vbQuestion, "確認")

    ' ユーザーの選択に応じて処理
    If response = vbOK Then
        ' OKボタンが押された場合、メッセージを表示
        Dim tstart_click As Double
        Dim dotPoint     As String
        Dim workbookName As String

        tstart_click = Timer

        dotPoint = InStrRev(ActiveWorkbook.Name, ".")
        workbookName = Left(ActiveWorkbook.Name, dotPoint - 1)

        '姿勢判定
        Call makeGraphJisya
        '作業分割、時間測定
        Call fixSheetJisya
        '修正シートの更新
        Call Module3.paintAll
        '字幕出力
        Call outputCaption(workbookName)
        '更新時間出力
        Debug.Print " 更新時間" & Format$(Timer - tstart_click, "0.00") & " sec."
        Sheets("工程評価シート").Activate
    Else
        ' キャンセルボタンが押された場合、何もしない
    End If
End Sub


' 概要 : 関節角度、3dデータのcsvをコピー貼り付けする
' 呼び元のシート : マクロテスト
' 補足 : 本ファイルと同じディレクトリにcsvファイルを置いておく
' 引数1 ：フレームレート
' 引数2 ：動画横幅の値
' 引数3 ：csvファイル名
' 引数4 ：動画縦の値 動画の向きによって字幕文字サイズを調整するために使用
' 戻り値：なし
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

    Sheets("ポイント計算シート").Select
    Range("D2").Select

    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & csv_file_name)

    With wb
        Set ws = .Sheets(1)

        Range("B2").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy

        'このブックの「貼付先」シートへ値貼り付け
        ThisWorkbook.Worksheets("ポイント計算シート").Range("D2").PasteSpecial _
            xlPasteValuesAndNumberFormats

        'コピー状態を解除
        Application.CutCopyMode = False

        '保存せず終了
        .Close False
    End With

    ' A から C の時間を表すセルを実体化させる
    ' angle.csvを張り付けたあとの最下行番号を取得する
    MaxRow = Range("D2").End(xlDown).row
    For i = 0 To MaxRow - 2
        Range("A" & i + 2).Value = i
        Range("B" & i + 2).Value = i * (1 / fps)
        Range("C" & i + 2).FormulaR1C1 = "=LEFT(TEXT(RC[-1]/(24*60*60), ""hh:mm:ss.000""), 8)"
    Next

    'fps値の保存
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 199) = fps
    'video_width値の保存
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 198) = video_width
    'video_height値の保存 動画の向きによって字幕文字サイズを調整するために使用
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 197) = video_height

    ThisWorkbook.Save

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


' 引数1 ：なし
' 引数2 ：なし
' 戻り値：なし
Sub VeryHiddenSheet()
    Sheets("ポイント計算シート").Visible = xlVeryHidden
    Sheets("条件設定シート").Visible = xlVeryHidden
End Sub


'Pythonから呼び出しされる
' 引数1 ：動画名
' 引数2 ：フレームレート
' 戻り値：なし
Sub MacroUpdateData(movieName As String, fps As Double)

    Dim tstart_first As Double

    If MEAGERE_TIME_MACROUPDATEDATA = True Then 'MacroUpdateDataの処理時間を測定する
        tstart_first = Timer
    End If

    With ThisWorkbook.Sheets("ポイント計算シート")
        Dim max_row_num As Long
        Dim i As Long

        '処理する行数を取得（3列目の最終セル）
        max_row_num = getLastRow()

        '★★★本処理は、将来的にPythonコード側で行う予定★★★
        'フラグが入力されるセルに入力されているスペースを検索して消去する
        'メイン字幕の姿勢素点の色が全て緑になる不具合の暫定対策
        'セル範囲が広すぎてメモリ不足になるため、for文で処理を細分化
        For i = 4 To 253
            .Range(.Cells(2, i), .Cells(max_row_num, i)).Replace " ", ""
        Next

        'fps値の保存
        fps = .Cells(2, 199)

    End With

    '姿勢判定
    Call makeGraphJisya


    'ノイズ除去
    Call removeCaptionNoise(fps)

    '作業分割、時間測定
    Call fixSheetJisya


    '修正シートの更新
    'Call Module1.paintAll
    Call Module3.paintAll

    '字幕生成
    Call outputCaption(movieName)

    'シートを隠す
    Call VeryHiddenSheet

    'MacroUpdateDataの処理時間を測定する
    If MEAGERE_TIME_MACROUPDATEDATA = True Then
        ThisWorkbook.Sheets("ポイント計算シート").Cells(2, COLUMN_MEAGERE_TIME_MACROUPDATEDATA) = Format$(Timer - tstart_first, "0.00")
    End If

    '初回分析済みのフラグを立てる
    ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 196) = 1

End Sub



'姿勢重量点調査票の選択と保存
' 引数1 ：動画名
' 戻り値：なし
Sub MacroSaveData(movieName As String)

    '工程評価シートを表示
    ThisWorkbook.Worksheets("工程評価シート").Select

    '工程評価シートの調査日を記入する
    ThisWorkbook.Worksheets("工程評価シート").Cells(GH_HYOUKA_SHEET_ROW_DATE, GH_HYOUKA_SHEET_COLUMN_KOUTEI_NAME).Value = Date

    '姿勢重量点調査票をアクティブにして保存する
    Sheets("工程評価シート").Activate
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

    '--- 拡張子を除いたパス（ファイル名）を格納する変数 ---'
    Dim strFileExExt As String

    If (0 < PosExt) Then
        StrFileName = Left(StrFileName, PosExt - 1)
    End If

    'Now関数で取得した現在日付をFormatで整形して変数に格納
    strYYYYMMDD = Format(Now, "yyyymmdd_HHMMSS_")

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    Set ReturnBook = ActiveWorkbook
    destFilePath = ActiveWorkbook.Path & "\" & StrFileName & "_otrs.xlsx"

    'もしotrs用ファイルがあれば、一度削除しておく
    If Dir(destFilePath) <> "" Then
        Kill destFilePath
    End If

    '作業用のワークブックのインスタンスを作る

    If Dir(destFilePath) = "" Then
        '新しいファイルを作成
        Set targetWorkbook = Workbooks.Add
        '新しいファイルをVBAを実行したファイルと同じフォルダ保存
        targetWorkbook.SaveAs destFilePath
    Else
        Set targetWorkbook = Workbooks.Open(destFilePath)
    End If

    ReturnBook.Activate
    lastPoseNum = -1
    lastTime = 0

    Dim CaptionName2(10) As String

    With ThisWorkbook.Worksheets("条件設定シート")
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
    End With

    CaptionName2(0) = "データなし"
        '以下のパターン以外はその他とする。
        '(10) 膝を曲げ上半身前屈(30°〜90°)
        '(9) 膝を曲げ上半身前屈(15°〜30°)
        '(8) 上半身前屈(45°〜90°)
        '(7) 上半身前屈(30°〜45°)
        '(6) 上半身前屈(90°〜180°)
        '(4) 蹲踞または片膝つき蹲踞
        '(2) 上半身前屈(15°〜30°)
        '(1) 基本の立ち姿勢
        '(0) 他"

    With ThisWorkbook.Sheets("ポイント計算シート")
        max_row_num = getLastRow()
        targetRowCount = 1
        Dim lastI As Long

        For i = 2 To max_row_num
            'COLUMN_DATA_RESULT_ORIGINが空白の可能性があるため一旦その他を入れておく
            writePoseNum = 0
            On Error Resume Next
            writePoseNum = .Cells(i, COLUMN_DATA_RESULT_ORIGIN).Value 'キャプション番号のセル代入

            '最初に別のポーズに変わった時が欲しいので一回目は同一にする。
            If i = 2 Then
                lastPoseNum = writePoseNum
                lastI = i - 2
            End If

            If lastPoseNum <> writePoseNum Then
                '同一ポーズを取っていた時間が必要（切り替わった一個前の時間）
                currentTime = .Cells(i - 1, 2).Value
                '書き込み処理
                targetWorkbook.Activate
                targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = CaptionName2(lastPoseNum)
                targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = Round(currentTime - lastTime, 5)
                lastI = i - 2
                targetRowCount = targetRowCount + 1

                lastTime = currentTime
                lastPoseNum = writePoseNum

                ReturnBook.Activate
            End If
        Next

        'ループ終了後に最後に取っていた姿勢が継続しているならそれを書き込む
        If lastPoseNum = writePoseNum Then
            currentTime = .Cells(i - 1, 2).Value
            '書き込み処理
            targetWorkbook.Activate
            targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_NAME).Value = CaptionName2(writePoseNum)
            targetWorkbook.Worksheets("Sheet1").Cells(targetRowCount, COLUMN_POSE_KEEP_TIME).Value = Round(currentTime - lastTime, 5)
            ReturnBook.Activate
        End If
    End With
    Sheet4.Activate
    ThisWorkbook.Save
    targetWorkbook.Close savechanges:=True
End Sub


Sub InputOtrs()

    Dim openFileName As String
    Dim lastTime     As Double

    Dim ex           As New Excel.Application    '// 処理用Excel
    Dim wb           As Workbook                 '// ワークブック
    Dim r            As Range                    '// 取得対象のセル範囲
    Dim sht          As Worksheet                '// 参照シート

    Dim i            As Long
    Dim max_row_num  As Long
    Dim max_row      As Long

    'OTRSエクスポートファイルを開く
    openFileName = Application.GetOpenFilename("OTRSエクスポートファイル,*.xlsx?")

    '正しくファイルが開かれた場合の処理
    If openFileName <> "False" Then

        '// 読み取り専用で開く
        Set wb = ex.Workbooks.Open(FileName:=openFileName, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)

        '処理する行数を取得（1列目の最終セル）
        max_row_num = wb.Worksheets(1).Cells(1, 1).End(xlDown).row

        For i = 2 To max_row_num
            '要素名のコピー、セル背景色は白にする
            ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 3).Value = wb.Worksheets(1).Cells(i, 1).Value
            ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 3).Interior.Color = RGB(255, 255, 255)

            '作業終了時間のコピー、セル背景色は白にする
            If i = 2 Then
                '処理する行数を取得
                max_row = getLastRow()
                '秒数を切り上げて代入
                lastTime = Application.WorksheetFunction.RoundUp(ThisWorkbook.Worksheets("ポイント計算シート").Cells(max_row, 2), 0)

            End If

            If i <> max_row_num Then
                ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 36).Value = "−"
                ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 37).Value = wb.Worksheets(1).Cells(i + 1, 2).Value
                ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 37).Interior.Color = RGB(255, 255, 255)
            Else
                ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 36).Value = "−"
                ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 37).Value = lastTime
                ThisWorkbook.Worksheets("姿勢重量点調査票").Cells(9 + i - 2, 37).Interior.Color = RGB(255, 255, 255)
            End If
        Next

        '// ブックを閉じる
        Call wb.Close

        '// Excelアプリケーションを閉じる
        Call ex.Application.Quit

        'データ更新
        ClickUpdateData

    End If
End Sub

'秒をhh:mm:ss:msに変換する
Function timeConvert(seconds As Double) As String

    Dim milliseconds        As Long
    Dim remainingSeconds    As Long
    Dim minutes             As Long
    Dim hours               As Long

    'ずれ防止のために小数点以下を切り捨てミリ秒・秒から先に出す
    milliseconds = (seconds - Int(seconds)) * 1000
    seconds = Int(seconds)

    remainingSeconds = seconds Mod 60
    minutes = (seconds Mod 3600) \ 60
    hours = seconds \ 3600

    timeConvert = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(remainingSeconds, "00") & "." & Format(milliseconds, "000")
End Function

'fpsの値を取得する
'戻り値：fpsの値
Function getFps() As Double
    getFps = ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 199).Value
End Function


'最終行を取得する
'戻り値：最終行
Function getLastRow() As Long
    getLastRow = ThisWorkbook.Sheets("ポイント計算シート").Cells(1, 3).End(xlDown).row
End Function