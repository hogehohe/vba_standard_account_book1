Option Explicit

'---------------------------------------------
'姿勢評価修正シートで使う定数
'---------------------------------------------
'1マスの秒数を定義
Const UNIT_TIME                         As Double = 0.1

'0秒の列
Const COLUMN_ZERO_NUM                   As Long = 6

'行
'信頼性上端
Const ROW_RELIABILITY_TOP               As Long = 2
'信頼性下端
Const ROW_RELIABILITY_BOTTOM            As Long = 7
'姿勢点上端
Const ROW_POSTURE_SCORE_TOP             As Long = 9
'姿勢点下端
Const ROW_POSTURE_SCORE_KOSHIMAGEOTTOM  As Long = 17

'A_姿勢点
Const ROW_POSTURE_SCORE_KOBUSHIAGE      As Long = 12
'B_姿勢点
Const ROW_POSTURE_SCORE_KOSHIMAGE       As Long = 14
'C_姿勢点
Const ROW_POSTURE_SCORE_HIZAMAGE        As Long = 16

'---------------------------------------------
'ポイント計算シートの列
'---------------------------------------------

'姿勢点が保存されている列
Const COLUMN_POSTURE_SCORE_ALL          As Long = 203

'測定
Const COLUMN_MEASURE_SECTION            As Long = 204
'推定
Const COLUMN_PREDICT_SECTION            As Long = 205
'除外区間
Const COLUMN_REMOVE_SECTION             As Long = 206
'強制区間
Const COLUMN_FORCED_SECTION_TOTAL       As Long = 207
'元データ
Const COLUMN_BASE_SCORE                 As Long = 208
'姿勢素点緑色
Const COLUMN_POSTURE_GREEN              As Long = 209
'姿勢素点黄色
Const COLUMN_POSTURE_YELLOW             As Long = 210
'姿勢素点赤色
Const COLUMN_POSTURE_RED                As Long = 211

'欠損
Const COLUMN_MISSING_SECTION            As Long = 219

'拳上強制区間
Const COLUMN_FORCED_SECTION_KOBUSHIAGE  As Long = 223
'腰曲げ強制区間
Const COLUMN_FORCED_SECTION_KOSHIMAGE   As Long = 228
'膝曲げ強制区間
Const COLUMN_FORCED_SECTION_HIZAMAGE    As Long = 233

'条件A(拳上)が保存されている列
Const COLUMN_POSTURE_SCORE_KOBUSHIAGE   As Long = 245
'条件B(腰曲げ)が保存されている列
Const COLUMN_POSTURE_SCORE_KOSHIMAGE    As Long = 247
'条件C(膝曲げ)が保存されている列
Const COLUMN_POSTURE_SCORE_HIZAMAGE     As Long = 249

'---------------------------------------------
'姿勢評価修正シート　関連
'---------------------------------------------
'LIMIT_COLUMNの設定値は3の倍数とする必要がある
'30fps×60秒×9分＝16200
'姿勢評価修正シートは9分毎に次のシートを生成する
Const LIMIT_COLUMN                      As Long = 16200

Const SHEET_LIMIT_COLUMN                As Long = LIMIT_COLUMN + COLUMN_ZERO_NUM

'時刻表示セルの幅
Const TIME_WIDTH                        As Long = 30
'時刻表示セルが存在する行
Const TIME_ROW                          As Long = 22
'一つ目の時刻表示セルの左端
Const TIME_COLUMN_LEFT                  As Long = 22
'一つ目の時刻表示セルの右端
Const TIME_COLUMN_RIGHT                 As Long = 51
'データ調整用のテーブルの下端
Const BOTTOM_OF_TABLE                   As Long = 22

'列幅用の列挙
Private Enum widthSize
    Small = 1
    Medium = 2
    Large = 4
    LL = 6
End Enum

'列幅調整ボタン名前
Const EXPANDBTN_NAME                    As String = "expandBtn"
Const REDUCEBTN_NAME                    As String = "reduceBtn"

'---------------------------------------------
'複数モジュールで使用する変数
'---------------------------------------------
'再生・停止ボタンで使用
'指定した時間が経過すると処理を実行する
Private ResTime As Date
Private scrollTime As Double


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


'姿勢評価修正シート
'罫線と映像時刻をオートフィル
Sub autoFillTemplate()
    'ラベルの列数
    Dim startColumnNum      As Long
    '10秒の列数
    Dim unit10SecColumnNum  As Long

    '変数定義
    Dim workTime            As Double
    Dim fps                 As Double
    Dim maxFrameNum         As Long
    Dim ruleLineColumnNum   As Long
    Dim ruleLineColumnAlf   As String

    startColumnNum = COLUMN_ZERO_NUM + 1
    unit10SecColumnNum = 10 / UNIT_TIME

    '作業時間を取得する
    With ThisWorkbook.Sheets("ポイント計算シート")
        'フレームレートを取得
        fps = getFps()
        '最終行の値を取得
        maxFrameNum = getLastRow()
    End With
End Sub


'------------------------------------------------------------
' 罫線の複製処理
'
' 「G2:EZ21」のレイアウトをベースとして、指定列まで右方向に複製する。
' 主に罫線や装飾を含むオートフィル処理を実行。
'
' 引数:
'   ws       - 対象のワークシート
'   endline  - 罫線を引く対象の最終列（上限あり）
'------------------------------------------------------------
Private Sub autoFillLine(ws As Worksheet, endline As Long)
    Dim targetColNum   As Long   ' 処理対象の列番号（制限後）
    Dim targetColAlpha As String ' 列番号をアルファベット表記に変換したもの

    ' 列番号が上限を超える場合、制限値までに調整
    targetColNum = endline
    If targetColNum > SHEET_LIMIT_COLUMN Then
        targetColNum = SHEET_LIMIT_COLUMN
    End If

    ' オートフィル終了列をアルファベット表記に変換（例: 160 → "FT"）
    targetColAlpha = Split(ws.Cells(1, targetColNum).Address(True, False), "$")(0)

    ' 対象シートを一度クリア（色や罫線含め全消去）
    Call clear(ws)

    ' レイアウトの基準範囲（G2:EZ21）を、指定列までオートフィル
    ws.Range("G2:EZ21").AutoFill _
        Destination:=ws.Range("G2:" & targetColAlpha & 21), _
        Type:=xlFillDefault

    ' 複製された範囲の右側の不要な罫線を削除（XFD列まで）
    Dim clearStartColAlpha As String
    clearStartColAlpha = Split(ws.Cells(1, targetColNum + 1).Address(True, False), "$")(0)
    ws.Range(clearStartColAlpha & "2:XFD21").Borders.LineStyle = xlLineStyleNone

End Sub


'------------------------------------------------------------
' 時刻を時間セルに挿入する処理
'
' 引数:
'   ws      - 対象のワークシート
'   min     - 分単位（例: 5 → 00:05:01のように開始）
'   endclm  - 処理対象の最終列
'------------------------------------------------------------
Private Sub autoFillTime(ws As Worksheet, min As Long, endclm As Long)

    ' 変数定義
    Dim tmp         As Long: tmp = endclm
    Dim boldcnt     As Long: boldcnt = 0
    Dim r           As Range
    Dim timeStr     As String
    Dim frame30Mod  As Long
    Dim i           As Long

    ' 最終列の調整（上限制限を考慮）
    If 30 <= tmp - TIME_COLUMN_LEFT Then
        If tmp > LIMIT_COLUMN Then
            tmp = LIMIT_COLUMN
        End If
    End If

    ' 結合セルがあるとオートフィル時にエラーになるため、事前に解除・クリア
    ws.Range(ws.Cells(TIME_ROW, 12), ws.Cells(TIME_ROW, 16384)).Clear

    ' 時間セルの書式設定と結合処理
    For i = TIME_COLUMN_LEFT To SHEET_LIMIT_COLUMN Step TIME_WIDTH
        Set r = ws.Range(ws.Cells(TIME_ROW, i), ws.Cells(TIME_ROW, i + TIME_WIDTH - 1))

        boldcnt = boldcnt + 1

        With r
            .Merge True                      ' セル結合（横方向）
            .Orientation = -90              ' 縦書き（90度回転）
            .ReadingOrder = xlContext       ' 文字方向：自動判定
            .HorizontalAlignment = xlCenter ' 横位置：中央
            .NumberFormatLocal = "hh:mm:ss" ' 時刻形式にする

            ' 5回に1回は太字にする
            If boldcnt = 5 Then
                .Font.FontStyle = "bold"
                boldcnt = 0
            End If
        End With
    Next i

    ' 初期の2つの時刻を直接入力（例: 00:05:01, 00:05:02）
    timeStr = "00:" & Format(min, "00") & ":01"
    ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT), _
             ws.Cells(TIME_ROW, TIME_COLUMN_RIGHT)).Value = timeStr

    timeStr = "00:" & Format(min, "00") & ":02"
    ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT + TIME_WIDTH), _
             ws.Cells(TIME_ROW, TIME_COLUMN_RIGHT + TIME_WIDTH)).Value = timeStr

    ' フレーム幅の余りを調整し、オートフィル範囲をフレーム単位に丸める
    frame30Mod = (tmp - TIME_COLUMN_LEFT) Mod TIME_WIDTH
    If frame30Mod Then
        tmp = tmp + TIME_WIDTH - frame30Mod
    End If

    ' 2つ目の時刻より右側が存在する場合にのみ、オートフィルを実行
    If (TIME_COLUMN_LEFT + TIME_WIDTH) < tmp Then
        ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT), _
                 ws.Cells(TIME_ROW, TIME_COLUMN_RIGHT + TIME_WIDTH)).AutoFill _
            Destination:=ws.Range(ws.Cells(TIME_ROW, TIME_COLUMN_LEFT), _
                                  ws.Cells(TIME_ROW, tmp - 1)), _
            Type:=xlFillValues
    End If

End Sub


'単位時間当たり最も多い姿勢点・信頼性を調べてセルに色を塗る
'processingRange　1:選択範囲（部分的に強制をキャンセル） 2:全体 else:特定の1セルごと
Sub paintPostureScore(processingRange As Long)
    '---------------------------------------------
    'RGBを指定するための変数を定義
    '---------------------------------------------
    '信頼性
    Dim colorMeasureSection    As Long: colorMeasureSection = RGB(0, 176, 240)   '水色
    Dim colorPredictSection    As Long: colorPredictSection = RGB(252, 246, 0)   '黄色
    Dim colorMissingSection    As Long: colorMissingSection = RGB(255, 124, 128) 'ピンク
    Dim colorForcedSection     As Long: colorForcedSection  = RGB(0, 51, 204)    '青色
    Dim colorRemoveSection     As Long: colorRemoveSection  = RGB(191, 191, 191) 'グレー

    '姿勢点
    Dim colorResultGreen       As Long: colorResultGreen    = RGB(0, 176, 80)    '緑色
    Dim colorResultYellow      As Long: colorResultYellow   = RGB(255, 192, 0)   '黄色
    Dim colorResultRed         As Long: colorResultRed      = RGB(192, 0, 0)     '赤色
    Dim colorResultGlay        As Long: colorResultGlay     = RGB(191, 191, 191) 'グレー
    Dim colorResultWhite       As Long: colorResultWhite    = RGB(255, 255, 255) '白色 20221219_下里
    Dim colorResultBrown       As Long: colorResultBrown    = RGB(64, 0, 0)      '茶色 20221222_下里
    Dim colorResultOFFGlay     As Long: colorResultOFFGlay  = RGB(217, 217, 217) 'グレー 20221222_下里

    '---------------------------------------------
    '配列
    '---------------------------------------------
    'ポイント計算シートの姿勢点を保管
    Dim postureScoreDataArray()    As Long

    '条件A,B,Cごとの姿勢点を保管
    Dim postureScoreDataArray_A()  As Long
    Dim postureScoreDataArray_B()  As Long
    Dim postureScoreDataArray_C()  As Long

    '0～1点のフレーム数をそれぞれ合計

    '条件AからCのフレーム数をそれぞれ合計
    Dim postureScoreCounterArray_A(0 To 1)      As Long
    Dim postureScoreCounterArray_B(0 To 1)      As Long
    Dim postureScoreCounterArray_C(0 To 1)      As Long

    'ポイント計算シートの信頼性を保管
    '1:測定、2:推定、3:欠損
    Dim reliabilityDataArray()     As Long
    '信頼性 1 ~ 3 のフレーム数をそれぞれ合計
    Dim reliabilityCounterArray(3) As Long

    '---------------------------------------------
    'その他の変数
    '---------------------------------------------
    'ポイント計算シート最大行数の変数定義
    Dim maxRowNum               As Long

    '変数定義
    Dim i                       As Long
    Dim j                       As Long
    Dim fps                     As Double

    '単位時間の繰り返し処理の開始終了地点を定義
    Dim wholeStart              As Long
    Dim wholeEnd                As Long

    '姿勢点一時記憶用の変数
    Dim postureScoreFlag        As Long

    Dim postureScoreFlag_A      As Long
    Dim postureScoreFlag_B      As Long
    Dim postureScoreFlag_C      As Long

    '単位時間の中で一番多い姿勢点を保管
    Dim mostOftenPostureScore   As Long

    'AからCの姿勢点一時記憶用の変数
    Dim mostOftenPostureScore_A As Long
    Dim mostOftenPostureScore_B As Long
    Dim mostOftenPostureScore_C As Long

    '信頼性一時記憶用の変数
    Dim reliabilityFlag         As Long
    '単位時間の中で一番多い信頼性を保管
    Dim mostOftenReliability    As Long

    '次ページにいく制限
    Dim thisPageLimit           As Long: thisPageLimit = LIMIT_COLUMN

    '前のページの最終列を保存する
    Dim preClm  As Long: preClm = 0

    Call stopUpdate

    Dim baseClm     As Long
    Dim shtPage     As Long

    'カラーを保持する変数
    Dim colorStr    As String
    Dim colorStr1   As String '条件A
    Dim colorStr2   As String '条件B
    Dim colorStr3   As String '条件C

    '動画時間(秒)により列の初期幅を変更する

    Dim wSize       As widthSize

    '---------------------------------------------
    '変数、配列に値を入力
    '---------------------------------------------
    With ThisWorkbook.Sheets("ポイント計算シート")
        '最終行を取得
        maxRowNum = getLastRow()
        '配列の最後尾
        '余分を削除
        maxRowNum = maxRowNum - 1

        '配列を再定義
        ReDim postureScoreDataArray_A(maxRowNum, 0)
        ReDim postureScoreDataArray_B(maxRowNum, 0)
        ReDim postureScoreDataArray_C(maxRowNum, 0)

        '信頼性区間用
        ReDim reliabilityDataArray(maxRowNum, 0)

        '配列の中に値を入れる
        For i = 1 To maxRowNum

            '条件Aからの配列を入れる
            postureScoreDataArray_A(i - 1, 0) = .Cells(i + 1, COLUMN_POSTURE_SCORE_KOBUSHIAGE - 1).Value
            postureScoreDataArray_B(i - 1, 0) = .Cells(i + 1, COLUMN_POSTURE_SCORE_KOSHIMAGE - 1).Value
            postureScoreDataArray_C(i - 1, 0) = .Cells(i + 1, COLUMN_POSTURE_SCORE_HIZAMAGE - 1).Value

            '信頼性を配列に入れる
            '1:測定、2:推定、3:欠損

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
        'フレームレートを取得
        fps = getFps()
        Dim video_sec As Double: video_sec = wholeEnd / fps

    End With

    '---------------------------------------------
    '処理範囲を決める
    '---------------------------------------------
    'キャンセル(戻る)ボタンから呼ばれたとき

    If processingRange = 1 Then
        'アクティブセルの一番左が6列目以下の時
        'エラーメッセージを出して処理をやめる

        shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
        baseClm = LIMIT_COLUMN * shtPage

        'pageLimitを次のページとなる閾値まで更新
        thisPageLimit = (shtPage + 1) * LIMIT_COLUMN
        preClm = (LIMIT_COLUMN * shtPage) * -1

        Dim lCol As Long, rCol As Long
        If Not CropSelectionToDataArea(lCol, rCol) Then
            MsgBox "範囲外です", vbCritical
            Exit Sub
        End If

        wholeStart = lCol - COLUMN_ZERO_NUM + baseClm
        wholeEnd = rCol - COLUMN_ZERO_NUM + baseClm

        If wholeStart < 1 Then
            wholeStart = 1
        End If

    'メインの処理から呼ばれたとき
    ElseIf processingRange = 2 Then

        '先頭から
        wholeStart = 1
        '末尾まで
        wholeEnd = maxRowNum

        '基準のworkSheet、合わせて初期化
        ThisWorkbook.Sheets("姿勢評価修正シート").Activate
        preClm = 0
        If maxRowNum >= 150 Then
            Call autoFillLine(ActiveSheet, wholeEnd + COLUMN_ZERO_NUM) '230206 + COLUMN_ZERO_NUMを追加
            Call autoFillTime(Worksheets("姿勢評価修正シート"), 0, wholeEnd)
        End If

        Call addPageShape(ActiveSheet, False, True)

        '15秒以下を列幅2とする
        If video_sec <= 15 Then
            wSize = LL
            Call changeBtnState(EXPANDBTN_NAME, False)
            Call changeBtnState(REDUCEBTN_NAME, True)
        Else
            wSize = Small
            Call changeBtnState(REDUCEBTN_NAME, False)
            Call changeBtnState(EXPANDBTN_NAME, True)
        End If

        Call DataAjsSht.SetCellsHW(CInt(wSize), Worksheets("姿勢評価修正シート"))

    '除外があるフレームに強制を上書きしたとき（1セルずつ実行）
    Else
        shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
        baseClm = LIMIT_COLUMN * shtPage

        'pageLimitを次のページとなる閾値まで更新
        thisPageLimit = (shtPage + 1) * LIMIT_COLUMN
        preClm = (LIMIT_COLUMN * shtPage) * -1

        wholeStart = processingRange - COLUMN_ZERO_NUM + baseClm

        wholeEnd = wholeStart
    End If

    For i = wholeStart To wholeEnd

        '姿勢点のカウンターをリセット
        'AからCの姿勢点のカウンターをリセット
        Erase postureScoreCounterArray_A
        Erase postureScoreCounterArray_B
        Erase postureScoreCounterArray_C
'
        '信頼性のカウンターをリセット
        Erase reliabilityCounterArray

        '姿勢点を確認
        'AからCの姿勢点を確認
        postureScoreFlag_A = postureScoreDataArray_A(i - 1, 0)
        postureScoreFlag_B = postureScoreDataArray_B(i - 1, 0)
        postureScoreFlag_C = postureScoreDataArray_C(i - 1, 0)

        '姿勢点フラグを立てる
        'AからCの姿勢点フラグを立てる
        postureScoreCounterArray_A(postureScoreFlag_A) = 1
        postureScoreCounterArray_B(postureScoreFlag_B) = 1
        postureScoreCounterArray_C(postureScoreFlag_C) = 1

        '信頼性を確認
        reliabilityFlag = reliabilityDataArray(i, 0)
        '信頼性フラグを立てる
        reliabilityCounterArray(reliabilityFlag) = 1

        '---------------------------------------------
        'フレーム数が最も多いものを探す
        '---------------------------------------------
        mostOftenPostureScore = 0
        mostOftenPostureScore_A = 0
        mostOftenPostureScore_B = 0
        mostOftenPostureScore_C = 0

        '姿勢点 0 ~ 1 の先頭から順に比較
        For j = 0 To 1
            'フレーム数の合計が多い姿勢点を選ぶ
            '合計が同じ場合は辛い姿勢を優先する

            '条件AからC
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

        '初期は1
        mostOftenReliability = 1
            '信頼性1～3の先頭から順に比較
            '1:測定、2:推定、3:欠損
        For j = 2 To 3
            'フレーム数の合計が多い姿勢点を選ぶ
            '合計が同じ場合は信頼性が低い方を優先する
            If reliabilityCounterArray(mostOftenReliability) <= reliabilityCounterArray(j) Then
                mostOftenReliability = j
            End If
        Next

        'active sheetを変更する基準
        If i <= thisPageLimit Then
            '何もしない
        Else
            ThisWorkbook.ActiveSheet.Next.Activate
            If InStr(ThisWorkbook.ActiveSheet.Name, "姿勢評価修正シート") > 0 Then
                '何もしない
            Else
                '戻る
                ThisWorkbook.ActiveSheet.Previous.Activate
                Call createSheet(wholeEnd - i)
            End If
            '更新
            thisPageLimit = thisPageLimit + LIMIT_COLUMN
            preClm = preClm - LIMIT_COLUMN
            Call clear(ActiveSheet)
            Call autoFillLine(ActiveSheet, wholeEnd - i)
            Call autoFillTime(ThisWorkbook.ActiveSheet, ((thisPageLimit / LIMIT_COLUMN) - 1) * 9, wholeEnd - i)
            Call addPageShape(ActiveSheet, True, True)
        End If
        '---------------------------------------------
        '姿勢評価修正シートのセルに色を塗る
        '---------------------------------------------
        With ThisWorkbook.ActiveSheet

            '-------------条件A
            '0点の場合、白
            If mostOftenPostureScore_A = 0 Then
                colorStr1 = colorResultWhite

            '1点の場合、赤
            ElseIf mostOftenPostureScore_A = 1 Then
                colorStr1 = colorResultRed
            End If

'            -------------条件B
            '0点の場合、白
            If mostOftenPostureScore_B = 0 Then
                colorStr2 = colorResultWhite

            '1点の場合、赤
            ElseIf mostOftenPostureScore_B = 1 Then
                colorStr2 = colorResultRed
            End If

'            -------------条件C
            '0点の場合、白
            If mostOftenPostureScore_C = 0 Then
                colorStr3 = colorResultWhite

            '1点の場合、赤
            ElseIf mostOftenPostureScore_C = 1 Then
                colorStr3 = colorResultRed
            End If

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            '色をクリア
            .Range _
            ( _
                .Cells(ROW_POSTURE_SCORE_KOSHIMAGEOTTOM, COLUMN_ZERO_NUM + i + preClm), _
                .Cells(ROW_POSTURE_SCORE_TOP, COLUMN_ZERO_NUM + i + preClm) _
            ) _
            .Interior.ColorIndex = 0

            '~~~~~~~~~~~~~~~追加~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '色を塗る
            '条件A
            If mostOftenPostureScore_A = 0 Or 1 Then
                .Range _
                ( _
                    .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, COLUMN_ZERO_NUM + i + preClm), _
                    .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, COLUMN_ZERO_NUM + i + preClm) _
                ) _
                .Interior.Color = colorStr1
            End If

            '条件B
            If mostOftenPostureScore_B = 0 Or 1 Then

                .Range _
                ( _
                    .Cells(ROW_POSTURE_SCORE_KOSHIMAGE, COLUMN_ZERO_NUM + i + preClm), _
                    .Cells(ROW_POSTURE_SCORE_KOSHIMAGE, COLUMN_ZERO_NUM + i + preClm) _
                ) _
                .Interior.Color = colorStr2
            End If

            '条件C
            If mostOftenPostureScore_C = 0 Or 1 Then

                .Range _
                ( _
                    .Cells(ROW_POSTURE_SCORE_HIZAMAGE, COLUMN_ZERO_NUM + i + preClm), _
                    .Cells(ROW_POSTURE_SCORE_HIZAMAGE, COLUMN_ZERO_NUM + i + preClm) _
                ) _
                .Interior.Color = colorStr3
            End If

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            '最も多かった信頼性に応じて
            '色を変更
            '1:測定、2:推定、3:欠損
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

    ' キャンセルボタン以外からの処理の時
    If 1 < processingRange Then
        If calcSheetNamePlace(ThisWorkbook.ActiveSheet) = 0 Then ' 0 = Base sheet
            Call addPageShape(ActiveSheet, False, False)
        Else
            Call addPageShape(ActiveSheet, True, False)
        End If
    End If

    '各シートを更新
    Call checkReliabilityRatio
    Call restartUpdate

End Sub


'『全体を処理』ボタンが押されたとき
'全体の姿勢点を計算して、色を塗る
Sub paintAll()
    Call paintPostureScore(2)
End Sub


'『Cancel』ボタンが押されたとき
'選択範囲の姿勢点を計算して、色を塗る（強制ボタンのキャンセル）
Sub paintSelected()

    ' 選択範囲の左端の列が「0列目（＝姿勢点列）」以下なら処理をスキップ
    If DataAjsSht.activeCells <= COLUMN_ZERO_NUM Then
        Exit Sub
    End If

    ' 選択範囲のみ再描画
    Call paintPostureScore(1)
End Sub


'塗りつぶしを全てクリア
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


'結果の修正ボタン
'姿勢点を強制的に変更する
'ボタン別で引数postureScorebuttonが変わる
Sub forceResult(postureScorebutton As Long)
    '---------------------------------------------
    'RGBを指定するための変数を定義
    '---------------------------------------------
    ' 色設定：信頼性
    Dim colorMeasureSection As Long: colorMeasureSection = RGB(0, 176, 240)
    Dim colorPredictSection As Long: colorPredictSection = RGB(252, 246, 0)
    Dim colorMissingSection As Long: colorMissingSection = RGB(255, 124, 128)
    Dim colorForcedSection  As Long: colorForcedSection  = RGB(0, 51, 204)
    Dim colorResetSection   As Long: colorResetSection   = RGB(191, 191, 191)

    ' 色設定：姿勢点
    Dim colorResultGreen    As Long: colorResultGreen   = RGB(0, 176, 80)
    Dim colorResultYellow   As Long: colorResultYellow  = RGB(255, 192, 0)
    Dim colorResultRed      As Long: colorResultRed     = RGB(192, 0, 0)
    Dim colorResultGlay     As Long: colorResultGlay    = RGB(191, 191, 191)
    Dim colorResultWhite    As Long: colorResultWhite   = RGB(255, 255, 255)
    Dim colorResultOFFGlay  As Long: colorResultOFFGlay = RGB(217, 217, 217)

    ' 現在のシート位置から列オフセットを計算
    Dim shtPage             As Long: shtPage = calcSheetNamePlace(ThisWorkbook.ActiveSheet)
    Dim baseClm             As Long: baseClm = LIMIT_COLUMN * shtPage

    '選択範囲内のセル読み込み用　20221222_下里
    '変数定義
    Dim SelectCells     As Variant
    Dim k               As Long
    Dim m               As Long
    Dim lCol            As Long
    Dim rCol            As Long
    Dim MaxRightCell    As Variant
    Dim MinLeftCell     As Variant

    '一時的にSelection.rowの価を保存しておく変数
    Dim postur_row      As Long

    '---------------------------------------------
    'ここから強制処理
    '---------------------------------------------
    With ThisWorkbook.ActiveSheet
        '修正シートの選択範囲はポイント計算シートからはみ出さない範囲にあること
        '修正シートの選択範囲は色塗りできる範囲にあること
        If CropSelectionToDataArea(lCol, rCol) Then

            '選択範囲の左端と右端を取得
            MinLeftCell = lCol
            MaxRightCell = rCol

            '戻る(Removeボタン)
            If postureScorebutton = -1 Then
                Call postureUpdate(MinLeftCell + baseClm, MaxRightCell + baseClm, 0, CInt(postureScorebutton))
                '下が今まで戻るボタンを押したときにキックされるマクロ
                Call paintPostureScore(1)

            '強制(0 ~ 11の姿勢点ボタン)
            ElseIf postureScorebutton >= 0 Then

                Call postureUpdate(MinLeftCell + baseClm, MaxRightCell + baseClm, 1, CInt(postureScorebutton))

                If postureScorebutton = 99 Then
                    '除外を99に変更　20221219_下里
                    '最初に背景塗りつぶし無しにしているので、処理をしない
                    '信頼性のセルに除外の色を塗る
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

                    '===強制時、ほかの列の除外を解除し、除外だった場所に元のデータどおりに色を付けなおす処理===

                    For m = MinLeftCell To MaxRightCell
                        If .Cells(ROW_POSTURE_SCORE_KOBUSHIAGE, m).Interior.Color = colorResultGlay Then
                            Call paintPostureScore(m)
                        End If
                    Next
                    '====================================================================================

                    '姿勢点のセルに押されたボタンの姿勢点
                    '1点の場合は赤
                    If postureScorebutton = 1 Then
                        .Range _
                        ( _
                            .Cells(postur_row, MinLeftCell), _
                            .Cells(postur_row, MaxRightCell) _
                        ) _
                        .Interior.Color = colorResultBrown

                    '0点の場合は白
                    ElseIf postureScorebutton = 0 Then
                        .Range _
                        ( _
                            .Cells(postur_row, MinLeftCell), _
                            .Cells(postur_row, MaxRightCell) _
                        ) _
                        .Interior.Color = colorResultOFFGlay
                    End If

                    '信頼性のセルに強制色をぬる
                    .Range _
                    ( _
                        .Cells(ROW_RELIABILITY_TOP, MinLeftCell), _
                        .Cells(ROW_RELIABILITY_BOTTOM, MaxRightCell) _
                    ) _
                    .Interior.Color = colorForcedSection

                End If
                '強制のときは単独で実行
                Call checkReliabilityRatio
            End If
        Else
            MsgBox "範囲はグラフ内から選択してください", vbOKOnly + vbCritical, "範囲選択エラー"
        End If
    End With

    Call checkReliabilityRatio

End Sub


'------------------------------------------------------------
' ポイント計算シートの姿勢点・信頼性を更新
'
' 【引数】
'   sclm  - 選択範囲の左端列番号（実際の列位置）
'   fclm  - 選択範囲の右端列番号（実際の列位置）
'   bit   - 初期化フラグ（0:初期化 / 1:強制）※未使用？
'   score - 姿勢点（-1:初期化, 1?10:強制, 99:除外）
'
' 【処理概要】
'   - 選択行に応じて「姿勢スコアの列番号（強制）」を決定
'   - 指定範囲の各列について姿勢点を更新
'   - 信頼性の更新も実施
'------------------------------------------------------------
Sub postureUpdate(sclm As Long, fclm As Long, bit As Long, score As Long)

    Dim s                As Long
    Dim last             As Long
    Dim i                As Long
    Dim fbit             As Long
    Dim vle              As Long
    Dim column_forced_num As Long

    ' 選択行に応じて強制スコアの対象列番号を設定
    If Selection.row = ROW_POSTURE_SCORE_KOBUSHIAGE Then
        column_forced_num = COLUMN_POSTURE_SCORE_KOBUSHIAGE
    ElseIf Selection.row = ROW_POSTURE_SCORE_KOSHIMAGE Then
        column_forced_num = COLUMN_POSTURE_SCORE_KOSHIMAGE
    ElseIf Selection.row = ROW_POSTURE_SCORE_HIZAMAGE Then
        column_forced_num = COLUMN_POSTURE_SCORE_HIZAMAGE
    End If

    ' データ列へのオフセット変換（データは2行目から）
    s = sclm - COLUMN_ZERO_NUM + 1
    last = fclm - COLUMN_ZERO_NUM + 1

    ' 姿勢点の更新ループ（s ～ last 列）
    For i = s To last
        With ThisWorkbook.Sheets("ポイント計算シート")

            ' 初期化フラグに応じたスコア決定
            fbit = .Cells(i, COLUMN_FORCED_SECTION_TOTAL).Value

            If bit = 0 Then ' 初期化処理
                If fbit = 0 Then
                    vle = .Cells(i, COLUMN_POSTURE_SCORE_ALL).Value
                Else
                    vle = .Cells(i, COLUMN_BASE_SCORE).Value
                End If

                ' 姿勢素点除外区間（除外フラグあり）
                If .Cells(i, COLUMN_REMOVE_SECTION).Value = 1 Then
                    vle = .Cells(i, COLUMN_BASE_SCORE).Value
                End If

            Else ' 強制設定
                vle = score
            End If

            ' 基礎スコア処理（呼び出し）および更新
            Call baseScore(i, bit)
            .Cells(i, COLUMN_POSTURE_SCORE_ALL).Value = vle

        End With

        ' 信頼性の更新
        Call reliabilityUpdate(i, bit, vle, column_forced_num)

    Next

End Sub


'------------------------------------------------------------
' 姿勢素点の信頼性を更新（強制／除外／リセット）
'
' 【引数】
'   row               - 対象となる行番号（ポイント計算シート）
'   bit               - 呼び出し種別フラグ（0:リセット, 1:強制）
'   vle               - 姿勢スコア（-1:初期化, 1?10:強制, 99:除外）
'   column_forced_num - 姿勢スコアを入力する対象列（拳上・腰曲げ・膝曲げ）
'
' 【呼び出し例】
'   - bit = 0   → リセット（姿勢スコア・信頼性の初期化）
'   - bit = 1   → 強制（vleにスコア、または99:除外）
'------------------------------------------------------------
Sub reliabilityUpdate(row As Long, bit As Long, vle As Long, column_forced_num As Long)
    '変数定義
    Dim column_reliability_forced_num As Long

    With ThisWorkbook.Sheets("ポイント計算シート")
        '除外
        If vle = 99 And bit = 1 Then
            '姿勢素点除外区間
            .Cells(row, COLUMN_REMOVE_SECTION).Value = bit

            '姿勢の強制を解除
            .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = 0
            .Cells(row, COLUMN_POSTURE_SCORE_KOSHIMAGE).Value = 0
            .Cells(row, COLUMN_POSTURE_SCORE_HIZAMAGE).Value = 0

            '信頼性の強制を解除
            .Cells(row, COLUMN_FORCED_SECTION_KOBUSHIAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_KOSHIMAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_HIZAMAGE).Value = 0

        'リセット
        ElseIf bit = 0 Then
            '姿勢素点強制区間と姿勢素点除外区間のビットを消す
            .Cells(row, COLUMN_FORCED_SECTION_TOTAL).Value = bit
            .Cells(row, COLUMN_REMOVE_SECTION).Value = bit

            '姿勢をリセット
            .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE).Value = .Cells(row, COLUMN_POSTURE_SCORE_KOBUSHIAGE - 1).Value
            .Cells(row, COLUMN_POSTURE_SCORE_KOSHIMAGE).Value = .Cells(row, COLUMN_POSTURE_SCORE_KOSHIMAGE - 1).Value
            .Cells(row, COLUMN_POSTURE_SCORE_HIZAMAGE).Value = .Cells(row, COLUMN_POSTURE_SCORE_HIZAMAGE - 1).Value

            '信頼性の強制を解除
            .Cells(row, COLUMN_FORCED_SECTION_KOBUSHIAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_KOSHIMAGE).Value = 0
            .Cells(row, COLUMN_FORCED_SECTION_HIZAMAGE).Value = 0

        '強制
        Else
            '信頼性を強制にする列を決める
            If column_forced_num = COLUMN_POSTURE_SCORE_KOBUSHIAGE Then
                column_reliability_forced_num = COLUMN_FORCED_SECTION_KOBUSHIAGE
            ElseIf column_forced_num = COLUMN_POSTURE_SCORE_KOSHIMAGE Then
                column_reliability_forced_num = COLUMN_FORCED_SECTION_KOSHIMAGE
            ElseIf column_forced_num = COLUMN_POSTURE_SCORE_HIZAMAGE Then
                column_reliability_forced_num = COLUMN_FORCED_SECTION_HIZAMAGE
            End If

            '除外を解除
            .Cells(row, COLUMN_REMOVE_SECTION).Value = 0
            '拳上腰曲げ膝曲げのいずれかを強制
            .Cells(row, column_forced_num).Value = vle
            '全体の信頼性強制
            .Cells(row, COLUMN_FORCED_SECTION_TOTAL).Value = bit
            '拳上腰曲げ膝曲げのいずれか信頼性を強制にする
            .Cells(row, column_reliability_forced_num).Value = 1
        End If
    End With

End Sub


'------------------------------------------------------------
' 姿勢スコアの基準値（Base）と全体スコアとの同期処理
'
' 引数:
'   row - 対象となる行番号（ポイント計算シート）
'   bit - 処理モードフラグ
'         1：現在の全体スコアを基準値として記録（上書きはしない）
'         0：基準値を全体スコアに復元（リセット）
'------------------------------------------------------------
Sub baseScore(row As Long, bit As Long)
    With ThisWorkbook.Sheets("ポイント計算シート")

        If bit = 1 Then ' 保存（基準スコアを未設定なら記録）
            If .Cells(row, COLUMN_BASE_SCORE).Value = "" Then
                .Cells(row, COLUMN_BASE_SCORE).Value = .Cells(row, COLUMN_POSTURE_SCORE_ALL).Value
            End If

        Else ' 復元（基準スコア → 姿勢スコアに戻す）
            .Cells(row, COLUMN_POSTURE_SCORE_ALL).Value = .Cells(row, COLUMN_BASE_SCORE).Value
        End If

    End With
End Sub


'『初期化』ボタンが押されたとき
Sub reset()
    Call forceResult(-1)
End Sub


'姿勢点『0』強制ボタンが押されたとき
Sub force0()
    Call forceResult(0)
End Sub


'姿勢点『1』強制ボタンが押されたとき
Sub force1()
    Call forceResult(1)
End Sub


'姿勢点『除外』強制ボタンが押されたとき
Sub jogai()
    Call forceResult(99)
End Sub


'データ区間の割合を計算
Sub checkReliabilityRatio()
    '変数定義
    Dim i                               As Long
    'フレームレート
    Dim fps                             As Double
    'ポイント計算シート最終行
    Dim maxRowNum                       As Long
    '姿勢評価修正シートの最終列
    Dim ColumnNum                       As Long
    '配列の最終値
    Dim maxArrayNum                     As Long
    '信頼性の番号
    '1:測定、2:推定、3:欠損、4:強制、5:除外
    Dim reliabilityFlag                 As Long
    '信頼性の割合
    Dim measurementSectionRatio         As Double
    Dim predictSectionRatio             As Double
    Dim missingSectionRatio             As Double
    Dim coercionSectionRatio            As Double
    Dim exclusionSectionRatio           As Double
    Dim totalRatio                      As Double

    '配列定義
    '色を保存する配列
    Dim reliabilityColorDataArray()     As Long
    '色をカウントする配列
    '信頼性1 ~ 3のフレーム数をそれぞれ合計
    '1:測定、2:推定、3:欠損、4:強制、5:除外
    Dim reliabilityColorCounterArray(5) As Long

    '---------------------------------------------
    'RGBを指定するための変数を定義
    '---------------------------------------------
    '信頼性
    Dim colorMeasureSection    As Long: colorMeasureSection = RGB(0, 176, 240)  '水色
    Dim colorPredictSection    As Long: colorPredictSection = RGB(252, 246, 0)  '黄色
    Dim colorMissingSection    As Long: colorMissingSection = RGB(255, 124, 128)'ピンク
    Dim colorForcedSection     As Long: colorForcedSection  = RGB(0, 51, 204)   '青色
    Dim colorRemoveSection     As Long: colorRemoveSection  = RGB(191, 191, 191)'グレー

    '---------------------------------------------
    '変数・配列準備
    '---------------------------------------------
    'フレームレートを取得
    fps = getFps()
    '最終行を取得
    maxRowNum = getLastRow()

    '姿勢評価修正シート
    Dim sName()  As String
    Dim n        As Long
    Dim actSheet As Worksheet

    '余分を消す
    maxRowNum = maxRowNum - 1

    '一番右の列数を取得
    With ThisWorkbook.Sheets("姿勢評価修正シート")
        ColumnNum = Columns.Count - 6
    End With

    '最初の列数(6列まで)分を追加する
    ColumnNum = 16206
    maxArrayNum = ColumnNum - 1

    '配列を再定義
    ReDim reliabilityColorDataArray(maxArrayNum, 0)

    'カウンターを初期化
    Erase reliabilityColorCounterArray

    '---------------------------------------------
    'ここから信頼性の割合を計算
    '---------------------------------------------

    For i = 2 To maxRowNum + 1 '230208

        With ThisWorkbook.Sheets("ポイント計算シート")

            '除外
            If .Cells(i, COLUMN_REMOVE_SECTION).Value > 0 Then
                reliabilityColorCounterArray(5) = reliabilityColorCounterArray(5) + 1
                GoTo CONTINUE:
            '強制
            ElseIf .Cells(i, COLUMN_FORCED_SECTION_TOTAL).Value > 0 Then
                reliabilityColorCounterArray(4) = reliabilityColorCounterArray(4) + 1
                GoTo CONTINUE:
            '欠損
            ElseIf .Cells(i, COLUMN_MISSING_SECTION).Value > 0 Then
                reliabilityColorCounterArray(3) = reliabilityColorCounterArray(3) + 1
                GoTo CONTINUE:
            '推定
            ElseIf .Cells(i, COLUMN_PREDICT_SECTION).Value > 0 Then
                reliabilityColorCounterArray(2) = reliabilityColorCounterArray(2) + 1
                GoTo CONTINUE:
            '測定
            ElseIf .Cells(i, COLUMN_MEASURE_SECTION).Value > 0 Then
                reliabilityColorCounterArray(1) = reliabilityColorCounterArray(1) + 1
                GoTo CONTINUE:

            End If
        End With

CONTINUE:
    Next

    '割合を計算
    '推定
    predictSectionRatio = reliabilityColorCounterArray(2) / maxRowNum * 100
    '欠損
    missingSectionRatio = reliabilityColorCounterArray(3) / maxRowNum * 100
    '除外
    exclusionSectionRatio = reliabilityColorCounterArray(5) / maxRowNum * 100
    '測定
    measurementSectionRatio = reliabilityColorCounterArray(1) / maxRowNum * 100
    '強制
    coercionSectionRatio = reliabilityColorCounterArray(4) / maxRowNum * 100

    Set actSheet = ActiveSheet
    sName() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢評価修正シート")
    For n = 1 To UBound(sName)
        '割合をセルに入力
        With ThisWorkbook.Sheets(sName(n))
            '測定
            .Cells(3, 4) = Round(measurementSectionRatio, 1) & "%"
            '強制
            .Cells(4, 4) = Round(coercionSectionRatio, 1) & "%"
            '除外
            .Cells(5, 4) = Round(exclusionSectionRatio, 1) & "%"
            '推定
            .Cells(6, 4) = Round(predictSectionRatio, 1) & "%"
            '欠損
            .Cells(7, 4) = Round(missingSectionRatio, 1) & "%"
            '測定+強制+除外
            .Cells(3, 5) = Round(measurementSectionRatio + coercionSectionRatio + exclusionSectionRatio, 1) & "%"
            '推定+欠損
            .Cells(6, 5) = Round(predictSectionRatio + missingSectionRatio, 1) & "%"

        End With
    Next
End Sub


'------------------------------------------------------------
' 幅調整処理（拡大・縮小）
'
' 概要:
'   ボタン操作に応じて姿勢素点修正シートの列幅を拡大／縮小する。
'
' 引数:
'   expansionFlag - 拡大/縮小のフラグ
'     True  = 拡大
'     False = 縮小
'
' 備考:
'   初回呼び出し時に現在の幅サイズを基準として記録する。
'------------------------------------------------------------
Sub adjustWidth(expansionFlag As Boolean)
    Static initFin          As Boolean      ' 初回呼び出しフラグ
    Static wSize            As widthSize    ' 現在の幅サイズ

    Call stopUpdate

    '拡大・縮小どちらのフラグか確認（ボタンから引数受け取る）
    '縮小ボタン

    '初めて呼ばれた時だけ処理
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

    ' シート一覧取得と処理実行
    Dim sName() As String
    Dim n As Long
    Dim actSheet    As Worksheet: Set actSheet = ActiveSheet

    Set actSheet = ActiveSheet
    sName() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢評価修正シート")

    For n = 1 To UBound(sName)
        Call DataAjsSht.SetCellsHW(CInt(wSize), ThisWorkbook.Sheets(sName(n)))
    Next

    ' 元のシートへ戻す
    actSheet.Activate

    Call restartUpdate
End Sub


'『幅拡大』ボタンが押されたとき
Sub expandWidth()
    '引数：expansionFlag As Long　幅の拡大or縮小を決める
    'False：縮小　True:拡大
    Call adjustWidth(True)
End Sub


'『幅縮小』ボタンが押されたとき
Sub reduceWidth()
    '引数：expansionFlag As Boolean　幅の拡大or縮小を決める
    'False：縮小　True:拡大
    Call adjustWidth(False)
End Sub


'1画面左へスクロール
Sub scrollToLeftPage()
        ActiveWindow.LargeScroll ToLeft:=1
End Sub


'1画面右へスクロール
Sub scrollToRightPage()
        If ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count <= _
        ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column Then
            ActiveWindow.LargeScroll ToRight:=1
        End If
End Sub


'最も左へスクロール
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


'------------------------------------------------------------
' 右端スクロール処理
'
' 概要:
'   姿勢素点修正シートの右端までスクロールする。
'   すでに右端なら次ページへ移動（nextPageボタン相当）する。
'
' 備考:
'   TIME_ROW：時刻行（列の終端判定に使用）
'   getClm, getPageShapeState, finCellPlace は補助関数
'------------------------------------------------------------
Sub scrollToRightEnd()
    Dim keepColumn As Long

    ' 右端に達しているかをチェック
    If getClm(ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column) Then
        ' ページ移動が可能なら次のページへ
        If getPageShapeState(ActiveSheet, "nextPage") Then
            Call nextPage_Click
        End If
    Else
        ' 現在の最右列を取得（※必要ならkeepColumnに保持可能）
        keepColumn = ActiveSheet.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column

        ' スクロール位置を設定（見やすい位置まで29列戻す）
        ActiveWindow.Panes(2).ScrollColumn = keepColumn - 29

        ' 小スクロールでスクロール範囲を表示に収める
        ActiveWindow.SmallScroll ToLeft:=ActiveWindow.Panes(2).VisibleRange.Columns.Count

        ' カーソル・セル位置の最終調整
        Call finCellPlace(ActiveSheet)
    End If
End Sub

'------------------------------------------------------------
' 同じ列に対して連続でスクロール処理が呼ばれたかを判定する
'
' 引数:
'   clm - 現在のカラム番号（列位置）
'
' 戻り値:
'   True  - 直前のカラムと同じ（連続呼び出し）
'   False - カラムが変わった（初回または位置変更）
'
' 備考:
'   Static変数 keepColumn によって、前回呼び出し時の列位置を記憶する。
'------------------------------------------------------------
Private Function getClm(clm As Long) As Boolean
    Static keepColumn   As Long
    Dim isSameColumn    As Boolean

    If keepColumn = clm Then
        isSameColumn = True
    Else
        keepColumn = clm
        isSameColumn = False
    End If

    getClm = isSameColumn
End Function


'表示倍率を画面にフィット
Sub fit()
    '見えている列範囲を取得
    Dim visibleColumn As String

    '見えている列範囲のうち左から7番目の列を取得（編集ボタンが置かれている1～6列を飛ばす）
    visibleColumn = Split(ActiveWindow.VisibleRange.Cells(7, 1).Address(True, False), "$")(0)
    '1～時刻の１行下までを選択
    Range(visibleColumn & "1:" & visibleColumn & BOTTOM_OF_TABLE + 1).Select
    '画面にフィット
    ActiveWindow.Zoom = True
    'A1セルを選択する
    Range("A1").Select
    '画面を一番上までスクロール
    ActiveWindow.ScrollRow = 1

End Sub


'------------------------------------------------------------
' 再生ボタン処理
'
' 概要:
'   姿勢素点修正シート上で再生ボタンが押された際に、
'   時間カラムの自動選択処理を定期的に実行する。
'
' 引数:
'   なし（グローバル変数 ResTime を使用して再帰的に呼び出される）
'
' 備考:
'   - シートが「姿勢素点修正シート」でなければ再帰処理を停止する。
'   - 再生ボタンの非表示処理を実行する。
'------------------------------------------------------------
Sub RegularInterval1()
    Dim iend            As Long
    Dim i               As Long
    Dim dajsht()        As String
    Dim currentColumn   As Long
    Dim ws              As Worksheet
    Set ws = ActiveSheet

    ' 対象となる修正シート一覧を取得
    dajsht = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    iend = UBound(dajsht)

    ' 全シートで再生ボタンを非表示にする
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes("playBtn").Visible = False
        End With
    Next

    ' カラム位置確認
    currentColumn = ActiveCell.Column
    If currentColumn < TIME_COLUMN_LEFT Then
        ActiveSheet.Cells(BOTTOM_OF_TABLE, TIME_COLUMN_LEFT).Select
        ' 初期スタートに見せるため1秒待機
        Application.Wait Now() + TimeValue("00:00:01")
    End If

    ' 次回実行時刻（1秒後）を設定
    ResTime = Now + TimeValue("00:00:01")

    ' 自身を1秒後に再実行する設定
    Application.OnTime EarliestTime:=ResTime, Procedure:="RegularInterval3"

    ' シート名確認（姿勢素点修正シートのみ継続）
    If ActiveSheet.Name Like "姿勢素点修正シート*" Then
        Call nextTimeSelect
    Else
        Call Cancel1
    End If
End Sub


'------------------------------------------------------------
' 時刻選択処理
'
' 概要:
'   アクティブセルの次の時間セルへ移動し、
'   時刻が表示されていなければ次シートへ遷移する。
'
' 引数:
'   なし
'
' 備考:
'   - アクティブセルの行はそのままで、右方向へ時刻を進める。
'   - 時刻が存在しない場合は終了または次ページへ遷移。
'------------------------------------------------------------
'時刻選択処理
Sub nextTimeSelect()

    'アクティブセルの一番左の列数を取得
    '取得した列数の時刻(23行目）をアクティブにする
    Cells(TIME_ROW, Selection.Column).Select

    '一つ右のセル（次の時刻）に移動
    ActiveCell.Offset(0, 1).Select

    '表示を1秒分右へスクロール
    ActiveWindow.SmallScroll ToRight:=TIME_WIDTH

    ' 選択されたセルが空白なら処理継続 or 終了判定
    If IsEmpty(ActiveCell.Value) Then
        'arrowが見えている時、すなわち次のシートが存在する。
        If getPageShapeState(ActiveSheet, "nextPage") Then
            '次のページが存在すれば遷移
            Call nextPage_Click
        Else
            ' 時刻が存在せず、次ページもなければ処理終了
            Call Cancel3
        End If
    End If

End Sub


'------------------------------------------------------------
' 停止ボタン処理
'
' 概要:
'   姿勢素点修正シート上で「停止ボタン」が押されたときに、
'   再生の自動実行を止め、各シートの再生ボタンを再表示する。
'
' 備考:
'   - 再生制御用の ResTime を用いて Application.OnTime を解除。
'   - 姿勢素点修正シートすべての再生ボタンを再表示。
'------------------------------------------------------------
Sub Cancel3()
    Dim iend As Long, i As Long
    Dim dajsht() As String

    ' 姿勢素点修正シート一覧を取得
    dajsht = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢素点修正シート")
    iend = UBound(dajsht)

    ' 各シートの再生ボタンを再表示
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes("playBtn").Visible = True
        End With
    Next

    ' Application.OnTime を使ってタイマーを停止
    On Error Resume Next
    Application.OnTime EarliestTime:=ResTime, _
                       Procedure:="RegularInterval1", _
                       Schedule:=False
    On Error GoTo 0
End Sub

'メッセージボックスの表示
'戻り値：メッセージボックス
Function YesorNo() As VbMsgBoxResult
    YesorNo = MsgBox("この場所に" & ActiveWorkbook.Name & _
                        "という名前のファイルが既にあります。置き換えますか？", _
                        vbInformation + vbYesNoCancel + vbDefaultButton2)
End Function


'ブック全体の保存
Sub Savebook()
    Dim dotPoint     As String
    Dim workbookName As String
    Dim fps          As Double

    'フレームレートを取得
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


'sheetの左から何番に属するか判定する
'引数1：シート
'戻り値：シートが何番目に属しているか
Function calcSheetNamePlace(ws As Worksheet)
    Dim shNameArray()   As String
    Dim i               As Long
    Dim iend            As Long
    Dim ret             As Long: ret = 0

    shNameArray() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢評価修正シート")
    iend = UBound(shNameArray)
    For i = 1 To iend
        If ws.Name = shNameArray(i) Then
            ret = i - 1
        End If
    Next
    calcSheetNamePlace = ret
End Function


'------------------------------------------------------------
' 対象名を含むワークシート名を配列として返す関数
'
' 概要:
'   指定された名前を含むワークシートを左から順に検索し、
'   一致したシート名を配列として返す。
'
' 引数:
'   wb  : Workbook オブジェクト
'   str : シート名に含まれる文字列（例: "姿勢素点修正シート"）
'
' 戻り値:
'   一致したワークシート名の配列（String型）
'------------------------------------------------------------
Function call_GetSheetNameToArrayspecific(wb As Workbook, str As String) As String()
    Dim tmp()   As String
    Dim ws      As Worksheet
    Dim i       As Long: i = 0

    ' 全ワークシートを走査して対象名を含むものを追加
    For Each ws In wb.Worksheets
        If InStr(ws.Name, str) > 0 Then
            i = i + 1
            ReDim Preserve tmp(1 To i)
            tmp(i) = ws.Name
        End If
    Next

    call_GetSheetNameToArrayspecific = tmp
End Function


'簡易的なシート切替処理も兼ねた矢印の図形
'引数1：姿勢評価修正シート
'引数2：前ページに移動するアイコンを非表示にするかどうか（true or false）
'引数3：次ページに移動するアイコンを非表示にするかどうか（true or false）
Private Sub addPageShape(ws As Worksheet, pPageState As Boolean, nPageState As Boolean)
    Const pPage As String = "prevPage"
    Const nPage As String = "nextPage"

    Call initCellPlace(ws)

    ws.Shapes(pPage).Visible = pPageState
    ws.Shapes(nPage).Visible = nPageState
End Sub


'図形がVisibleかどうか判定する
'引数1：ワークシート
'引数2：図形の名前
'戻り値:Visibleかどうか（0 or 1）
Private Function getPageShapeState(ws As Worksheet, shapeName As String)
    getPageShapeState = ws.Shapes(shapeName).Visible
End Function


'ワークシートをコピーし、右に挿入
Sub createSheet(endclm As Long)
    ThisWorkbook.ActiveSheet.Copy After:=ActiveSheet
End Sub


'ひとつ前のシートをアクティブにし、データの最後尾まで行く
Sub prevPage_Click()
    ThisWorkbook.ActiveSheet.Previous.Activate
    Call finCellPlace(ThisWorkbook.ActiveSheet)
End Sub


'ひとつ次のシートをアクティブにし、データの最初に行く
Sub nextPage_Click()
    ThisWorkbook.ActiveSheet.Next.Activate
    Call initCellPlace(ThisWorkbook.ActiveSheet)
End Sub


'セルの初期位置
Private Sub initCellPlace(ws As Worksheet)
    ws.Cells(TIME_ROW, TIME_COLUMN_LEFT).Select
End Sub


'セルの最終位置
Private Sub finCellPlace(ws As Worksheet)
    ws.Cells(TIME_ROW, ws.Cells(TIME_ROW, Columns.Count).End(xlToLeft).Column).Select
End Sub


'段階的にサイズの変更処理をする為の関数
'引数1：画面の拡大率
'引数2：サイズを変更できるかどうか
'戻り値：Small = 1
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
'                ベースファイルの保存が悪かった時用
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
            '前にならないとき
            If Not nextChange Then
                tmpsize = widthSize.Large
                Call changeBtnState(EXPANDBTN_NAME, True)
            Else
                tmpsize = widthSize.LL
                Call changeBtnState(REDUCEBTN_NAME, True)
'                ベースファイルの保存が悪かった時用
                Call changeBtnState(EXPANDBTN_NAME, False)
            End If
    End Select
    sizeNext = tmpsize
End Function

Sub doNothing_btn()
    'なにもしない
End Sub

'幅調整用のボタンに使う予定。実際名前さえ決めることができればなんとでもなる。

'引数1：ボタンの名前（EXPANDBTN_NAME or REDUCEBTN_NAME）
'引数2：ボタンを押せるかどうか
Private Sub changeBtnState(btnName As String, btnstate As Boolean)
    Dim iend, i As Long
    Dim dajsht() As String

    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢評価修正シート")
    iend = UBound(dajsht)
    For i = 1 To iend
        With Worksheets(dajsht(i))
            .Shapes(btnName).Visible = btnstate
        End With
    Next
End Sub

'シートをリセットする
Sub resetSheet()
    Const pPage As String = "prevPage"
    Const nPage As String = "nextPage"
    Dim iend, i As Long
    Dim dajsht() As String
    dajsht() = call_GetSheetNameToArrayspecific(ThisWorkbook, "姿勢評価修正シート")
    iend = UBound(dajsht)
    For i = 1 To iend
        With Worksheets(dajsht(i))
            '全て隠す
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


'非表示の名前の定義を再表示　20230215　早川　シートコピー時に発生するエラー対策
Public Sub ShowInvisibleNames()
    Dim oName As Object
    For Each oName In Names
        If oName.Visible = False Then
            oName.Visible = True
        End If
    Next
    MsgBox "非表示の名前の定義を表示しました。", vbOKOnly
End Sub


Private Sub UserForm_Terminate()
    MsgBox "Excelの画面を表示します"
    Application.Visible = True
End Sub


' 選択範囲をデータ有効域とチェックし有効内の値を返す
' 戻り値 : True → 交差あり（ leftCol/rightCol が返る ）
'          False → 交差なし（メッセージは呼び出し側で）
Public Function CropSelectionToDataArea(ByRef leftCol As Long, ByRef rightCol As Long) As Boolean
    Const PAGE_FRAME_MAX    As Long = LIMIT_COLUMN '16200
    Dim shtPage             As Long
    Dim baseClm             As Long
    Dim selR                As Long '選択列
    Dim frmR                As Long '選択フレーム
    Dim pageFrmR            As Long 'ページの有効フレーム
    Dim totalFrm            As Long

    'ボタン列を一緒に選んだら無視
    If Selection.Column > Columns.Count Then
        Exit Function
    End If

    selR = Selection.Column + Selection.Columns.Count - 1 '選択列の長さ

    shtPage = calcSheetNamePlace(ActiveSheet)
    baseClm = LIMIT_COLUMN * shtPage

    With Worksheets("ポイント計算シート")
        totalFrm = .Cells(1, 3).End(xlDown).row - 1
    End With

    '列 → フレームへ
    frmR = selR - COLUMN_ZERO_NUM + baseClm
    pageFrmR = WorksheetFunction.min(baseClm + PAGE_FRAME_MAX, totalFrm)
    frmR = WorksheetFunction.min(frmR, pageFrmR)    '右辺においてページ内の有効フレーム数を超えないようにする

    '姿勢素点修正シートで始まりの列かどうかをチェックし、最低値以下が選択されていた場合はCOLUMN_ZERO+1
    leftCol = WorksheetFunction.Max(Selection.Column, COLUMN_ZERO_NUM + 1)
    rightCol = frmR - baseClm + COLUMN_ZERO_NUM

    If leftCol > rightCol Then
        CropSelectionToDataArea = False   '重なりなし
    Else
        'フレーム → 列番号へ戻す
        CropSelectionToDataArea = True
    End If
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