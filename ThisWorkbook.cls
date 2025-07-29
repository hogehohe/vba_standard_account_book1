VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Option Explicit '変数の宣言を強制
'
''======================================================================================
''字幕情報の定義
''======================================================================================
'Const CAPTION_TRACK2_FILE_NAME_SOEJI           As String = "2"   '字幕トラック２用のファイル名末尾につける添字
'Const CAPTION_CHUKAN_FILE_NAME_SOEJI           As String = "tmp" '中間ファイルにつける添字
'
'
'Private Sub Workbook_Open()
'    Dim PosExt         As Long
'    Dim FileName       As String
'    Dim CaptionRow1    As String
'    Dim TrackNum       As Long
'
'    FileName = ThisWorkbook.Name 'ファイル名を取得
'    PosExt = InStrRev(FileName, ".") '.以前の文字数を取得
'
'    '--- 拡張子を除いたパス（ファイル名）を格納する変数 ---'
'    If (0 < PosExt) Then
'        FileName = Left(FileName, PosExt - 1)
'    End If
'
'    '1回でもMacroUpdateDataが動いたことがあれば（初回分析時には起動させないため）
'    If ThisWorkbook.Sheets("ポイント計算シート").Cells(2, 196).Value = 1 Then
'    'If ThisWorkbook.Sheets("条件設定シート").Cells(1, 80).Value = 1 Then
'
'        If Dir(ActiveWorkbook.Path & "\" & FileName & ".srt") <> "" Then
'            'エクセルと同じファイル名の字幕ファイルを開く
'            Open ActiveWorkbook.Path & "\" & FileName & ".srt" For Input As #1
'
'            '字幕ファイルの１行目を取得して、字幕ファイルを閉じる
'            Line Input #1, CaptionRow1
'            Close #1
'
'            '字幕ファイル１行目の文字列を確認する
'            '半角スペース1個(字幕トラック1）
'            If CaptionRow1 = " 1" Then
'                TrackNum = 1
'            '半角スペース2個(字幕トラック2）
'            ElseIf CaptionRow1 = "  1" Then
'                TrackNum = 2
'            Else
'                TrackNum = -1
'            End If
'        Else
'            MsgBox "字幕ファイル(.srt)が存在しません。管理者にお問い合わせください。"
'        End If
'    End If
'
'    '字幕トラックとチェックボックスの組み合わせが合っていなければ、字幕ファイルの名前を入れ替える
'    If ThisWorkbook.Worksheets("姿勢重量点調査票").CheckBox1 = True _
'        And TrackNum = 1 _
'        Or ThisWorkbook.Worksheets("姿勢重量点調査票").CheckBox1 = False _
'        And TrackNum = 2 Then
'        '---字幕ファイルをリネームする---
'        'Name 変更前のファイル名 As 変更後のファイル名
'        'トラック2用字幕を中間ファイルに変更
'        Name ActiveWorkbook.Path & "\" & FileName & CAPTION_TRACK2_FILE_NAME_SOEJI & ".srt" _
'          As ActiveWorkbook.Path & "\" & FileName & CAPTION_CHUKAN_FILE_NAME_SOEJI & ".srt"
'
'        'トラック１用字幕をトラック２用字幕に変更
'        Name ActiveWorkbook.Path & "\" & FileName & ".srt" _
'          As ActiveWorkbook.Path & "\" & FileName & CAPTION_TRACK2_FILE_NAME_SOEJI & ".srt"
'
'        '中間ファイルをトラック１用字幕に変更
'        Name ActiveWorkbook.Path & "\" & FileName & CAPTION_CHUKAN_FILE_NAME_SOEJI & ".srt" _
'          As ActiveWorkbook.Path & "\" & FileName & ".srt"
'    End If
'End Sub

Private Sub Workbook_Open()

    '実行時エラー防止の為にブックを開いたときにシート保護をかける
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "姿勢素点修正シート") > 0 Then
            On Error Resume Next
            ws.Unprotect
            ws.Protect UserInterfaceOnly:=True
            On Error GoTo 0
        End If
    Next ws

End Sub
