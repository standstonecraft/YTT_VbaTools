Attribute VB_Name = "YTT_MergeCellHgs"
'''
''' 方眼紙上で表を作る際などに使用。
''' 選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す
'''
Sub YTT_方眼紙セル結合_中央寄せ()
Attribute YTT_方眼紙セル結合_中央寄せ.VB_Description = "方眼紙上で表を作る際などに使用。\r\n選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す"
    YTT_方眼紙セル結合 True, False
End Sub
'''
''' 方眼紙上で表を作る際などに使用
''' 選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す
'''
Sub YTT_方眼紙セル結合_罫線()
Attribute YTT_方眼紙セル結合_罫線.VB_Description = "方眼紙上で表を作る際などに使用。\r\n選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す"
    YTT_方眼紙セル結合 False, True
End Sub
'''
''' 方眼紙上で表を作る際などに使用
''' 選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す
'''
Sub YTT_方眼紙セル結合_中央寄せ_罫線()
Attribute YTT_方眼紙セル結合_中央寄せ_罫線.VB_Description = "方眼紙上で表を作る際などに使用。\r\n選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す"
    YTT_方眼紙セル結合 True, True
End Sub
'''
''' 方眼紙上で表を作る際などに使用
''' 選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す
'''
Sub YTT_方眼紙セル結合(centering As Boolean, surrounding As Boolean)
Attribute YTT_方眼紙セル結合.VB_Description = "方眼紙上で表を作る際などに使用。\r\n選択範囲内で文字が入っている列から次に文字が入っている列の手前までを結合することを繰り返す"
    
    If Selection.Count > 300 Then
        MsgBox "範囲が大きすぎます（>300）", vbCritical
        Exit Sub
    End If
    If Selection.Rows.Count > 4 Then
        MsgBox "行数が大きすぎます（>4）", vbCritical
        Exit Sub
    End If

    '現在のシート
    Dim sh As Worksheet
    Set sh = ActiveSheet
    'ターゲット行
    Dim thisRow As Integer
    thisRow = Selection(1).Row
    '開始列、ポインタ列
    Dim pointerCol As Integer
    pointerCol = Selection(1).Column
    '終了列
    Dim endCol As Integer
    endCol = Selection.Column + Selection.Columns.Count
    '行数
    Dim rowSize As Integer
    rowSize = Selection.Rows.Count
    
    
    'まず結合解除
    Selection.MergeCells = False
    
    'ポインターが最終列を越えるまで繰り返す
    While pointerCol < endCol
        'ポインタ列の値の個数
        Dim baseValCnt As Integer
        baseValCnt = valCount(sh.Cells(thisRow, pointerCol).Resize(rowSize))
        
        '結合する列数を調べる
        Dim colSize As Integer
        For colSize = 2 To 100
            '最終列を越えた
            If pointerCol + colSize - 1 >= endCol Then
                Exit For
            End If
            '新しい範囲の値の個数
            Dim rngValCnt As Integer
            rngValCnt = valCount(sh.Cells(thisRow, pointerCol).Resize(rowSize, colSize))
            '次の列に進入した
            If baseValCnt < rngValCnt Then
                Exit For
            End If
        Next colSize
        '次の列または最終の次の列になっているので一つ減らす
        colSize = colSize - 1
        
        '結合するセル
        Dim rngToMerge As range
        Set rngToMerge = sh.Cells(thisRow, pointerCol).Resize(rowSize, colSize)
        
        '結合する
        rngToMerge.MergeCells = True
        
        '中央寄せ
        If centering Then
            rngToMerge.VerticalAlignment = xlCenter
            rngToMerge.HorizontalAlignment = xlCenter
        End If
        
        '罫線で囲む
        If surrounding Then
            rngToMerge.Borders.LineStyle = xlContinuous
        End If
        
        'ポインタを進める
        pointerCol = pointerCol + colSize
    Wend
End Sub

'''
''' 範囲内の値の個数を返す
'''
Private Function valCount(rng As range) As Integer
    valCount = WorksheetFunction.CountIf(rng, "<>")
End Function

