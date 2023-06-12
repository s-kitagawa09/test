Attribute VB_Name = "Module1"

Sub 区分入力1(dummy)

Dim LastRow As Long
Dim i As Long
    
    '最終行取得
    'Gitテスト
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '最終行までループ
    For i = 17 To LastRow
    
        '店番号を使ってO列入力
        On Error Resume Next
        
        With Cells(i, "D")
            .Offset(0, 11) = WorksheetFunction.VLookup(.Value, Worksheets("Sheet1").Range("A:C"), 3, False)
    
        End With
        
        '営業所を使ってP列入力
        With Cells(i, "E")
            .Offset(0, 11) = WorksheetFunction.VLookup(.Value, Worksheets("Sheet1").Range("B:C"), 3, False)
    
        End With
        
        On Error GoTo 0
        
    Next

End Sub

Sub 区分入力2(dummy)

Dim LastRow As Long
Dim i As Long

    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        '最終行までループ
        For i = 17 To LastRow

            '特定の送付先の時の区分入力
            If InStr(Cells(i, "E"), "小林工芸社") >= 1 Then
                .Cells(i, "B") = "移管分"
                
            ElseIf InStr(Cells(i, "E"), "DNPロジスティクス") >= 1 Then
                .Cells(i, "B") = "転送分"
                
            End If
            
        Next
        
    End With
    
End Sub

Sub 非請求削除(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim Target As Variant

    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '対象文字
    Target = Array("非請求")

    Application.DisplayAlerts = False
        With ActiveSheet

            '最終行から先頭へループ
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target)

                    '特定文字があったら行削除
                    If InStr(Cells(i, 13), Target(j)) >= 1 Then
                        .Rows(i).Delete
                    End If

                Next
                
            Next
            
        End With
    Application.DisplayAlerts = True

End Sub

Sub コーヒーカタログ置換(dummy)

Dim LastRow As Long
Dim i As Long

    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        '最終行から先頭へループ
        For i = LastRow To 1 Step -1

            On Error Resume Next
            
            'コーヒーとカタログ置換
            If InStr(Cells(i, "F"), "コーヒー") >= 1 Then
                .Cells(i, "F").ClearContents
                .Cells(i, "F") = "FFコーヒー関連"
                
            ElseIf InStr(Cells(i, "F"), "FAMIMACAFE") >= 1 Then
                .Cells(i, "F").ClearContents
                .Cells(i, "F") = "FFコーヒー関連"
                
            ElseIf InStr(Cells(i, "F"), "カタログ") >= 1 Then
                .Cells(i, "F").ClearContents
                .Cells(i, "F") = "カタログ関連"
                
            End If
            
            On Error GoTo 0
            
        Next
        
    End With
    
End Sub

Sub 別送追加(dummy)

Dim LastRow As Long
Dim myDate As Date '発行日
Dim fstDate As Date '月初
Dim i As Long
Dim bs As Borders
Dim myMon As Date

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    myDate = Range("L5")
    fstDate = Year(myDate) & "/" & Month(myDate) & "/" & 1
    
    i = GetWeekDay(fstDate, Range("L5"), vbTuesday)

    '行の追加
    Rows(LastRow + 1).Insert
    
    '背景色を黄色に変更
    Range(Cells(LastRow + 1, "A"), Cells(LastRow + 1, "L")).Interior.ColorIndex = 6
    
    '枠線の追加
    Set bs = Range(Cells(LastRow + 1, "A"), Cells(LastRow + 1, "L")).Borders
    bs.LineStyle = xlContinuous
    
    '文字の追加
    Cells(LastRow + 1, "E") = "山形南営業所"
    Cells(LastRow + 1, "F") = "FXSSショーカード別送"
    Cells(LastRow + 1, "G") = 1
    Cells(LastRow + 1, "H") = 185
    Cells(LastRow + 1, "I") = 185
    Cells(LastRow + 1, "J") = 600
    Cells(LastRow + 1, "K") = 600
    Cells(LastRow + 1, "L") = 785
    
    '行の複製
    Rows(LastRow + 1).Copy
    Range(Rows(LastRow + i - 1), Rows(LastRow + 1)).Insert
    
End Sub

Sub 日付順ソート(dummy)

Dim LastRow As Long
Dim sht As Worksheet

    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        'ソートのクリア
        .Sort.SortFields.Clear
    
        'ソートの設定
        .Sort.SortFields.Add _
            Key:=ActiveSheet.Cells(16, "A"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        
        '日付列を基準にソート
        With .Sort
            .SetRange Range(Cells(16, "A"), Cells(LastRow - 1, "L"))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
        End With
        
    End With

    Set sht = ActiveSheet
    sht.Cells.Interior.ColorIndex = xlNone

End Sub

Sub エリア毎単価設定(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim Target1 As Variant
Dim Target2 As Variant
Dim Target3 As Variant
    
    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '対象文字
    Target1 = Array("中四国", "九州1", "九州2", "南九州")
    Target2 = Array("沖縄")
    Target3 = Array("西日本")
    
    Application.DisplayAlerts = False

        With ActiveSheet

            '最終行から先頭へループ
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target1)
                For k = 0 To UBound(Target2)
                For l = 0 To UBound(Target3)

                    'Target1
                    If InStr(Cells(i, "O"), Target1(j)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If
                    
                    If InStr(Cells(i, "P"), Target1(j)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If
                    
                    If InStr(Cells(i, "Q"), Target1(j)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If
                    
                    'Target2
                    If InStr(Cells(i, "O"), Target2(k)) >= 1 Then
                        .Cells(i, "J").ClearContents
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "P"), Target2(k)) >= 1 Then
                        .Cells(i, "J").ClearContents
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "Q"), Target2(k)) >= 1 Then
                        .Cells(i, "J").ClearContents
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                
                    'Target3
                    If InStr(Cells(i, "Q"), Target3(l)) >= 1 Then
                        .Cells(i, "J") = 900
                    End If

                Next l
                Next k
                Next j
            
            Next i
        
        End With
    
    Application.DisplayAlerts = True

End Sub

Sub 単価色付け(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim Target1 As Variant
Dim Target2 As Variant
Dim Target3 As Variant
Dim Target4 As Variant

    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '対象文字
    Target1 = Array("ヤマト", "ヤマトタイムサービス")
    Target2 = Array("DM")
    Target3 = Array("航空便")
    Target4 = Array("セール品のみ送付", "青箱一式", "酒販促物一式")
    
    Application.DisplayAlerts = False
    
        With ActiveSheet

            '最終行から先頭へループ
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target1)
                For k = 0 To UBound(Target2)
                For l = 0 To UBound(Target3)
                For m = 0 To UBound(Target4)

                    '特定文字があったらセルの色変更
                    If InStr(Cells(i, "N"), Target1(j)) >= 1 Then
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "N"), Target2(k)) >= 1 Then
                        .Cells(i, "J").Interior.ColorIndex = xlNone
                    End If
                    
                    If InStr(Cells(i, "M"), Target3(l)) >= 1 Then
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If
                    
                    If InStr(Cells(i, "M"), Target4(m)) >= 1 Then
                        .Cells(i, "G").Interior.ColorIndex = 6
                        .Cells(i, "H").Interior.ColorIndex = 6
                        .Cells(i, "J").Interior.ColorIndex = 6
                    End If

                Next m
                Next l
                Next k
                Next j
                
            Next
            
        End With
        
    Application.DisplayAlerts = True

End Sub

Sub 青箱一式に文言追加(dummy)

Dim LastRow As Long
Dim i As Long

    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    With ActiveSheet

        '最終行から先頭へループ
        For i = LastRow To 1 Step -1

            On Error Resume Next
            
                '"青箱一式"があれば備考に文言追加
                If InStr(Cells(i, "M"), "青箱一式") >= 1 Then
                    Cells(i, "M") = Cells(i, "M").Value & "／梱包費のみ週次作業と同一単価"
                End If
            
            On Error GoTo 0
            
        Next
        
    End With
    
End Sub

Sub 梱包作業費変更(dummy)

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim Target As Variant

    '最終行取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    '対象文字
    Target = Array("青箱一式", "セール品のみ送付")

    Application.DisplayAlerts = False
        With ActiveSheet

            '最終行から先頭へループ
            For i = LastRow To 1 Step -1

                For j = 0 To UBound(Target)

                    '特定文字があったら梱包作業費を変更
                    If InStr(Cells(i, "M"), Target(j)) >= 1 Then
                        .Cells(i, "H") = 24
                    End If

                Next
                
            Next
            
        End With
    Application.DisplayAlerts = True

End Sub

Sub シート複製と不要行削除(dummy)

Dim LastRow1 As Long
Dim LastRow2 As Long
Dim i As Long
Dim j As Long
Dim targetSheet As Worksheet

    'シートの名前変更
    ActiveSheet.Name = "追加発送_DM便"
    
    'シートの複製
    Worksheets("追加発送_DM便").Copy Before:=Worksheets(1)
    
    '複製したシートの名前変更
    ActiveSheet.Name = "追加発送_宅配便"

    '最終行取得
    LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row

        '最終行から先頭へループ
        For i = LastRow1 To 1 Step -1
        
            On Error Resume Next

                '"ヤマトDM便"の行を削除
                If InStr(Cells(i, "N"), "ヤマトDM便") >= 1 Then
                    Rows(i).Delete
                End If
                
            On Error GoTo 0
                    
        Next
        
    '不要列を削除
    Range("N:Q").Delete
            
    'アクティブシートを"追加発送_DM便"に変更
    Sheets("追加発送_DM便").Select
    
    '最終行取得
    LastRow2 = Cells(Rows.Count, 1).End(xlUp).Row

        '最終行から先頭へループ
        For j = LastRow2 To 1 Step -1
        
            On Error Resume Next

                '"ヤマトDM便"以外の行を削除
                If InStr(Cells(j, "N"), "ヤマトDM便") = 0 Then
                    Rows(j).Delete
                End If
                
            On Error GoTo 0
                    
        Next
        
    '不要列を削除
    Range("N:Q").Delete
    
    '消えてしまったヘッダー分の行を挿入
    Range("1:16").Insert
        
    'ヤマトDM便以外を消した事で消えたタイトル等のコピペ
    Sheets("追加発送_宅配便").Select
    Range("1:16").Copy Sheets("追加発送_DM便").Range("A1")
    
    
    '不要ワークシート削除
    Application.DisplayAlerts = False
    
    For Each targetSheet In Worksheets
    
        If Not targetSheet.Name Like "追加発送*" Then
            targetSheet.Delete
            
        End If
    
    Next
    
    Application.DisplayAlerts = True
            
End Sub


