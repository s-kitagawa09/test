Attribute VB_Name = "Module1"
Sub 処理01_非請求削除()

Dim Fso As FileSystemObject


    Call 非請求削除(dummy)
    
    
    Awb = ActiveWorkbook.Name
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Ost = Fso.GetBaseName(Awb)

    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & Left(Ost, Len(Ost) - 2) & "01_非請求削除.xlsx"

End Sub

Sub 処理02_区分入力()

Dim Fso As FileSystemObject


    Call 区分入力1(dummy)
    Call 区分入力2(dummy)
    
    
    Awb = ActiveWorkbook.Name
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Ost = Fso.GetBaseName(Awb)

    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & Left(Ost, Len(Ost) - 8) & "02_区分入力.xlsx"

End Sub

Sub 処理03_備考1調整()

Dim Fso As FileSystemObject


    Call コーヒーカタログ置換(dummy)
    
    
    Awb = ActiveWorkbook.Name
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Ost = Fso.GetBaseName(Awb)

    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & Left(Ost, Len(Ost) - 7) & "03_備考1調整.xlsx"

End Sub

Sub 処理04_別送追加()

Dim Fso As FileSystemObject


    Call 別送追加(dummy)
    
    
    Awb = ActiveWorkbook.Name
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Ost = Fso.GetBaseName(Awb)

    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & Left(Ost, Len(Ost) - 8) & "04_別送追加.xlsx"

End Sub

Sub 処理05_ソートとエリア分け()

Dim Fso As FileSystemObject


    Call 日付順ソート(dummy)
    Call エリア毎単価設定(dummy)
    Call 単価色付け(dummy)
    Call 青箱一式に文言追加(dummy)
    Call 梱包作業費変更(dummy)
    
    
    Awb = ActiveWorkbook.Name
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Ost = Fso.GetBaseName(Awb)

    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & Left(Ost, Len(Ost) - 7) & "05_ソートとエリア分け.xlsx"

End Sub

Sub 処理06_確認用()

Dim Fso As FileSystemObject


    Call シート複製と不要行削除(dummy)
    
    
    Awb = ActiveWorkbook.Name
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Ost = Fso.GetBaseName(Awb)

    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & Left(Ost, Len(Ost) - 12) & "06_確認用.xlsx"

End Sub
