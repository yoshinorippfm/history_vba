VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CloseNextForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   OleObjectBlob   =   "CloseNextForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CloseNextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
   If Sum.Text = "" Or Coach.Text = "" Then
      MsgBox "袋の中の金額またはレジ担当コーチの名前が入力されていません。"
   
   Else
     Dim cnt As Long ' コピー先ブックのシートの枚数
     Dim Y As Long
     Dim M As Long
     Dim D As Long
     Dim shp As Shape
     Dim Stoday As Long
     Dim LastRow As Long
     
     
     Y = Year(Now)
     M = Month(Now)
     D = Day(Now)
     
     
     Workbooks.Open "\\192.168.3.63\share\garden\" & Y & "年売上管理\" & M & "月売上管理.xlsx"
     Workbooks.Open "\\192.168.3.63\share\garden\" & Y & "年収支表\" & M & "月収支表.xlsx"
     
     
    ' 売上管理のシートの保存
     cnt = Workbooks(M & "月売上管理.xlsx").Sheets.Count
     Workbooks("レジ.xlsm").Worksheets("日計取引表").Copy after:=Workbooks(M & "月売上管理.xlsx").Sheets(cnt)
  
     For Each shp In Workbooks(M & "月売上管理.xlsx").Sheets("日計取引表").Shapes
       shp.Delete
     Next shp
     
     With Workbooks("レジ.xlsm").Sheets("日計取引表")
         Stoday = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
     End With
     
     With Workbooks(M & "月売上管理.xlsx").Sheets("日計取引表")
        LastRow = .Cells(Rows.Count, 9).End(xlUp).Row + 1
        .Range("C1").Value = Y & "年" & M & "月" & D & "日" & "売上"
        .Cells(1, 4).Value = Coach.Text
        .Cells(LastRow, 8).Value = "合計"
        .Cells(LastRow, 9).Value = Stoday
     End With
     
     Workbooks(M & "月売上管理.xlsx").Sheets("日計取引表").Name = Y & "年" & M & "月" & D & "日" & "売上"
     
     
     
     '収支表の作成
     Dim i As Variant
     Dim S As Long
     Dim Scar As Long
     Dim Sins As Long
     Dim Sitm As Long
     Dim Soth As Long
     Dim Srev As Long
     Dim W As Single
    
     
     S = Scar = Sins = Sitm = Soth = Srev = W = 0
     
     
     cnt = Workbooks(M & "月収支表.xlsx").Sheets.Count + 1
     
     With Workbooks("レジ.xlsm").Worksheets("日計取引表")
     LastRow = .Cells(Rows.Count, 9).End(xlUp).Row
     For i = 3 To LastRow
       If .Cells(i, 4) = "100" Then
          W = .Cells(i, 9).Value
          Scar = Scar + W
      ElseIf .Cells(i, 4) = "110" Then
         W = .Cells(i, 9).Value
         Sins = Sins + W
      ElseIf .Cells(i, 4) = "120" Then
         W = .Cells(i, 9).Value
         Sitm = Sitm + W
      ElseIf .Cells(i, 4) = "140" Then
         W = .Cells(i, 9).Value
         Soth = Soth + W
      ElseIf .Cells(i, 4) = "150" Or .Cells(i, 4) = "160" Or .Cells(i, 4) = "170" Or .Cells(i, 4) = "180" Then
        W = .Cells(i, 9).Value
        Srev = Srev + W
      Else
         W = .Cells(i, 9).Value
         S = S + W
      End If
      W = 0
    Next i
    End With
    
    Workbooks(M & "月収支表.xlsx").Worksheets.Add after:=Worksheets(cnt - 1)
    
    
    With Workbooks(M & "月収支表.xlsx").Sheets(cnt)
       .Range("B1").Value = Y & "年" & M & "月" & D & "日" & "収支表"
       .Range("C1").Value = Coach.Text
              
       .Range("B2").Value = "収入"
       
       .Range("B3").Value = "フェンシング売上"
       .Range("B3").Interior.Color = RGB(200, 215, 255)
       .Range("C3").Value = S
       
       .Range("B4").Value = "駐車場"
       .Range("B4").Interior.Color = RGB(200, 215, 255)
       .Range("C4").Value = Scar
       
       .Range("B5").Value = "スポーツ保険"
       .Range("B5").Interior.Color = RGB(200, 215, 255)
       .Range("C5").Value = Sins
       
       .Range("B6").Value = "用具購入代"
       .Range("B6").Interior.Color = RGB(200, 215, 255)
       .Range("C6").Value = Sitm
       
       .Range("B7").Value = "その他"
       .Range("B7").Interior.Color = RGB(200, 215, 255)
       .Range("C7").Value = Soth
       
       .Range("B9").Value = "支出"
              
       .Range("B10").Value = "支出合計"
       .Range("B10").Interior.Color = RGB(255, 150, 150)
       .Range("C10").Value = Srev
       
       .Range("B12").Value = "最終売り上げ"
       .Range("C12").Value = Stoday
       .Range("B13").Value = "袋の中の金額"
       .Range("C13").Value = Sum.Value
       
       
        .Columns(2).ColumnWidth = 18.75
        .Columns(3).ColumnWidth = 10
        
        
        .Range("B16").Value = "金種表(経理用)"
        .Range("B17").Value = "10,000"
        .Range("B18").Value = "5,000"
        .Range("B19").Value = "1,000"
        .Range("B20").Value = "500"
        .Range("B21").Value = "100"
        .Range("B22").Value = "50"
        .Range("B23").Value = "10"
        .Range("B24").Value = "5"
        .Range("B25").Value = "1"
        .Range("B26").Value = "合計"
        
        
        .Range("B17:C26").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Range("B17:B26").Borders(xlEdgeRight).LineStyle = xlDouble
        .Range("B17:C17").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B18:C18").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B19:C19").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B20:C20").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B21:C21").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B22:C22").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B23:C23").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B24:C24").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B25:C25").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
       End With
            
        Workbooks(M & "月収支表.xlsx").Sheets(cnt).Name = Y & "年" & M & "月" & D & "日" & "収支表"
     
     'エクセルシートのリセット
     
     Call レジ開始
     
     Workbooks(M & "月売上管理.xlsx").Save
     Workbooks(M & "月収支表.xlsx").Save
     
     Unload CloseNextForm
     
     '収支表の印刷
     
     'MsgBox "収支表の印刷をします。プリンターの電源をonにしてOKボタンを押してください。"
     
      'Workbooks(M & "月収支表.xlsx").Sheets(Y & "年" & M & "月" & D & "日" & "収支表").PrintOut
     
     '締め作業の終了
     
     MsgBox "レジ締め終了しました。すべてのエクセルのウィンドウを閉じてからタブレットをシャットダウンしてください。"
     
     Workbooks("レジ.xlsm").Save
      
   End If
End Sub

