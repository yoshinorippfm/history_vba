Attribute VB_Name = "OpenAct"
Option Explicit
Sub レジ開始()
        
  Dim CancelAmount As Long
  Dim i As Long
  Dim j As Long
  
  With Workbooks("レジ.xlsm").Worksheets("登録用シート")
      CancelAmount = .Cells(Rows.Count, 3).End(xlUp).Row
      For i = 3 To CancelAmount
         For j = 3 To 10
            .Cells(i, j).ClearContents
            .Cells(i, j).ClearFormats
         Next
      Next
  End With
  
  With Workbooks("レジ.xlsm").Worksheets("日計取引表")
      CancelAmount = .Cells(Rows.Count, 2).End(xlUp).Row
      For i = 3 To CancelAmount
         For j = 2 To 10
            .Cells(i, j).ClearContents
            .Cells(i, j).ClearFormats
         Next
      Next
  End With
  
End Sub
