Attribute VB_Name = "OpenAct"
Option Explicit
Sub ���W�J�n()
        
  Dim CancelAmount As Long
  Dim i As Long
  Dim j As Long
  
  With Workbooks("���W.xlsm").Worksheets("�o�^�p�V�[�g")
      CancelAmount = .Cells(Rows.Count, 3).End(xlUp).Row
      For i = 3 To CancelAmount
         For j = 3 To 10
            .Cells(i, j).ClearContents
            .Cells(i, j).ClearFormats
         Next
      Next
  End With
  
  With Workbooks("���W.xlsm").Worksheets("���v����\")
      CancelAmount = .Cells(Rows.Count, 2).End(xlUp).Row
      For i = 3 To CancelAmount
         For j = 2 To 10
            .Cells(i, j).ClearContents
            .Cells(i, j).ClearFormats
         Next
      Next
  End With
  
End Sub
