VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CloseNextForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   OleObjectBlob   =   "CloseNextForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CloseNextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
   If Sum.Text = "" Or Coach.Text = "" Then
      MsgBox "�܂̒��̋��z�܂��̓��W�S���R�[�`�̖��O�����͂���Ă��܂���B"
   
   Else
     Dim cnt As Long ' �R�s�[��u�b�N�̃V�[�g�̖���
     Dim Y As Long
     Dim M As Long
     Dim D As Long
     Dim shp As Shape
     Dim Stoday As Long
     Dim LastRow As Long
     
     
     Y = Year(Now)
     M = Month(Now)
     D = Day(Now)
     
     
     Workbooks.Open "\\192.168.3.63\share\garden\" & Y & "�N����Ǘ�\" & M & "������Ǘ�.xlsx"
     Workbooks.Open "\\192.168.3.63\share\garden\" & Y & "�N���x�\\" & M & "�����x�\.xlsx"
     
     
    ' ����Ǘ��̃V�[�g�̕ۑ�
     cnt = Workbooks(M & "������Ǘ�.xlsx").Sheets.Count
     Workbooks("���W.xlsm").Worksheets("���v����\").Copy after:=Workbooks(M & "������Ǘ�.xlsx").Sheets(cnt)
  
     For Each shp In Workbooks(M & "������Ǘ�.xlsx").Sheets("���v����\").Shapes
       shp.Delete
     Next shp
     
     With Workbooks("���W.xlsm").Sheets("���v����\")
         Stoday = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
     End With
     
     With Workbooks(M & "������Ǘ�.xlsx").Sheets("���v����\")
        LastRow = .Cells(Rows.Count, 9).End(xlUp).Row + 1
        .Range("C1").Value = Y & "�N" & M & "��" & D & "��" & "����"
        .Cells(1, 4).Value = Coach.Text
        .Cells(LastRow, 8).Value = "���v"
        .Cells(LastRow, 9).Value = Stoday
     End With
     
     Workbooks(M & "������Ǘ�.xlsx").Sheets("���v����\").Name = Y & "�N" & M & "��" & D & "��" & "����"
     
     
     
     '���x�\�̍쐬
     Dim i As Variant
     Dim S As Long
     Dim Scar As Long
     Dim Sins As Long
     Dim Sitm As Long
     Dim Soth As Long
     Dim Srev As Long
     Dim W As Single
    
     
     S = Scar = Sins = Sitm = Soth = Srev = W = 0
     
     
     cnt = Workbooks(M & "�����x�\.xlsx").Sheets.Count + 1
     
     With Workbooks("���W.xlsm").Worksheets("���v����\")
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
    
    Workbooks(M & "�����x�\.xlsx").Worksheets.Add after:=Worksheets(cnt - 1)
    
    
    With Workbooks(M & "�����x�\.xlsx").Sheets(cnt)
       .Range("B1").Value = Y & "�N" & M & "��" & D & "��" & "���x�\"
       .Range("C1").Value = Coach.Text
              
       .Range("B2").Value = "����"
       
       .Range("B3").Value = "�t�F���V���O����"
       .Range("B3").Interior.Color = RGB(200, 215, 255)
       .Range("C3").Value = S
       
       .Range("B4").Value = "���ԏ�"
       .Range("B4").Interior.Color = RGB(200, 215, 255)
       .Range("C4").Value = Scar
       
       .Range("B5").Value = "�X�|�[�c�ی�"
       .Range("B5").Interior.Color = RGB(200, 215, 255)
       .Range("C5").Value = Sins
       
       .Range("B6").Value = "�p��w����"
       .Range("B6").Interior.Color = RGB(200, 215, 255)
       .Range("C6").Value = Sitm
       
       .Range("B7").Value = "���̑�"
       .Range("B7").Interior.Color = RGB(200, 215, 255)
       .Range("C7").Value = Soth
       
       .Range("B9").Value = "�x�o"
              
       .Range("B10").Value = "�x�o���v"
       .Range("B10").Interior.Color = RGB(255, 150, 150)
       .Range("C10").Value = Srev
       
       .Range("B12").Value = "�ŏI����グ"
       .Range("C12").Value = Stoday
       .Range("B13").Value = "�܂̒��̋��z"
       .Range("C13").Value = Sum.Value
       
       
        .Columns(2).ColumnWidth = 18.75
        .Columns(3).ColumnWidth = 10
        
        
        .Range("B16").Value = "����\(�o���p)"
        .Range("B17").Value = "10,000"
        .Range("B18").Value = "5,000"
        .Range("B19").Value = "1,000"
        .Range("B20").Value = "500"
        .Range("B21").Value = "100"
        .Range("B22").Value = "50"
        .Range("B23").Value = "10"
        .Range("B24").Value = "5"
        .Range("B25").Value = "1"
        .Range("B26").Value = "���v"
        
        
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
            
        Workbooks(M & "�����x�\.xlsx").Sheets(cnt).Name = Y & "�N" & M & "��" & D & "��" & "���x�\"
     
     '�G�N�Z���V�[�g�̃��Z�b�g
     
     Call ���W�J�n
     
     Workbooks(M & "������Ǘ�.xlsx").Save
     Workbooks(M & "�����x�\.xlsx").Save
     
     Unload CloseNextForm
     
     '���x�\�̈��
     
     'MsgBox "���x�\�̈�������܂��B�v�����^�[�̓d����on�ɂ���OK�{�^���������Ă��������B"
     
      'Workbooks(M & "�����x�\.xlsx").Sheets(Y & "�N" & M & "��" & D & "��" & "���x�\").PrintOut
     
     '���ߍ�Ƃ̏I��
     
     MsgBox "���W���ߏI�����܂����B���ׂẴG�N�Z���̃E�B���h�E����Ă���^�u���b�g���V���b�g�_�E�����Ă��������B"
     
     Workbooks("���W.xlsm").Save
      
   End If
End Sub

