VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Insert 
   Caption         =   "Insert"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5100
   OleObjectBlob   =   "Insert.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
   If NameBox = "" Or KanaBox = "" Or NoBox = "" Then
      MsgBox "�S�Ă̗��ɓ��͂��Ă���ēx�o�^���Ă�������"
      
   Else
      Dim Name As String
      Dim Kana As String
      Dim No As String
      Dim LastRow As Long
      Dim Ins1 As String
      Dim Ins2 As String
      Dim Ins3 As String
      
      Name = NameBox.Text
      Kana = KanaBox.Text
      No = NoBox.Text
      Ins1 = Unconfirmed.Value
      Ins2 = Bought.Value
      Ins3 = Exemption.Value
      
      With Worksheets("����")
         LastRow = .Cells(Rows.Count, 2).End(xlUp).Row + 1
         
         .Cells(LastRow, 2).Value = No
         .Cells(LastRow, 3).Value = Name
         .Cells(LastRow, 4).Value = Kana
         
         If Ins1 = True Then
            .Cells(LastRow, 5).Value = "���m�F"
         ElseIf Ins2 = True Then
            .Cells(LastRow, 5).Value = "������"
         Else
            .Cells(LastRow, 5).Value = "�Ə�"
         End If
         
         
     End With
     
     Unload Insert
    
  End If

End Sub

