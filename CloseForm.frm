VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CloseForm 
   Caption         =   "CloseForm"
   ClientHeight    =   2380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "CloseForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CloseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
   Dim S As Long
   
   If TextBox1 = "close" Then
      With Worksheets("���v����\")
      S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
      End With
      CloseNextForm.TodaySum.Text = S & "�~�ł��B"
      
      CloseNextForm.Show
      
      Unload CloseForm
   Else
      MsgBox "�p�X���[�h���Ⴂ�܂��B"
      Unload CloseForm
   End If
End Sub

