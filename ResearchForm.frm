VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResearchForm 
   Caption         =   "ResearchForm"
   ClientHeight    =   2840
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5670
   OleObjectBlob   =   "ResearchForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ResearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

      Dim Name1 As String
      Dim Name2 As String
      Dim Name As String
      Dim No As Long
      Dim Nocell As Range
      
      'Name1 = NameBox1.Text
      'Name2 = NameBox2.Text
      Name = NameBox1.Text & " " & NameBox2.Text
      With Worksheets("名簿")
      Set Nocell = .Range("D3", .Range("D3").End(xlDown)).Find(what:=Name, LookAt:=xlWhole)
      
      If Nocell Is Nothing Then
         MsgBox "会員番号が見つかりませんでした。"
         
      Else
      
      'Nocell.Offset(0, -2).Value
      
      MsgBox "会員番号は" & Nocell.Offset(0, -2).Value & "です。"
      End If
      End With
      
      Unload ResearchForm
End Sub

