VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReverseForm 
   Caption         =   "Reverse"
   ClientHeight    =   9270.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15180
   OleObjectBlob   =   "ReverseForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ReverseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MsgBox "［閉じる］ボタンを使用してください"
        Cancel = True
    End If
End Sub

Private Sub AlredyButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim i As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      i = Agecell.Offset(0, 1).Value
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "支払い済"
            .Cells(LastRow, 4).Value = 0 & 6 & i
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = 0
        End With
         
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
            End With
   
            PriceBox3.Text = S
    End If
        
End Sub

Private Sub CancelButton_Click()
  Dim CancelAmount As Long
  Dim i As Long
  Dim j As Long

  Number.Text = ""
  NameTextBox.Text = ""
  PriceBox.Value = 0
  PriceBox2.Value = 0
  PriceBox3.Value = 0
  AgeCombo.ListIndex = 0
  SwordCombo.ListIndex = 0
  LessonTime.ListIndex = 0
  LessonTime2.ListIndex = 0
  CarTime.ListIndex = 0
  FamilyCombo.ListIndex = 0
  U64Option.Value = True
  CheckBox.Value = False
  
  With Worksheets("登録用シート")
      CancelAmount = .Cells(Rows.Count, 3).End(xlUp).Row
      For i = 3 To CancelAmount
         For j = 3 To 10
            .Cells(i, j).ClearContents
         Next
      Next
  End With
End Sub

Private Sub CarButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim i As Long
   Dim j As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      i = Agecell.Offset(0, 1).Value
      j = CarTime.Value / 3
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "駐車場"
            .Cells(LastRow, 4).Value = 10 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = 500 * j
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
            End With
   
            PriceBox3.Text = S
    End If
        
End Sub

Private Sub ClassButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim i As Long
   Dim k As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      i = Agecell.Offset(0, 1).Value
      k = FamilyCombo.Value
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "初心者教室"
            .Cells(LastRow, 4).Value = 0 & 4 & i
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            If i = 4 Then
               .Cells(LastRow, 9).Value = 2000 * k
            Else
               .Cells(LastRow, 9).Value = 2500
            End If
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
            End With
   
            PriceBox3.Text = S
    End If
End Sub

Private Sub ClassMButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim i As Long
   Dim k As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      i = Agecell.Offset(0, 1).Value
      k = FamilyCombo.Value
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "初心者教室(月謝払い)"
            .Cells(LastRow, 4).Value = 0 & 5 & i
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            If i = 4 Then
               .Cells(LastRow, 9).Value = 7750 * k
            Else
               .Cells(LastRow, 9).Value = 9000
            End If
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub CloseCommand_Click()

If Worksheets("登録用シート").Range("C3").Value = "" Then
   Unload ReverseForm
   
Else
   MsgBox "登録用シートに登録が未完了のデータがあります。登録か一括取消をして操作を完了させてからレジを閉じてください。"
   
End If

End Sub

Private Sub CommandButton2_Click()
Unload Me
Call メニューの表示

End Sub

Private Sub DownButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim Price As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
  ElseIf PriceBox2 = 0 Then
     MsgBox "金額を入力してください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      Price = PriceBox2.Text
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "値引き"
            .Cells(LastRow, 4).Value = 17 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = Price * (-1)
        End With
        PriceBox2 = 0
        
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub ExperienceButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "初回体験"
            .Cells(LastRow, 4).Value = 0 & 9 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = 1000
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub FirstButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "入会金"
            .Cells(LastRow, 4).Value = 0 & 8 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = 2000
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub InsuranceButton11_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 4).Value = 11 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            If U15Option.Value = True Then
               .Cells(LastRow, 3).Value = "スポーツ保険(中学生以下)"
               .Cells(LastRow, 9).Value = 1000
            ElseIf U64Option = True Then
               .Cells(LastRow, 3).Value = "スポーツ保険(64歳以下)"
               .Cells(LastRow, 9).Value = 2000
            Else
               .Cells(LastRow, 3).Value = "スポーツ保険(65歳以上)"
               .Cells(LastRow, 9).Value = 1400
            End If
               
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub ItemButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim Price As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
  ElseIf PriceBox = 0 Then
     MsgBox "金額を入力してください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      Price = PriceBox.Text
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "用具購入代"
            .Cells(LastRow, 4).Value = 12 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = Price
        End With
        PriceBox = 0
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub LackButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim Price As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
  ElseIf PriceBox2 = 0 Then
     MsgBox "金額を入力してください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      Price = PriceBox2.Text
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "不足"
            .Cells(LastRow, 4).Value = 16 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = Price * (-1)
        End With
        PriceBox2 = 0
        
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub LessonpracticeButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim i As Long
   Dim j As Long
   Dim S As Long
   Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
   i = Agecell.Offset(0, 1).Value
    
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
   ElseIf i = 4 Then
        MsgBox "家族割引ではご利用いただけません"
   
   Else
      j = LessonTime2.Value
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "個人レッスン＋練習会"
            .Cells(LastRow, 4).Value = 0 & 3 & i
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            If i = 1 Then
               .Cells(LastRow, 9).Value = 2500 * j + 500
            ElseIf i = 2 Then
               .Cells(LastRow, 9).Value = 2000 * j + 300
            ElseIf i = 3 Then
               .Cells(LastRow, 9).Value = 2000 * j

            End If
            
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
        
End Sub

Private Sub LessonButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim i As Long
   Dim j As Long
   Dim S As Long
   Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
   i = Agecell.Offset(0, 1).Value
    
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
   ElseIf i = 4 Then
     MsgBox "家族割引ではご利用いただけません。"
   Else
      j = LessonTime.Value
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "個人レッスン"
            .Cells(LastRow, 4).Value = 0 & 2 & i
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            If i = 1 Then
               .Cells(LastRow, 9).Value = 2500 * j
            Else
               .Cells(LastRow, 9).Value = 2000 * j
            End If
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
    
End Sub

Private Sub Money2Button_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim Price As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
  ElseIf PriceBox2 = 0 Then
     MsgBox "金額を入力してください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      Price = PriceBox2.Text
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "出金(その他)"
            .Cells(LastRow, 4).Value = 18 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = Price * (-1)
        End With
        PriceBox2 = 0
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub MoneyButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim Price As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
  ElseIf PriceBox = 0 Then
     MsgBox "金額を入力してください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      Price = PriceBox.Text
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "入金(その他)"
            .Cells(LastRow, 4).Value = 14 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = Price
        End With
        PriceBox = 0
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub OtherClassesButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim Price As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
  ElseIf PriceBox = 0 Then
     MsgBox "金額を入力してください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      Price = PriceBox.Text
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "教室/講演会"
            .Cells(LastRow, 4).Value = 13 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = Price
        End With
        PriceBox = 0
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub PracticeButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim i As Long
   Dim S As Long
   Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
   i = Agecell.Offset(0, 1).Value
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
   ElseIf i = 4 Then
      MsgBox "家族割引では利用できません"
      
   Else

      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "練習会"
            .Cells(LastRow, 4).Value = 0 & 1 & i
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            If i = 2 Then
               .Cells(LastRow, 9).Value = 300

            Else
               .Cells(LastRow, 9).Value = 1000
            End If
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
        
End Sub

Private Sub ReadButton1_Click()
   Dim No As Long
   Dim Nocell As Range
   Dim Name As String
   
   If Number = "" Then
      MsgBox "会員番号を入力してください。"

   Else
  
      No = Number.Text
      With Worksheets("名簿")
      Set Nocell = .Range("B3", .Range("B3").End(xlDown)).Find(what:=No, LookAt:=xlWhole)
      
      NameTextBox = Nocell.Offset(0, 1).Value
      End With
      
      
   End If
        
End Sub

Private Sub RefundButton_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim Price As Long
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
  ElseIf PriceBox2 = 0 Then
     MsgBox "金額を入力してください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      Price = PriceBox2.Text
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "返金"
            .Cells(LastRow, 4).Value = 15 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = Price * (-1)
        End With
        PriceBox2 = 0
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub

Private Sub ResultButton_Click()
  Dim RegistrationAmount As Long
  Dim LastRow As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  
  If Worksheets("登録用シート").Range("C3").Value = "" Then
     MsgBox "登録するものがありません。" & vbCrLf & "利用詳細を入力してから再度登録をしてください。", vbCritical
  Else
      l = MsgBox("戻し操作をしてもよろしいですか？", vbOKCancel)
      
      If l = 1 Then
      
     With Worksheets("登録用シート")
         RegistrationAmount = .Cells(Rows.Count, 3).End(xlUp).Row
         LastRow = Worksheets("日計取引表").Cells(Rows.Count, 2).End(xlUp).Row
         For i = 3 To RegistrationAmount
            For j = 3 To 10
               Worksheets("日計取引表").Cells(LastRow - 2 + i, j).Value = .Cells(i, j).Value
               Worksheets("日計取引表").Cells(LastRow - 2 + i, j).Interior.ColorIndex = 3
            Next
            k = Worksheets("日計取引表").Cells(LastRow - 2 + i, 9).Value
           Worksheets("日計取引表").Cells(LastRow - 2 + i, 2).Value = LastRow - 4 + i
           Worksheets("日計取引表").Cells(LastRow - 2 + i, 2).Interior.ColorIndex = 3
           Worksheets("日計取引表").Cells(LastRow - 2 + i, 9).Value = k * -1
         Next
     End With
     
     End If
  
  Call CancelButton_Click
  
  Workbooks("レジ.xlsm").Save
   
  
  End If
End Sub

Private Sub SumButton_Click()
   Dim S As Long
   
   With Worksheets("登録用シート")
   S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
   End With
   
   PriceBox3.Text = S
End Sub

Private Sub WarterButton8_Click()
   Dim LastRow As Long
   Dim Agecell As Range
   Dim S As Long
      
   If Number = "" And Name = "" And CheckBox = False Then
      MsgBox "会員番号と名前を入力してください。"
   
   ElseIf Number = "" And CheckBox = False Then
      MsgBox "会員番号を入力するか、新規にチェックをしてください。"
      
   Else
      Set Agecell = Worksheets("コンボボックス用リスト").Range("B3:B6").Find(what:=AgeCombo.Text, LookAt:=xlWhole)
      With Worksheets("登録用シート")
         LastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
            .Cells(LastRow, 3).Value = "アクアクララ"
            .Cells(LastRow, 4).Value = 0 & 7 & 0
            .Cells(LastRow, 5).Value = Number.Text
            .Cells(LastRow, 6).Value = NameTextBox.Text
            .Cells(LastRow, 7).Value = AgeCombo.Text
            .Cells(LastRow, 8).Value = SwordCombo.Text
            .Cells(LastRow, 9).Value = 100
        End With
        With Worksheets("登録用シート")
            S = Application.WorksheetFunction.Sum(.Range("I3", .Range("I3").End(xlDown)))
        End With
   
        PriceBox3.Text = S
    End If
End Sub
