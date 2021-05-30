VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuForm 
   Caption         =   "Menu"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5715
   OleObjectBlob   =   "MenuForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

   OpenForm.Show
   
   Unload MenuForm

End Sub

Private Sub CommandButton3_Click()
   ResearchForm.Show
   Unload MenuForm
End Sub

Private Sub CommandButton4_Click()

   CloseForm.Show
   
   Unload MenuForm
End Sub

Private Sub CommandButton5_Click()
   Insert.Show
   
   Unload MenuForm
End Sub

Private Sub CommandButton6_Click()
   Unload Me
   Call 戻しの表示
   
End Sub

Private Sub RegisterCommand_Click()
      
      Unload Me
      
      Call レジの表示
      
End Sub
