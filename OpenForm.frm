VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpenForm 
   Caption         =   "OpenForm"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "OpenForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "OpenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
  If TextBox1 = "open" Then
  
    Call ���W�J�n
 
    Unload OpenForm
    
    RegisterForm.Show
  
  Else
     MsgBox "�p�X���[�h���Ⴂ�܂�"
     Unload OpenForm
     
End If
     
End Sub
