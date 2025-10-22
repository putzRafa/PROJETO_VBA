VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} telaInicial 
   Caption         =   "TELA INICIAL"
   ClientHeight    =   9150.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15570
   OleObjectBlob   =   "telaInicial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "telaInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ToggleButton1_Click()
cadastrarItem.Show
End Sub

Private Sub ToggleButton2_Click()
'estoque
Sheets("ITEM_CADASTRADOs").Activate
End Sub

Private Sub ToggleButton3_Click()
fazerPedido.Show
End Sub
Private Sub ToggleButton5_Click()
Sheets("HISTORICO").Activate
End Sub

Private Sub UserForm_Click()
ToggleButton1.SetFocus
End Sub
