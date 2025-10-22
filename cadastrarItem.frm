VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadastrarItem 
   Caption         =   "UserForm1"
   ClientHeight    =   9105.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12405
   OleObjectBlob   =   "cadastrarItem.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cadastrarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim proximaLinha As Long

    If TextBox1 = "" Or TextBox2 = "" Or TextBox3 = "" Then
    MsgBox "Preencha todos os campos"
    TextBox1.SetFocus
    Exit Sub
    End If
    
    'Encontra a próxima linha vazia na coluna A, B, C
    proximaLinha = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    Cells(proximaLinha, 1).Value = TextBox1.Text
    Cells(proximaLinha, 2).Value = TextBox2.Text
    Cells(proximaLinha, 3).Value = TextBox3.Text
    
        'ADICIONA O ITEM NO HISTORICO
        Sheets("HISTORICO").Cells(proximaLinha, 1).Value = "ENTRADA"
        Sheets("HISTORICO").Cells(proximaLinha, 2).Value = TextBox1.Text
        Sheets("HISTORICO").Cells(proximaLinha, 3).Value = TextBox3.Text
        
            'ADICIONA DATA E HORA
            Sheets("HISTORICO").Cells(proximaLinha, 4).Value = Date
            Sheets("HISTORICO").Cells(proximaLinha, 5).Value = Time
            
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox1.SetFocus
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub ToggleButton4_Click()
retirarItem.Show
End Sub
Private Sub UserForm_Click()
End Sub

Private Sub UserForm_Initialize()
    Label1.BackColor = RGB(0, 184, 79)
    Label1.ForeColor = RGB(255, 255, 255)
    
    Label2.BackColor = RGB(0, 184, 79)
    Label2.ForeColor = RGB(255, 255, 255)
    
    Label3.BackColor = RGB(0, 184, 79)
    Label3.ForeColor = RGB(255, 255, 255)
End Sub
