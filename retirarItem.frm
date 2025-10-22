VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} retirarItem 
   Caption         =   "RETIRAR ITENS"
   ClientHeight    =   9030.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13020
   OleObjectBlob   =   "retirarItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "retirarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim i As Long
    Dim ultimaLinha As Long
    Dim valorRetirado As Long
    Dim linhaVazia As Long
    
    If TextBox1 = "" Or ComboBox1 = "" Then
        MsgBox "Preencha todos os campos"
        ComboBox1.SetFocus
        Exit Sub
    End If
    
    valorRetirado = Val(TextBox1.Value)
    
    ultimaLinha = Sheets("ITEM_CADASTRADOS").Cells(Sheets("ITEM_CADASTRADOS").Rows.Count, 3).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If ComboBox1.Value = Sheets("ITEM_CADASTRADOS").Cells(i, 1) Then
    
            If Sheets("ITEM_CADASTRADOS").Cells(i, 3).Value < valorRetirado Then
                MsgBox "Quantidade insuficiente em estoque!"
                Exit Sub
            End If
            
            Sheets("ITEM_CADASTRADOS").Cells(i, 3) = Sheets("ITEM_CADASTRADOS").Cells(i, 3) - valorRetirado
           Exit For
        End If
    Next i
        
    With Sheets("HISTORICO")
        linhaVazia = 2 ' começa a busca na linha 2
        
        Do While .Cells(linhaVazia, 4).Value <> ""
            linhaVazia = linhaVazia + 1
        Loop
        
        If Not ComboBox1.Value = Sheets("ITEM_CADASTRADOS").Cells(i, 1) Then
            MsgBox "Este item não existe no estoque"
            
        Else
        ' registra os dados
        .Cells(linhaVazia, 1).Value = "RETIRADA"
        .Cells(linhaVazia, 2).Value = ComboBox1.Value
        .Cells(linhaVazia, 3).Value = TextBox1.Value
        .Cells(linhaVazia, 4).Value = Date
        .Cells(linhaVazia, 5).Value = Time
        End If
    End With
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()
 Dim i As Long
 Dim ultimaLinha As Long
 ComboBox1.MatchEntry = fmMatchEntryNone
 
 ultimaLinha = Sheets("ITEM_CADASTRADOS").Cells(Sheets("ITEM_CADASTRADOS").Rows.Count, 1).End(xlUp).Row '//essa utiliza a planilha item_cadastrado independentemente
 'ultimaLinha = Sheets("ITEM_CADASTRADOS").Cells(Rows.Count, 1).End(xlUp).Row // Essa pega a planilha ativa
 
 ComboBox1.Clear
 
    For i = 2 To ultimaLinha
        ComboBox1.AddItem Sheets("ITEM_CADASTRADOS").Cells(i, 1).Value
    Next i
    
End Sub

Private Sub ComboBox1_Change()
Dim i As Long
Dim j As Long
Dim ultimaLinha2 As Long
Dim itemProcurado As String
Dim valorCelula As String

itemProcurado = ComboBox1.Value


ListBox2.Clear
ultimaLinha2 = Sheets("ITEM_CADASTRADOS").Cells(Sheets("ITEM_CADASTRADOS").Rows.Count, 1).End(xlUp).Row '//essa utiliza a planilha item_cadastrado independentemente

    For i = 2 To ultimaLinha2
        valorCelula = Sheets("ITEM_CADASTRADOS").Cells(i, 1).Value
    
            If InStr(1, LCase(valorCelula), LCase(itemProcurado), vbTextCompare) > 0 Then
            ListBox2.AddItem valorCelula
            End If
    Next i
End Sub
