VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fazerPedido 
   Caption         =   "SOLICITAÇÃO DE COMPRA"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   OleObjectBlob   =   "fazerPedido.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fazerPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
End Sub
Private Sub LimparTexto()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                ctrl.Value = ""
            Case "ComboBox"
                ctrl.Value = ""
        End Select
    Next ctrl
End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Initialize()
    'PRA INICIALIZAÇÃO GERAL USA INITIALIZATE. SÓ COLOCA DENTRO DO COMANDO PARA EVENTOS
    ' INTERNOS
    Dim i As Long
    Dim j As Long
    Dim ultimaLinha As Long
    Dim ultimaLinha2 As Long
    
    ComboBox1.Clear
    
    Label1.BackColor = RGB(0, 184, 79)
    Label1.ForeColor = RGB(255, 255, 255)
    
    Label3.BackColor = RGB(0, 184, 79)
    Label3.ForeColor = RGB(255, 255, 255)
    
    Label4.BackColor = RGB(0, 184, 79)
    Label4.ForeColor = RGB(255, 255, 255)
    
    Label5.BackColor = RGB(0, 184, 79)
    Label5.ForeColor = RGB(255, 255, 255)
    
    Label6.BackColor = RGB(0, 184, 79)
    Label6.ForeColor = RGB(255, 255, 255)
    
    Label7.BackColor = RGB(0, 184, 79)
    Label7.ForeColor = RGB(255, 255, 255)
    
    ultimaLinha = Sheets("ITEM_CADASTRADOS").Cells(Rows.Count, 3).End(xlUp).Row
    
    For i = 1 To ultimaLinha
        If Sheets("ITEM_CADASTRADOS").Cells(i, 3).Value <= 10 Then
        ComboBox1.AddItem Sheets("ITEM_CADASTRADOS").Cells(i, 1).Value
        End If
    Next i
    
    TextBox3.Value = Date
    TextBox3.Locked = True
    
End Sub
Private Sub ComboBox1_Change()
    Dim j As Long
    Dim ultimaLinha2 As Long

    ultimaLinha2 = Sheets("ITEM_CADASTRADOS").Cells(Rows.Count, 1).End(xlUp).Row
    
    For j = 1 To ultimaLinha2
        If Sheets("ITEM_CADASTRADOS").Cells(j, 1).Value = ComboBox1.Value Then
        Sheets("SOLICITACAO_COMPRA").Cells(29, 4).Value = Sheets("ITEM_CADASTRADOS").Cells(j, 2).Value
        End If
    Next j
End Sub
Private Sub CommandButton1_Click()

'PREENCHENDO O FORMULARIO
    Sheets("SOLICITACAO_COMPRA").Cells(4, 4).Value = TextBox2 'NOME DO SOLICITANTE
        Sheets("SOLICITACAO_COMPRA").Cells(5, 4).Value = TextBox3 'DATA DE SOLICITAÇÃO
            Sheets("SOLICITACAO_COMPRA").Cells(6, 4).Value = TextBox4 'DATA DE ENTREGA
                Sheets("SOLICITACAO_COMPRA").Cells(11, 3) = ComboBox1.Value 'NOME DO ITEM
                    Sheets("SOLICITACAO_COMPRA").Cells(12, 3) = TextBox1.Value 'QUANTIDADE
                        Sheets("SOLICITACAO_COMPRA").Cells(7, 4) = TextBox5.Value 'JUSTIFICATIVA
                        
                        
    If Not TextBox2 = "" And Not TextBox4 = "" And Not ComboBox1 = "" And Not TextBox1 = "" And Not TextBox5 = "" Then
        EnviaEmail
    Else
        MsgBox "Preencha todos os campos"
    End If

    LimparTexto
    TextBox3.Value = Date
    TextBox3.Locked = True
End Sub
Private Sub EnviaEmail()
    Dim ws As Worksheet
    Dim wbTemp As Workbook
    Dim caminhoTemp As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    
    ' Define a planilha que será enviada
    Set ws = ThisWorkbook.Sheets("SOLICITACAO_COMPRA")
    
    ' Caminho temporário para salvar a cópia
    caminhoTemp = Environ("TEMP") & "\" & ws.Name & ".xlsx"
    
    ' Cria uma cópia apenas da planilha desejada
    ws.Copy
    Set wbTemp = ActiveWorkbook
    
    ' Salva o arquivo temporário
    wbTemp.SaveAs Filename:=caminhoTemp, FileFormat:=51 ' 51 = .xlsx
    
    ' Cria o e-mail no Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)
    
    With outlookMail
        .To = "rafaellimafarias2802@gmail.com"
        .Subject = "Envio da planilha " & ws.Name
        .Body = "Segue em anexo a planilha solicitada."
        .Attachments.Add caminhoTemp
        .Send ' Envia direto (use .Display para revisar antes)
    End With
    
    ' Fecha o arquivo temporário sem salvar
    wbTemp.Close False
    
    ' Exclui o arquivo temporário (opcional)
    Kill caminhoTemp
    
    ' Limpa objetos
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    Set wbTemp = Nothing
    
    MsgBox "Planilha '" & ws.Name & "' enviada com sucesso!", vbInformation
End Sub
