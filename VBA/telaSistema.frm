VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} telaSistema 
   Caption         =   "Sistema"
   ClientHeight    =   6960
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9204.001
   OleObjectBlob   =   "telaSistema.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "telasistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub BTacomodacoesdisponiveis_Click()

    acomodacoesDisp.Show
    
End Sub
Private Sub BTcheckout_Click()

    Dim idCliente As Integer
    Dim idAcomodacao As Integer
    Dim totalLinhasQA As Integer
    Dim totalLinhasAC
    Dim i As Integer
    
    idCliente = TextId.Value
    idAcomodacao = TextIdAcomodacao.Value
    
    totalLinhasQA = Pquartosalugados.Range("A1").CurrentRegion.Rows.Count
    
    For i = 2 To totalLinhasQA
        If idCliente = Pquartosalugados.Cells(i, 1).Value Then
            Pquartosalugados.Cells(i, 9).Value = "Checkout"
            Exit For
        End If
    Next
    
    totalLinhasAC = Pacomodacoes.Range("A1").CurrentRegion.Rows.Count
    
    For i = 2 To totalLinhasAC
        If idAcomodacao = Pacomodacoes.Cells(i, 1).Value Then
            Pacomodacoes.Cells(i, 6).Value = "Disponível"
            Exit For
        End If
    Next
    
    MsgBox "Checkout concluído!"
    Call BTnovoregistro_Click
    
End Sub
Private Sub BTconsultaralugados_Click()

    consultarAlugados.Show
    
    BTcheckout.Enabled = True 'Habilita o botão
    BTreservar.Enabled = False ' Desabilita o botao
    
End Sub
Private Sub BTnovoregistro_Click()

'dispara ao clicar no botao "Novo Registro
    Dim i As Integer
    
    'Coloca no campo ID o próximo ID
    'Pega o valor da célula ID
    TextId.Value = Pquartosalugados.Range("ID").Value
    
    'Apaga os valores dos controles
    
    TextCliente.Value = ""
    TextContato.Value = ""
    TextIdAcomodacao.Value = ""
    TextQtdeCama.Value = ""
    TextQtdeQuartos.Value = ""
    TextQtdeBanheiros.Value = ""
    TextQtdeDiaria.Value = ""
    CBNumDias.Value = ""
    TextCheckin.Value = ""
    TextCheckout.Value = ""
    TextTotal.Value = ""
    '--------------------------------
    
    'Coloca a lista de dias
    CBNumDias.Clear
    For i = 1 To 30
        CBNumDias.AddItem i
    Next
    
    BTcheckout.Enabled = False 'Desabilita o botão
    BTreservar.Enabled = True ' Habilita o botao
    
End Sub
Private Sub BTreservar_Click()

    Dim nLin As Integer
    Dim Id As Integer
    Dim checkout As Date
    Dim checkin As Date
    Dim codAcom As Integer
    Dim totalLin As Integer
    
    'Verifica se os itens obrigatórios estão preenchidos
    If TextCliente.Value = "" Then
        MsgBox "Preencha o campo Cliente"
        TextCliente.SetFocus 'Posiciona o cursor no controle
        Exit Sub 'Abortar macro
    
    ElseIf TextContato.Value = "" Then
        MsgBox "Preencha o campo Contato"
        TextContato.SetFocus 'Posicionar o cursor no controle
        Exit Sub
    
    ElseIf CBNumDias.Value = "" Then
        MsgBox "Selecione a quantidade de dias"
        CBNumDias.SetFocus 'Posicionar o cursor no controle
        Exit Sub
    End If
    
    MsgBox "Registro salvo com sucesso!"
    
    nLin = Pquartosalugados.Range("A1").CurrentRegion.Rows.Count + 1
    
    'Atribui os valores a variavel de data
    checkout = telasistema.TextCheckout.Value
    checkin = telasistema.TextCheckin.Value
    
    
    'A variavel ID armazena o ID do registro qque está na celula ID
    Id = Pquartosalugados.Range("ID").Value
    Pquartosalugados.Range("ID").Value = Id + 1 'soma + 1 a cada novo registro
    
    'Cadastra nas células os valores q estão nos controles
    Pquartosalugados.Cells(nLin, 1).Value = telasistema.TextId.Value
    Pquartosalugados.Cells(nLin, 2).Value = telasistema.TextCliente.Value
    Pquartosalugados.Cells(nLin, 3).Value = "https://wa.me/55" + telasistema.TextContato.Value
    Pquartosalugados.Cells(nLin, 4).Value = telasistema.TextIdAcomodacao.Value
    Pquartosalugados.Cells(nLin, 5).Value = telasistema.CBNumDias.Value
    Pquartosalugados.Cells(nLin, 6).Value = checkin
    Pquartosalugados.Cells(nLin, 7).Value = checkout
    Pquartosalugados.Cells(nLin, 8).Value = telasistema.TextTotal.Value
    Pquartosalugados.Cells(nLin, 9).Value = "Alugado"
    
    'Deixar a acomodacao indisponível
    codAcom = telasistema.TextIdAcomodacao.Value
   
    'Total de linhas na planilha
    totalLin = Pacomodacoes.Range("A1").CurrentRegion.Rows.Count
    
    For nLin = 2 To totalLin
        
       If Pacomodacoes.Cells(nLin, 1).Value = codAcom Then
            Pacomodacoes.Cells(nLin, 6).Value = "Indisponível"
            Exit Sub
       End If
    
    Call BTnovoregistro_Click
    Next
    
End Sub
Private Sub CBNumDias_Change()

    Dim numDias As Integer
    Dim dataCheckout As Date
    Dim dataCheckin As Date
    Dim Total As Integer
    Dim Diaria As Integer
    
    'verifica numero valido
    If IsNumeric(CBNumDias.Value) = False Then
        Exit Sub
    Else
        numDias = CBNumDias.Value
    End If
    
    Diaria = TextQtdeDiaria.Value
    
    'verifica se Diaria está preenchido com um numero válido
    If IsNumeric(TextQtdeDiaria.Value) = True Then
        Diaria = TextQtdeDiaria.Value
    Else
        Diaria = 0
    End If
    
    numDias = CBNumDias.Value
    dataCheckin = Date
    dataCheckout = dataCheckin + numDias
    Total = Diaria * numDias
    
    
    'preenche os valores das datas
    TextCheckout.Value = dataCheckout
    TextCheckin.Value = dataCheckin
    
    'preenche o campo Total
    TextTotal.Value = Total
    
End Sub
Private Sub UserForm_Initialize()

    Call BTnovoregistro_Click
    
End Sub
