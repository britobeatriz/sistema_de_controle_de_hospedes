VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} consultarAlugados 
   Caption         =   "Consultar Quartos Alugados"
   ClientHeight    =   4392
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8628.001
   OleObjectBlob   =   "consultarAlugados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "consultarAlugados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub preencherListBox()

    Dim base As Range
    
    Set base = Pfiltroquartosalugados.Range("A5").CurrentRegion.Offset(1)
    
    ListBox1.RowSource = base.Address(, , , True)
    
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim i As Integer
    Dim checkin As Date
    Dim checkout As Date
    
    checkin = ListBox1.List(i, 5)
    checkout = ListBox1.List(i, 6)
    
    i = ListBox1.ListIndex
    
    telasistema.TextId.Value = ListBox1.List(i, 0)
    telasistema.TextCliente.Value = ListBox1.List(i, 1)
    telasistema.TextContato.Value = ListBox1.List(i, 2)
    telasistema.TextIdAcomodacao.Value = ListBox1.List(i, 3)
    telasistema.TextCheckin.Value = checkin
    telasistema.TextCheckout.Value = checkout
    telasistema.TextTotal.Value = ListBox1.List(i, 7)
    
    Unload Me 'fechar o formulario
    
End Sub
Private Sub UserForm_Initialize()

    Call preencherListBox
    Call FiltrarAlugados
    
End Sub
Sub preencherCriterios()
    'preenche criterios na planilha filtro disp
    
    Pfiltroquartosalugados.Range("B2").Value = TextNomeCliente.Value
    
    Call FiltrarAlugados
    Call preencherListBox

End Sub
Private Sub TextNomeCliente_Change()

    Call preencherCriterios
    
End Sub
