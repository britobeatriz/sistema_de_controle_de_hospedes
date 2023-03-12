VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} acomodacoesDisp 
   Caption         =   "Quartos Disponíveis"
   ClientHeight    =   4392
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8628.001
   OleObjectBlob   =   "acomodacoesDisp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "acomodacoesDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub preencherListBox()

    Dim listaDisp As Range
    
    Set listaDisp = Pfiltrodisp.Range("A5").CurrentRegion.Offset(1)
    
    ListBox1.RowSource = listaDisp.Address(, , , True)

End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim i As Integer
    
    i = ListBox1.ListIndex
    
    telasistema.TextIdAcomodacao.Value = ListBox1.List(i, 0)
    telasistema.TextQtdeCama.Value = ListBox1.List(i, 1)
    telasistema.TextQtdeQuartos.Value = ListBox1.List(i, 2)
    telasistema.TextQtdeBanheiros.Value = ListBox1.List(i, 3)
    telasistema.TextQtdeDiaria.Value = ListBox1.List(i, 4)
    
    Unload Me 'fechar o formulario
    
End Sub
Private Sub TextQtdeCamas_Change()

    Call preencherCriterios
    
End Sub
Private Sub TextQtdeBanheiros_Change()

    Call preencherCriterios
    
End Sub
Private Sub TextQtdeQuartos_Change()

    Call preencherCriterios
    
End Sub
Sub preencherCriterios()

    'preenche criterios na planilha filtro disp
    
    Pfiltrodisp.Range("B2").Value = TextQtdeCamas.Value
    Pfiltrodisp.Range("C2").Value = TextQtdeQuartos.Value
    Pfiltrodisp.Range("D2").Value = TextQtdeBanheiros.Value
    
    Call FiltrarDisponiveis
    Call preencherListBox

End Sub
Private Sub UserForm_Initialize()

    Call FiltrarDisponiveis
    Call preencherListBox
    
End Sub
