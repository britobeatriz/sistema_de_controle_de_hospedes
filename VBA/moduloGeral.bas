Attribute VB_Name = "moduloGeral"
Option Explicit
'abre o sistema ao clicar no botao "abrir sistema"
Sub abrir()
    telasistema.Show
End Sub
Sub FiltrarDisponiveis()
    
    Dim acomodacoes As Range
    Dim intCriterio As Range
    Dim intDestino As Range
    
    Set acomodacoes = Pacomodacoes.Range("A1").CurrentRegion
    Set intCriterio = Pfiltrodisp.Range("A1:F2")
    Set intDestino = Pfiltrodisp.Range("A5:F5")
    
    acomodacoes.AdvancedFilter xlFilterCopy, intCriterio, intDestino
    
End Sub
Sub FiltrarAlugados()
    
    Dim acomodacoes As Range
    Dim intCriterio As Range
    Dim intDestino As Range
    
    Set acomodacoes = Pquartosalugados.Range("A1").CurrentRegion
    Set intCriterio = Pfiltroquartosalugados.Range("A1:I2")
    Set intDestino = Pfiltroquartosalugados.Range("A5:I5")
    
    acomodacoes.AdvancedFilter xlFilterCopy, intCriterio, intDestino
    
End Sub

