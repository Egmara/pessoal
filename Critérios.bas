Attribute VB_Name = "Critérios"
Public CritérioEscolhido As Integer
Sub AddCriterio()

    CritérioEscolhido = 0
    SubcritérioEscolhido = 0
    Sheets("Novo Critério").Select
    Cells(2, 2).Select
   
End Sub
Sub SalvarCriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Critério")
  
    Dim Nome_Criterio As String
    Nome_Criterio = Questionario.Range("B5")
        
    '-------------------------------------------------------
    'Verificar campos obrigatorios
    If (Questionario.Range("B5") = "" Or Questionario.Range("B10") = "") Then
        MsgBox ("Preencha todos os campos obrigatórios antes de salvar!")
    Else
                
        If (CritérioEscolhido = 0) Then 'Adicionando criterio
            
            nova_linha = Sheets("Critérios").Range("A1").End(xlDown).Row + 1
            'Gravar no banco de dados
            GravarCamposCriterio (nova_linha)
            If (nova_linha = 3) Then
                Sheets("Critérios").Cells(nova_linha, 1) = "C1"
            Else
                Sheets("Critérios").Cells(nova_linha, 1) = "C" & Right$(Sheets("Critérios").Range("A1").End(xlDown).Value, Len(Sheets("Critérios").Range("A1").End(xlDown)) - 1) + 1
            End If
            
            MsgBox "Critério adicionado com sucesso!"
            LimparCriterio
            VoltarCriterio
        
        
        
        Else 'Editando criterio
        
            GravarCamposCriterio (CritérioEscolhido + 2) 'GravarCampos está recebendo a linha do banco de dados referente ao criterio escolhido
            MsgBox "Alterações salvas!"
            LimparCriterio
            VoltarCriterio

        End If
        
    End If
End Sub
Sub EditarCriterio()

    Dim Critérios As Worksheet
    Set Critérios = Sheets("Critérios")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Critério")
    
    '-------------------------------------------------------
    
    ultima = Critérios.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormCriterio.Show vbModeless 'Altera CriterioEscolhido
        
        If (CritérioEscolhido <> 0) Then
        
            'Questionário recebe os dados do banco de dados
            Questionario.Range("B5") = Critérios.Cells(CritérioEscolhido + 2, 2)
            Questionario.Range("B7") = Critérios.Cells(CritérioEscolhido + 2, 3)
            Questionario.Range("B10") = Critérios.Cells(CritérioEscolhido + 2, 4)
            Questionario.Range("B13") = Critérios.Cells(CritérioEscolhido + 2, 5)
            Questionario.Range("B15") = Critérios.Cells(CritérioEscolhido + 2, 6)
            
        
            Questionario.Select
            Cells(2, 2).Select
        End If
        
    Else
        MsgBox "Não há critérios cadastrados!"
    End If
        
End Sub


Sub GravarCamposCriterio(Linha As Integer)

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Critério")
    
    Dim Critérios As Worksheet
    Set Critérios = Sheets("Critérios")
    
    '-------------------------------------------------------
    
   
    Critérios.Cells(Linha, 2) = Questionario.Range("B5")
    Critérios.Cells(Linha, 3) = Questionario.Range("B7")
    Critérios.Cells(Linha, 4) = Questionario.Range("B10")
    Critérios.Cells(Linha, 5) = Questionario.Range("B13")
    Critérios.Cells(Linha, 6) = Questionario.Range("B15")
    
End Sub
Sub LimparCriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Critério")
    
    Questionario.Range("B5") = ""
    Questionario.Range("B7") = ""
    Questionario.Range("B10") = ""
    Questionario.Range("B13") = ""
    Questionario.Range("B15") = ""
    
    'Limpar Subcriterio
    Questionario.Range("B22") = ""
    Questionario.Range("B24") = ""
    
End Sub

'Função usada no UserForm. Retorna ID do criterio escolhida.
Sub ReceberCriterio(criterio As String)

    ultima = Sheets("Critérios").Range("A1").End(xlDown).Row - 2
    'Pegar a linha
    CritérioEscolhido = 0
        For i = 1 To ultima
           If (Sheets("Critérios").Cells(2 + i, 2) = criterio) Then
                CritérioEscolhido = i
            End If
        Next i
    
End Sub

Sub ExcluirSub_função(ID As String)


    
    'busca a linha da ID
    L = 3
    While (Sheets("Subcritérios").Cells(L, 1) <> ID)
        L = L + 1
    Wend
    
    'apaga o subcriterio atraves da linha da ID
    Sheets("Subcritérios").Rows(L).Delete Shift:=xlUp
    
    'Busca a coluna da ID
    C = 2
    While (Sheets("Pesos").Cells(1, C) <> ID)
        C = C + 1
    Wend
    'Exclui da planilha de pesos
    Sheets("Pesos").Columns(C).Delete Shift:=xlToLeft
    
    'Busca a coluna da ID
    C = 2
    While (Sheets("Notas").Cells(1, C) <> ID)
        C = C + 1
    Wend
    'Exclui da planilha de notas
    Sheets("Notas").Columns(C).Delete Shift:=xlToLeft
    
End Sub

Sub ExcluirCriterio()

    Dim Critérios As Worksheet
    Set Critérios = Sheets("Critérios")
        
    Dim Subcritérios As Worksheet
    Set Subcritérios = Sheets("Subcritérios")
    
    ultima = Critérios.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormCriterio.Show vbModeless
        
        If (CritérioEscolhido <> 0) Then
       
            'primeiro excluir os subcritérios dele
            k = 7
            While (Critérios.Cells(CritérioEscolhido + 2, k) <> "")
                ID = Critérios.Cells(CritérioEscolhido + 2, k)
                ExcluirSub_função (ID)
                k = k + 1
            Wend
            
            'depois exclui a linha do critério
            Critérios.Rows(CritérioEscolhido + 2).Delete Shift:=xlUp
            
        End If
    Else
        MsgBox "Não há novos critérios cadastrados!"
    End If
    
 End Sub



Sub VoltarCriterio()

    'apaga subcriterios caso volte ao menu (pois o criterio nao foi salvo)
    nova_linha = Sheets("Critérios").Range("A1").End(xlDown).Row + 1
    ultima_linha_sub = Sheets("Subcritérios").Range("A1").End(xlDown).Row
    i = 7
    While (Sheets("Critérios").Cells(nova_linha, i) <> "")
        Sheets("Critérios").Cells(nova_linha, i) = ""
        Sheets("Subcritérios").Cells(ultima_linha_sub + 7 - i, 1) = ""
        Sheets("Subcritérios").Cells(ultima_linha_sub + 7 - i, 2) = ""
        Sheets("Subcritérios").Cells(ultima_linha_sub + 7 - i, 3) = ""
        i = i + 1
    Wend
        Crit.Show vbModeless
End Sub

Sub VerCriterios()
    
    Dim Critérios As Worksheet
    Set Critérios = Sheets("Critérios")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Critério")
    
    '-------------------------------------------------------
    
    ultima = Critérios.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormCriterio.Show vbModeless 'Altera CriterioEscolhido
        
        If (CritérioEscolhido <> 0) Then
        
            'Questionário recebe os dados do banco de dados
            Questionario.Range("B5") = Critérios.Cells(CritérioEscolhido + 2, 2)
            Questionario.Range("B7") = Critérios.Cells(CritérioEscolhido + 2, 3)
            Questionario.Range("B10") = Critérios.Cells(CritérioEscolhido + 2, 4)
            Questionario.Range("B13") = Critérios.Cells(CritérioEscolhido + 2, 5)
            Questionario.Range("B15") = Critérios.Cells(CritérioEscolhido + 2, 6)
            
        
            Questionario.Select
            Cells(2, 2).Select
        End If
        
    Else
        MsgBox "Não há critérios cadastrados!"
    End If
End Sub

