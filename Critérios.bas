Attribute VB_Name = "Crit�rios"
Public Crit�rioEscolhido As Integer
Sub AddCriterio()

    Crit�rioEscolhido = 0
    Subcrit�rioEscolhido = 0
    Sheets("Novo Crit�rio").Select
    Cells(2, 2).Select
   
End Sub
Sub SalvarCriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Crit�rio")
  
    Dim Nome_Criterio As String
    Nome_Criterio = Questionario.Range("B5")
        
    '-------------------------------------------------------
    'Verificar campos obrigatorios
    If (Questionario.Range("B5") = "" Or Questionario.Range("B10") = "") Then
        MsgBox ("Preencha todos os campos obrigat�rios antes de salvar!")
    Else
                
        If (Crit�rioEscolhido = 0) Then 'Adicionando criterio
            
            nova_linha = Sheets("Crit�rios").Range("A1").End(xlDown).Row + 1
            'Gravar no banco de dados
            GravarCamposCriterio (nova_linha)
            If (nova_linha = 3) Then
                Sheets("Crit�rios").Cells(nova_linha, 1) = "C1"
            Else
                Sheets("Crit�rios").Cells(nova_linha, 1) = "C" & Right$(Sheets("Crit�rios").Range("A1").End(xlDown).Value, Len(Sheets("Crit�rios").Range("A1").End(xlDown)) - 1) + 1
            End If
            
            MsgBox "Crit�rio adicionado com sucesso!"
            LimparCriterio
            VoltarCriterio
        
        
        
        Else 'Editando criterio
        
            GravarCamposCriterio (Crit�rioEscolhido + 2) 'GravarCampos est� recebendo a linha do banco de dados referente ao criterio escolhido
            MsgBox "Altera��es salvas!"
            LimparCriterio
            VoltarCriterio

        End If
        
    End If
End Sub
Sub EditarCriterio()

    Dim Crit�rios As Worksheet
    Set Crit�rios = Sheets("Crit�rios")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Crit�rio")
    
    '-------------------------------------------------------
    
    ultima = Crit�rios.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormCriterio.Show vbModeless 'Altera CriterioEscolhido
        
        If (Crit�rioEscolhido <> 0) Then
        
            'Question�rio recebe os dados do banco de dados
            Questionario.Range("B5") = Crit�rios.Cells(Crit�rioEscolhido + 2, 2)
            Questionario.Range("B7") = Crit�rios.Cells(Crit�rioEscolhido + 2, 3)
            Questionario.Range("B10") = Crit�rios.Cells(Crit�rioEscolhido + 2, 4)
            Questionario.Range("B13") = Crit�rios.Cells(Crit�rioEscolhido + 2, 5)
            Questionario.Range("B15") = Crit�rios.Cells(Crit�rioEscolhido + 2, 6)
            
        
            Questionario.Select
            Cells(2, 2).Select
        End If
        
    Else
        MsgBox "N�o h� crit�rios cadastrados!"
    End If
        
End Sub


Sub GravarCamposCriterio(Linha As Integer)

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Crit�rio")
    
    Dim Crit�rios As Worksheet
    Set Crit�rios = Sheets("Crit�rios")
    
    '-------------------------------------------------------
    
   
    Crit�rios.Cells(Linha, 2) = Questionario.Range("B5")
    Crit�rios.Cells(Linha, 3) = Questionario.Range("B7")
    Crit�rios.Cells(Linha, 4) = Questionario.Range("B10")
    Crit�rios.Cells(Linha, 5) = Questionario.Range("B13")
    Crit�rios.Cells(Linha, 6) = Questionario.Range("B15")
    
End Sub
Sub LimparCriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Crit�rio")
    
    Questionario.Range("B5") = ""
    Questionario.Range("B7") = ""
    Questionario.Range("B10") = ""
    Questionario.Range("B13") = ""
    Questionario.Range("B15") = ""
    
    'Limpar Subcriterio
    Questionario.Range("B22") = ""
    Questionario.Range("B24") = ""
    
End Sub

'Fun��o usada no UserForm. Retorna ID do criterio escolhida.
Sub ReceberCriterio(criterio As String)

    ultima = Sheets("Crit�rios").Range("A1").End(xlDown).Row - 2
    'Pegar a linha
    Crit�rioEscolhido = 0
        For i = 1 To ultima
           If (Sheets("Crit�rios").Cells(2 + i, 2) = criterio) Then
                Crit�rioEscolhido = i
            End If
        Next i
    
End Sub

Sub ExcluirSub_fun��o(ID As String)


    
    'busca a linha da ID
    L = 3
    While (Sheets("Subcrit�rios").Cells(L, 1) <> ID)
        L = L + 1
    Wend
    
    'apaga o subcriterio atraves da linha da ID
    Sheets("Subcrit�rios").Rows(L).Delete Shift:=xlUp
    
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

    Dim Crit�rios As Worksheet
    Set Crit�rios = Sheets("Crit�rios")
        
    Dim Subcrit�rios As Worksheet
    Set Subcrit�rios = Sheets("Subcrit�rios")
    
    ultima = Crit�rios.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormCriterio.Show vbModeless
        
        If (Crit�rioEscolhido <> 0) Then
       
            'primeiro excluir os subcrit�rios dele
            k = 7
            While (Crit�rios.Cells(Crit�rioEscolhido + 2, k) <> "")
                ID = Crit�rios.Cells(Crit�rioEscolhido + 2, k)
                ExcluirSub_fun��o (ID)
                k = k + 1
            Wend
            
            'depois exclui a linha do crit�rio
            Crit�rios.Rows(Crit�rioEscolhido + 2).Delete Shift:=xlUp
            
        End If
    Else
        MsgBox "N�o h� novos crit�rios cadastrados!"
    End If
    
 End Sub



Sub VoltarCriterio()

    'apaga subcriterios caso volte ao menu (pois o criterio nao foi salvo)
    nova_linha = Sheets("Crit�rios").Range("A1").End(xlDown).Row + 1
    ultima_linha_sub = Sheets("Subcrit�rios").Range("A1").End(xlDown).Row
    i = 7
    While (Sheets("Crit�rios").Cells(nova_linha, i) <> "")
        Sheets("Crit�rios").Cells(nova_linha, i) = ""
        Sheets("Subcrit�rios").Cells(ultima_linha_sub + 7 - i, 1) = ""
        Sheets("Subcrit�rios").Cells(ultima_linha_sub + 7 - i, 2) = ""
        Sheets("Subcrit�rios").Cells(ultima_linha_sub + 7 - i, 3) = ""
        i = i + 1
    Wend
        Crit.Show vbModeless
End Sub

Sub VerCriterios()
    
    Dim Crit�rios As Worksheet
    Set Crit�rios = Sheets("Crit�rios")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Crit�rio")
    
    '-------------------------------------------------------
    
    ultima = Crit�rios.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormCriterio.Show vbModeless 'Altera CriterioEscolhido
        
        If (Crit�rioEscolhido <> 0) Then
        
            'Question�rio recebe os dados do banco de dados
            Questionario.Range("B5") = Crit�rios.Cells(Crit�rioEscolhido + 2, 2)
            Questionario.Range("B7") = Crit�rios.Cells(Crit�rioEscolhido + 2, 3)
            Questionario.Range("B10") = Crit�rios.Cells(Crit�rioEscolhido + 2, 4)
            Questionario.Range("B13") = Crit�rios.Cells(Crit�rioEscolhido + 2, 5)
            Questionario.Range("B15") = Crit�rios.Cells(Crit�rioEscolhido + 2, 6)
            
        
            Questionario.Select
            Cells(2, 2).Select
        End If
        
    Else
        MsgBox "N�o h� crit�rios cadastrados!"
    End If
End Sub

