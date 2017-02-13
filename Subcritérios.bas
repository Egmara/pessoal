Attribute VB_Name = "Subcritérios"
'Variável global
Public SubcritérioEscolhido As Integer

Sub SalvarSubCriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Critério")
    
    Dim Critérios As Worksheet
    Set Critérios = Sheets("Critérios")
    
    Dim Subcritérios As Worksheet
    Set Subcritérios = Sheets("Subcritérios")
  
    Dim Nome_SubCriterio As String
    Nome_SubCriterio = Questionario.Range("B22")
    
    Dim Nome_Criterio As String
    Nome_Criterio = Questionario.Range("B5")
    
    '-------------------------------------------------------
    If (Nome_SubCriterio = "") Then
    MsgBox ("Digite o nome do subcritério!")
    Else
        If (Nome_Criterio = "") Then
                MsgBox ("Digite o nome do critério referente!")
        Else
            'Descobri a linha do criterio
            If (CritérioEscolhido = 0) Then 'Adicionandom novo criteiro
                LinhaCrit = Critérios.Range("A1").End(xlDown).Row + 1
                If (Critérios.Range("A1").End(xlDown).Row = 3) Then
                    IDCrit = "C1"
                Else
                    IDCrit = "C" & Right$(Sheets("Critérios").Range("A1").End(xlDown).Value, Len(Sheets("Critérios").Range("A1").End(xlDown)) - 1) + 1
                End If
                    
            Else 'Editando criterio
                LinhaCrit = CritérioEscolhido + 2
                IDCrit = Sheets("Critérios").Cells(LinhaCrit, 1)
            End If
             
                
       
        
    
            If (SubcritérioEscolhido = 0) Then 'Adicionando subcriterio
                'Descobri a coluna do novo subcriteiro
                Colunasub = 7
                While (Critérios.Cells(LinhaCrit, Colunasub) <> "")
                    Colunasub = Colunasub + 1
                Wend
            
                'Descobri a linha do novo subcriteiro a ser gravada
                Linhasub = Subcritérios.Range("A1").End(xlDown).Row + 1
            
                'Gravar no banco de dados
                '*-----------------------------------------------------------------------------------------------------------------

                
                If (Colunasub = 7) Then 'primeiro subcriteiro
                    Sheets("Critérios").Cells(LinhaCrit, Colunasub) = IDCrit & "S1"
                Else
                    'acha a posição do S no ID anterior
                    IDAnt = Sheets("Critérios").Cells(LinhaCrit, Colunasub - 1)
                    For i = 1 To Len(IDAnt)
                    IndiceS = Len(IDAnt)
                    If (Mid(IDAnt, i, 1) = s) Then
                        IndiceS = i
                        End If
                    Next i
                    ID = IDCrit & "S" & Right$(IDAnt, Len(IDAnt) - IndiceS + 1) + 1
    
                    Critérios.Cells(LinhaCrit, Colunasub) = ID
                    Subcritérios.Cells(Linhasub, 1) = ID
    
                    Subcritérios.Cells(Linhasub, 2) = Questionario.Range("B22")
                    Subcritérios.Cells(Linhasub, 3) = Questionario.Range("B24")
                End If
                
                '*-----------------------------------------------------------------------------------------------------------------
                MsgBox "Subcritério adicionado com sucesso!"
                LimparSubcriterio
                    
            Else 'Editando subcriterio
                ID = Critérios.Cells(LinhaCrit, SubcritérioEscolhido + 6)
            
                'Coluna do subcriteiro na plan Criterios
                Colunasub = SubcritérioEscolhido + 6
                'Descobri a linha do subcriteiro a ser gravada
                Linhasub = 3
                While (Subcritérios.Cells(Linhasub, 1) <> ID)
                    Linhasub = Linhasub + 1
                Wend
            
                'Gravar no banco de dados
                Subcritérios.Cells(Linhasub, 2) = Questionario.Range("B22")
                Subcritérios.Cells(Linhasub, 3) = Questionario.Range("B24")

                MsgBox "Subcritério editado com sucesso!"
                SubcritérioEscolhido = 0
                LimparSubcriterio
        
            End If
        End If
    End If
    
End Sub

Sub LimparSubcriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Critério")
    
    Questionario.Range("B22") = ""
    Questionario.Range("B24") = ""
    
End Sub
'Função usada no UserForm.
Sub ReceberSubcriterio(subcriterio As String)

    ultima = Sheets("Subcritérios").Range("A1").End(xlDown).Row - 2
    
    'Pegar a linha
    SubcritérioEscolhido = 0
        For i = 1 To ultima
           If (Sheets("Subcritérios").Cells(2 + i, 2) = subcriterio) Then
                SubcritérioEscolhido = i
            End If
        Next i
    
End Sub

Sub ExcluirSubcriterio()
    
    Dim Subcritérios As Worksheet
    Set Subcritérios = Sheets("Subcritérios")
    
    Dim Critérios As Worksheet
    Set Critérios = Sheets("Critérios")

    ultimo = Critérios.Range("A1").End(xlDown).Row - 2
    
    If (ultimo > 0) Then
        UserFormCriterio.Show
        
        If CritérioEscolhido <> 0 Then
            
            ultimo_sub = 0
            C = 7
            While (Critérios.Cells(CritérioEscolhido + 2, C) <> "")
                ultimo_sub = ultimo_sub + 1
                C = C + 1
            Wend

            If (ultimo_sub > 0) Then 'se o criterio tem subcriterio
                UserFormSubcriterio.Show
                
                If (SubcritérioEscolhido <> 0) Then
                    
                    ID_sub = Subcritérios.Cells(SubcritérioEscolhido + 2, 1)
                    
                    'exclui a linha do Subcritério
                    Subcritérios.Rows(SubcritérioEscolhido + 2).Delete Shift:=xlUp
                    
                    'exclui a ID_sub da planilha Critérios
                    C = 7
                    'procura a coluna da ID_sub
                    While (Critérios.Cells(CritérioEscolhido + 2, C) <> ID_sub)
                        C = C + 1
                    Wend
                    'e desloca todos pra esquerda
                    Critérios.Cells(CritérioEscolhido + 2, C).Delete Shift:=xlToLeft
                    
                End If
                
            Else
                MsgBox "Não há subcritérios cadastrados para este critério!"
            End If
            
        End If
       
    Else
        MsgBox "Não há critérios cadastrados!"
    End If
 End Sub

