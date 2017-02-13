Attribute VB_Name = "Subcrit�rios"
'Vari�vel global
Public Subcrit�rioEscolhido As Integer

Sub SalvarSubCriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Crit�rio")
    
    Dim Crit�rios As Worksheet
    Set Crit�rios = Sheets("Crit�rios")
    
    Dim Subcrit�rios As Worksheet
    Set Subcrit�rios = Sheets("Subcrit�rios")
  
    Dim Nome_SubCriterio As String
    Nome_SubCriterio = Questionario.Range("B22")
    
    Dim Nome_Criterio As String
    Nome_Criterio = Questionario.Range("B5")
    
    '-------------------------------------------------------
    If (Nome_SubCriterio = "") Then
    MsgBox ("Digite o nome do subcrit�rio!")
    Else
        If (Nome_Criterio = "") Then
                MsgBox ("Digite o nome do crit�rio referente!")
        Else
            'Descobri a linha do criterio
            If (Crit�rioEscolhido = 0) Then 'Adicionandom novo criteiro
                LinhaCrit = Crit�rios.Range("A1").End(xlDown).Row + 1
                If (Crit�rios.Range("A1").End(xlDown).Row = 3) Then
                    IDCrit = "C1"
                Else
                    IDCrit = "C" & Right$(Sheets("Crit�rios").Range("A1").End(xlDown).Value, Len(Sheets("Crit�rios").Range("A1").End(xlDown)) - 1) + 1
                End If
                    
            Else 'Editando criterio
                LinhaCrit = Crit�rioEscolhido + 2
                IDCrit = Sheets("Crit�rios").Cells(LinhaCrit, 1)
            End If
             
                
       
        
    
            If (Subcrit�rioEscolhido = 0) Then 'Adicionando subcriterio
                'Descobri a coluna do novo subcriteiro
                Colunasub = 7
                While (Crit�rios.Cells(LinhaCrit, Colunasub) <> "")
                    Colunasub = Colunasub + 1
                Wend
            
                'Descobri a linha do novo subcriteiro a ser gravada
                Linhasub = Subcrit�rios.Range("A1").End(xlDown).Row + 1
            
                'Gravar no banco de dados
                '*-----------------------------------------------------------------------------------------------------------------

                
                If (Colunasub = 7) Then 'primeiro subcriteiro
                    Sheets("Crit�rios").Cells(LinhaCrit, Colunasub) = IDCrit & "S1"
                Else
                    'acha a posi��o do S no ID anterior
                    IDAnt = Sheets("Crit�rios").Cells(LinhaCrit, Colunasub - 1)
                    For i = 1 To Len(IDAnt)
                    IndiceS = Len(IDAnt)
                    If (Mid(IDAnt, i, 1) = s) Then
                        IndiceS = i
                        End If
                    Next i
                    ID = IDCrit & "S" & Right$(IDAnt, Len(IDAnt) - IndiceS + 1) + 1
    
                    Crit�rios.Cells(LinhaCrit, Colunasub) = ID
                    Subcrit�rios.Cells(Linhasub, 1) = ID
    
                    Subcrit�rios.Cells(Linhasub, 2) = Questionario.Range("B22")
                    Subcrit�rios.Cells(Linhasub, 3) = Questionario.Range("B24")
                End If
                
                '*-----------------------------------------------------------------------------------------------------------------
                MsgBox "Subcrit�rio adicionado com sucesso!"
                LimparSubcriterio
                    
            Else 'Editando subcriterio
                ID = Crit�rios.Cells(LinhaCrit, Subcrit�rioEscolhido + 6)
            
                'Coluna do subcriteiro na plan Criterios
                Colunasub = Subcrit�rioEscolhido + 6
                'Descobri a linha do subcriteiro a ser gravada
                Linhasub = 3
                While (Subcrit�rios.Cells(Linhasub, 1) <> ID)
                    Linhasub = Linhasub + 1
                Wend
            
                'Gravar no banco de dados
                Subcrit�rios.Cells(Linhasub, 2) = Questionario.Range("B22")
                Subcrit�rios.Cells(Linhasub, 3) = Questionario.Range("B24")

                MsgBox "Subcrit�rio editado com sucesso!"
                Subcrit�rioEscolhido = 0
                LimparSubcriterio
        
            End If
        End If
    End If
    
End Sub

Sub LimparSubcriterio()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Novo Crit�rio")
    
    Questionario.Range("B22") = ""
    Questionario.Range("B24") = ""
    
End Sub
'Fun��o usada no UserForm.
Sub ReceberSubcriterio(subcriterio As String)

    ultima = Sheets("Subcrit�rios").Range("A1").End(xlDown).Row - 2
    
    'Pegar a linha
    Subcrit�rioEscolhido = 0
        For i = 1 To ultima
           If (Sheets("Subcrit�rios").Cells(2 + i, 2) = subcriterio) Then
                Subcrit�rioEscolhido = i
            End If
        Next i
    
End Sub

Sub ExcluirSubcriterio()
    
    Dim Subcrit�rios As Worksheet
    Set Subcrit�rios = Sheets("Subcrit�rios")
    
    Dim Crit�rios As Worksheet
    Set Crit�rios = Sheets("Crit�rios")

    ultimo = Crit�rios.Range("A1").End(xlDown).Row - 2
    
    If (ultimo > 0) Then
        UserFormCriterio.Show
        
        If Crit�rioEscolhido <> 0 Then
            
            ultimo_sub = 0
            C = 7
            While (Crit�rios.Cells(Crit�rioEscolhido + 2, C) <> "")
                ultimo_sub = ultimo_sub + 1
                C = C + 1
            Wend

            If (ultimo_sub > 0) Then 'se o criterio tem subcriterio
                UserFormSubcriterio.Show
                
                If (Subcrit�rioEscolhido <> 0) Then
                    
                    ID_sub = Subcrit�rios.Cells(Subcrit�rioEscolhido + 2, 1)
                    
                    'exclui a linha do Subcrit�rio
                    Subcrit�rios.Rows(Subcrit�rioEscolhido + 2).Delete Shift:=xlUp
                    
                    'exclui a ID_sub da planilha Crit�rios
                    C = 7
                    'procura a coluna da ID_sub
                    While (Crit�rios.Cells(Crit�rioEscolhido + 2, C) <> ID_sub)
                        C = C + 1
                    Wend
                    'e desloca todos pra esquerda
                    Crit�rios.Cells(Crit�rioEscolhido + 2, C).Delete Shift:=xlToLeft
                    
                End If
                
            Else
                MsgBox "N�o h� subcrit�rios cadastrados para este crit�rio!"
            End If
            
        End If
       
    Else
        MsgBox "N�o h� crit�rios cadastrados!"
    End If
 End Sub

