Attribute VB_Name = "Fornecedor"
Public FornecedorEscolhido As Integer

Sub AddFornecedor()

    Sheets("Pré-questionário fornecedor").Select
    Cells(2, 2).Select
    FornecedorEscolhido = 0
    
End Sub
Sub SalvarFornecedor()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário fornecedor")
    
    Dim Fornecedores As Worksheet
    Set Fornecedores = Sheets("Fornecedores")
    
    Dim Nome_Fornecedor As String
    Nome_Fornecedor = UCase(Questionario.Range("B5"))
    
    '-------------------------------------------------------
    
    ultimo = Fornecedores.Range("A1").End(xlDown).Row - 2
    existe = 0
    
        If (FornecedorEscolhido = 0) Then 'Adicionando Fornecedor
            'Verificar campos obrigatorios
            If (Nome_Fornecedor = "" Or Questionario.Range("B15") = "" Or Questionario.Range("B39") = "") Then
                MsgBox ("Preencha todos os campos obrigatórios antes de salvar!")
            
            Else
                For i = 1 To ultimo
                    If Fornecedores.Cells(i + 2, 2) = Nome_Fornecedor Then
                        MsgBox ("O nome da empresa já existe!")
                        existe = 1
                    End If
                Next i
                
                If (existe = 0) Then
                    nova_linha = Fornecedores.Range("A1").End(xlDown).Row + 1
                    If (nova_linha = 3) Then
                        Fornecedores.Cells(nova_linha, 1) = "F1"
                    Else
                        Fornecedores.Cells(nova_linha, 1) = "F" & Right$(Fornecedores.Range("A1").End(xlDown).Value, Len(Fornecedores.Range("A1").End(xlDown)) - 1) + 1
                    End If
                    GravarCampos (nova_linha)
                    MsgBox "Empresa cadastrada!"
                    LimparFornecedor
                    VoltarFornecedor
                End If
                
            End If
            
        Else 'Editando Fornecedor
            'Verificar campos obrigatorios
            If (Nome_Fornecedor = "" Or Questionario.Range("B15") = "" Or Questionario.Range("B39") = "") Then
                MsgBox ("Preencha todos os campos obrigatórios antes de salvar!")
            Else
                'pode salvar alterações
                GravarCampos (FornecedorEscolhido + 2)
                MsgBox "Alterações salvas!"
                LimparFornecedor
                VoltarFornecedor
            End If
            
        End If

End Sub

Sub LimparFornecedor()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário fornecedor")
    
    Questionario.Range("B5") = ""
    Questionario.Range("B7") = ""
    Questionario.Range("B9") = ""
    Questionario.Range("B11") = ""
    Questionario.Range("B13") = ""
    Questionario.Range("B15") = ""
    Questionario.Range("B18") = ""
    Questionario.Range("B20") = ""
    Questionario.Range("B22") = ""
    Questionario.Range("B25") = ""
    Questionario.Range("B27") = ""
    Questionario.Range("B29") = ""
    Questionario.Range("B35") = ""
    Questionario.Range("B39") = ""
    Questionario.Range("B41") = ""
    

End Sub

Sub EditarFornecedor()

    Dim Fornecedores As Worksheet
    Set Fornecedores = Sheets("Fornecedores")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário fornecedor")
    
    '-------------------------------------------------------
    
    ultima = Fornecedores.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormFornecedor.Show 'Altera FornecedorEscolhido
        If (FornecedorEscolhido <> 0) Then
            Questionario.Range("B5") = Fornecedores.Cells(FornecedorEscolhido + 2, 2)
            Questionario.Range("B7") = Fornecedores.Cells(FornecedorEscolhido + 2, 3)
            Questionario.Range("B9") = Fornecedores.Cells(FornecedorEscolhido + 2, 4)
            Questionario.Range("B11") = Fornecedores.Cells(FornecedorEscolhido + 2, 5)
            Questionario.Range("B13") = Fornecedores.Cells(FornecedorEscolhido + 2, 6)
            Questionario.Range("B15") = Fornecedores.Cells(FornecedorEscolhido + 2, 7)
            Questionario.Range("B18") = Fornecedores.Cells(FornecedorEscolhido + 2, 8)
            Questionario.Range("B20") = Fornecedores.Cells(FornecedorEscolhido + 2, 9)
            Questionario.Range("B22") = Fornecedores.Cells(FornecedorEscolhido + 2, 10)
            Questionario.Range("B25") = Fornecedores.Cells(FornecedorEscolhido + 2, 11)
            Questionario.Range("B27") = Fornecedores.Cells(FornecedorEscolhido + 2, 12)
            Questionario.Range("B29") = Fornecedores.Cells(FornecedorEscolhido + 2, 13)
            Questionario.Range("B35") = Fornecedores.Cells(FornecedorEscolhido + 2, 14)
            Questionario.Range("B39") = Fornecedores.Cells(FornecedorEscolhido + 2, 15)
            Questionario.Range("B41") = Fornecedores.Cells(FornecedorEscolhido + 2, 16)
               
            Questionario.Select
            Cells(2, 2).Select
        End If
    Else
        MsgBox "Não há empresas cadastradas!"
    End If
        
End Sub

'Funçao usada no UserForm. Retorna ID em empresa escolhida
Sub ReceberFornecedor(Empresa As String)

    ultima = Sheets("Fornecedores").Range("A1").End(xlDown).Row - 2
    'Pegar a linha
    FornecedorEscolhido = 0
        For i = 1 To ultima
            If (Sheets("Fornecedores").Cells(2 + i, 2) = Empresa) Then
                FornecedorEscolhido = i
            End If
        Next i
    
End Sub
Sub GravarCampos(Linha As Integer)

    Dim Fornecedores As Worksheet
    Set Fornecedores = Sheets("Fornecedores")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário fornecedor")
    
    '-------------------------------------------------------
    
        
        ' gravar nome maiusulo
        Nome = UCase(Questionario.Range("B5"))
        
        Fornecedores.Cells(Linha, 2) = Nome
        Fornecedores.Cells(Linha, 3) = Questionario.Range("B7")
        Fornecedores.Cells(Linha, 4) = Questionario.Range("B9")
        Fornecedores.Cells(Linha, 5) = Questionario.Range("B11")
        Fornecedores.Cells(Linha, 6) = Questionario.Range("B13")
        Fornecedores.Cells(Linha, 7) = Questionario.Range("B15")
        Fornecedores.Cells(Linha, 8) = Questionario.Range("B18")
        Fornecedores.Cells(Linha, 9) = Questionario.Range("B20")
        Fornecedores.Cells(Linha, 10) = Questionario.Range("B22")
        Fornecedores.Cells(Linha, 11) = Questionario.Range("B25")
        Fornecedores.Cells(Linha, 12) = Questionario.Range("B27")
        Fornecedores.Cells(Linha, 13) = Questionario.Range("B29")
        Fornecedores.Cells(Linha, 14) = Questionario.Range("B35")
        Fornecedores.Cells(Linha, 15) = Questionario.Range("B39")
        Fornecedores.Cells(Linha, 16) = Questionario.Range("B41")
                        
 
End Sub
Sub VisualizarFornecedor()

    Dim Fornecedores As Worksheet
    Set Fornecedores = Sheets("Fornecedores")
    
    Dim Modelo As Worksheet
    Set Modelo = Sheets("Visualização Fornecedor")
    
    '-------------------------------------------------------
    
    ultima = Sheets("Fornecedores").Range("A1").End(xlDown).Row - 2

    If (ultima > 0) Then
        UserFormFornecedor.Show
        If (FornecedorEscolhido <> 0) Then
        
            Modelo.Range("B5") = Fornecedores.Cells(FornecedorEscolhido + 2, 2)
            Modelo.Range("B7") = Fornecedores.Cells(FornecedorEscolhido + 2, 3)
            Modelo.Range("B9") = Fornecedores.Cells(FornecedorEscolhido + 2, 4)
            Modelo.Range("B11") = Fornecedores.Cells(FornecedorEscolhido + 2, 5)
            Modelo.Range("B13") = Fornecedores.Cells(FornecedorEscolhido + 2, 6)
            Modelo.Range("B15") = Fornecedores.Cells(FornecedorEscolhido + 2, 7)
            
      
            Modelo.Range("B17") = Fornecedores.Cells(FornecedorEscolhido + 2, 8)
            If (Fornecedores.Cells(FornecedorEscolhido + 2, 9) <> "") Then
                Modelo.Range("B17") = Modelo.Range("B17") & "," & Fornecedores.Cells(FornecedorEscolhido + 2, 9)
            End If
            If (Fornecedores.Cells(FornecedorEscolhido + 2, 10) <> "") Then
                Modelo.Range("B17") = Modelo.Range("B17") & "," & Fornecedores.Cells(FornecedorEscolhido + 2, 10)
            End If
            
            Modelo.Range("B19") = Fornecedores.Cells(FornecedorEscolhido + 2, 11)
            If (Fornecedores.Cells(FornecedorEscolhido + 2, 12) <> "") Then
                Modelo.Range("B19") = Modelo.Range("B19") & "," & Fornecedores.Cells(FornecedorEscolhido + 2, 12)
            End If
            If (Fornecedores.Cells(FornecedorEscolhido + 2, 13) <> "") Then
                Modelo.Range("B19") = Modelo.Range("B19") & "," & Fornecedores.Cells(FornecedorEscolhido + 2, 13)
            End If
            
            Modelo.Range("B25") = Fornecedores.Cells(FornecedorEscolhido + 2, 14)
            Modelo.Range("B29") = Fornecedores.Cells(FornecedorEscolhido + 2, 15)
            Modelo.Range("B31") = Fornecedores.Cells(FornecedorEscolhido + 2, 16)
            
            Modelo.Select
        
        Cells(2, 2).Select
        End If
    Else
        MsgBox "Não há empresas cadastradas!"
    End If

End Sub

Sub ExcluirFornecedor()

    Dim Fornecedores As Worksheet
    Set Fornecedores = Sheets("Fornecedores")
        
    ultima = Fornecedores.Range("A1").End(xlDown).Row - 2

    If (ultima > 0) Then
        UserFormFornecedor.Show
        
        If (FornecedorEscolhido <> 0) Then
            
            'Apagar o fornecedor do danco de dados (mover a lista para cima)
            Fornecedores.Rows(FornecedorEscolhido + 2).Delete Shift:=xlUp
            
            'Apagar onde tem o fornecedor do banco de dados de Notas
            ultima_linha_notas = Sheets("Notas").Range("A1").End(xlDown).Row
            If ultima_linha_notas > 2 Then
            
                L = 3
                For v = 1 To ultima_linha_notas - 2
                    If (Sheets("Notas").Cells(L, 2) = ID) Then
                        Sheets("Notas").Rows(L).Delete Shift:=xlUp
                    Else
                        L = L + 1
                    End If
                Next v
                
            End If
            
        End If
    Else
        MsgBox "Não há empresas cadastradas!"
    End If
End Sub
Sub VoltarFornecedor()
    Forn.Show
End Sub


