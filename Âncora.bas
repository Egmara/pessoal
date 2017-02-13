Attribute VB_Name = "Âncora"
Public EmpresaEscolhida As Integer

Sub AddAncora()

    Sheets("Pré-questionário âncora").Select
    Cells(2, 2).Select
    EmpresaEscolhida = 0
    
End Sub
Sub SalvarAncora()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário âncora")
    
    Dim Âncoras As Worksheet
    Set Âncoras = Sheets("Âncoras")
    
    Dim Nome_Âncora As String
    Nome_Âncora = UCase(Questionario.Range("B5"))
    
    '-------------------------------------------------------
    
    ultima = Âncoras.Range("A1").End(xlDown).Row - 2
    existe = 0
    
        If (EmpresaEscolhida = 0) Then 'Adicionando ancora
        
            'Verificar campos obrigatorios
            If (Nome_Âncora = "" Or Questionario.Range("B15") = "" Or Questionario.Range("B81") = "") Then
                MsgBox ("Preencha todos os campos obrigatórios antes de salvar!")
            
            Else
                For i = 1 To ultima
                    If Âncoras.Cells(i + 2, 2) = Nome_Âncora Then
                        MsgBox ("O nome da empresa já existe!")
                        existe = 1
                    End If
                Next i
                
                If (existe = 0) Then
                    nova_linha = Âncoras.Range("A1").End(xlDown).Row + 1
                    If (nova_linha = 3) Then
                        Âncoras.Cells(nova_linha, 1) = "A1"
                    Else
                        Âncoras.Cells(nova_linha, 1) = "A" & Right$(Âncoras.Range("A1").End(xlDown).Value, Len(Âncoras.Range("A1").End(xlDown)) - 1) + 1
                    End If
                    GravarCampos (nova_linha)
                    MsgBox "Empresa cadastrada!"
                    
                    LimparAncora
                    VoltarAncora
                End If
                
            End If
            
        Else 'Editando ancora
            'Verificar campos obrigatorios
            If (Nome_Âncora = "" Or Questionario.Range("B15") = "" Or Questionario.Range("B81") = "") Then
                MsgBox ("Preencha todos os campos obrigatórios antes de salvar!")
            Else
                'pode salvar alterações
                GravarCampos (EmpresaEscolhida + 2) 'GravarCampos está recebendo a linha do banco de dados referente à empresa escolhida
                MsgBox "Alterações salvas!"
                LimparAncora
                VoltarAncora
            End If
            
        End If

End Sub

Sub LimparAncora()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário âncora")
    
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
    Questionario.Range("B43") = ""
    Questionario.Range("B45") = ""
    Questionario.Range("B47") = ""
    Questionario.Range("B55") = ""
    Questionario.Range("B57") = ""
    Questionario.Range("B59") = ""
    Questionario.Range("B61") = ""
    Questionario.Range("B63") = ""
    Questionario.Range("B65") = ""
    Questionario.Range("B67") = ""
    Questionario.Range("B69") = ""
    Questionario.Range("B71") = ""
    Questionario.Range("B73") = ""
    Questionario.Range("B75") = ""
    Questionario.Range("B77") = ""
    Questionario.Range("B81") = ""
    Questionario.Range("B83") = ""

End Sub

Sub EditarAncora()

    Dim Âncoras As Worksheet
    Set Âncoras = Sheets("Âncoras")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário âncora")
    
    '-------------------------------------------------------
    
    ultima = Âncoras.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormAncora.Show 'Altera EmpresaEscolhida
        If (EmpresaEscolhida <> 0) Then
            'Questionário recebe os dados do banco de dados
            Questionario.Range("B5") = Âncoras.Cells(EmpresaEscolhida + 2, 2)
            Questionario.Range("B7") = Âncoras.Cells(EmpresaEscolhida + 2, 3)
            Questionario.Range("B9") = Âncoras.Cells(EmpresaEscolhida + 2, 4)
            Questionario.Range("B11") = Âncoras.Cells(EmpresaEscolhida + 2, 5)
            Questionario.Range("B13") = Âncoras.Cells(EmpresaEscolhida + 2, 6)
            Questionario.Range("B15") = Âncoras.Cells(EmpresaEscolhida + 2, 7)
            Questionario.Range("B18") = Âncoras.Cells(EmpresaEscolhida + 2, 8)
            Questionario.Range("B20") = Âncoras.Cells(EmpresaEscolhida + 2, 9)
            Questionario.Range("B22") = Âncoras.Cells(EmpresaEscolhida + 2, 10)
            Questionario.Range("B25") = Âncoras.Cells(EmpresaEscolhida + 2, 11)
            Questionario.Range("B27") = Âncoras.Cells(EmpresaEscolhida + 2, 12)
            Questionario.Range("B29") = Âncoras.Cells(EmpresaEscolhida + 2, 13)
            Questionario.Range("B35") = Âncoras.Cells(EmpresaEscolhida + 2, 14)
            Questionario.Range("B39") = Âncoras.Cells(EmpresaEscolhida + 2, 15)
            Questionario.Range("B41") = Âncoras.Cells(EmpresaEscolhida + 2, 16)
            Questionario.Range("B43") = Âncoras.Cells(EmpresaEscolhida + 2, 17)
            Questionario.Range("B45") = Âncoras.Cells(EmpresaEscolhida + 2, 18)
            Questionario.Range("B47") = Âncoras.Cells(EmpresaEscolhida + 2, 19)
            Questionario.Range("B55") = Âncoras.Cells(EmpresaEscolhida + 2, 20)
            Questionario.Range("B57") = Âncoras.Cells(EmpresaEscolhida + 2, 21)
            Questionario.Range("B59") = Âncoras.Cells(EmpresaEscolhida + 2, 22)
            Questionario.Range("B61") = Âncoras.Cells(EmpresaEscolhida + 2, 23)
            Questionario.Range("B63") = Âncoras.Cells(EmpresaEscolhida + 2, 24)
            Questionario.Range("B65") = Âncoras.Cells(EmpresaEscolhida + 2, 25)
            Questionario.Range("B67") = Âncoras.Cells(EmpresaEscolhida + 2, 26)
            Questionario.Range("B69") = Âncoras.Cells(EmpresaEscolhida + 2, 27)
            Questionario.Range("B71") = Âncoras.Cells(EmpresaEscolhida + 2, 28)
            Questionario.Range("B73") = Âncoras.Cells(EmpresaEscolhida + 2, 29)
            Questionario.Range("B75") = Âncoras.Cells(EmpresaEscolhida + 2, 30)
            Questionario.Range("B77") = Âncoras.Cells(EmpresaEscolhida + 2, 31)
            Questionario.Range("B81") = Âncoras.Cells(EmpresaEscolhida + 2, 32)
            Questionario.Range("B83") = Âncoras.Cells(EmpresaEscolhida + 2, 33)
        
            Questionario.Select
            Cells(2, 2).Select
        End If
    Else
        MsgBox "Não há empresas cadastradas!"
    End If
        
End Sub

'Função usada no UserForm.'Retorna ID em empresa escolhida
Sub ReceberAncora(Empresa As String)

    ultima = Sheets("Âncoras").Range("A1").End(xlDown).Row - 2
    'Pegar a linha
    EmpresaEscolhida = 0
        For i = 1 To ultima
            If (Sheets("Âncoras").Cells(2 + i, 2) = Empresa) Then
                EmpresaEscolhida = i
            End If
        Next i
         
End Sub
Sub GravarCampos(Linha As Integer)

    Dim Âncoras As Worksheet
    Set Âncoras = Sheets("Âncoras")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pré-questionário âncora")
    
    '-------------------------------------------------------
    
        'Banco de dados recebe os dados do Questionário
        
        
        'gravar maiuscula
        Nome = UCase(Questionario.Range("B5"))
        Âncoras.Cells(Linha, 2) = Nome
        Âncoras.Cells(Linha, 3) = Questionario.Range("B7")
        Âncoras.Cells(Linha, 4) = Questionario.Range("B9")
        Âncoras.Cells(Linha, 5) = Questionario.Range("B11")
        Âncoras.Cells(Linha, 6) = Questionario.Range("B13")
        Âncoras.Cells(Linha, 7) = Questionario.Range("B15")
        Âncoras.Cells(Linha, 8) = Questionario.Range("B18")
        Âncoras.Cells(Linha, 9) = Questionario.Range("B20")
        Âncoras.Cells(Linha, 10) = Questionario.Range("B22")
        Âncoras.Cells(Linha, 11) = Questionario.Range("B25")
        Âncoras.Cells(Linha, 12) = Questionario.Range("B27")
        Âncoras.Cells(Linha, 13) = Questionario.Range("B29")
        Âncoras.Cells(Linha, 14) = Questionario.Range("B35")
        Âncoras.Cells(Linha, 15) = Questionario.Range("B39")
        Âncoras.Cells(Linha, 16) = Questionario.Range("B41")
        Âncoras.Cells(Linha, 17) = Questionario.Range("B43")
        Âncoras.Cells(Linha, 18) = Questionario.Range("B45")
        Âncoras.Cells(Linha, 19) = Questionario.Range("B47")
        Âncoras.Cells(Linha, 20) = Questionario.Range("B55")
        Âncoras.Cells(Linha, 21) = Questionario.Range("B57")
        Âncoras.Cells(Linha, 22) = Questionario.Range("B59")
        Âncoras.Cells(Linha, 23) = Questionario.Range("B61")
        Âncoras.Cells(Linha, 24) = Questionario.Range("B63")
        Âncoras.Cells(Linha, 25) = Questionario.Range("B65")
        Âncoras.Cells(Linha, 26) = Questionario.Range("B67")
        Âncoras.Cells(Linha, 27) = Questionario.Range("B69")
        Âncoras.Cells(Linha, 28) = Questionario.Range("B71")
        Âncoras.Cells(Linha, 29) = Questionario.Range("B73")
        Âncoras.Cells(Linha, 30) = Questionario.Range("B75")
        Âncoras.Cells(Linha, 31) = Questionario.Range("B77")
        Âncoras.Cells(Linha, 32) = Questionario.Range("B81")
        Âncoras.Cells(Linha, 33) = Questionario.Range("B83")
        
 
End Sub
Sub VisualizarAncora()

    Dim Âncoras As Worksheet
    Set Âncoras = Sheets("Âncoras")
    
    Dim Modelo As Worksheet
    Set Modelo = Sheets("Visualização Âncora")
    
    '-------------------------------------------------------
    
    ultima = Sheets("Âncoras").Range("A1").End(xlDown).Row - 2

    If (ultima > 0) Then
        UserFormAncora.Show
        
        If (EmpresaEscolhida <> 0) Then
            'Mostrar os dados do banco de dados para o usuário
            
            Modelo.Range("B5") = Âncoras.Cells(EmpresaEscolhida + 2, 2)
            Modelo.Range("B7") = Âncoras.Cells(EmpresaEscolhida + 2, 3)
            Modelo.Range("B9") = Âncoras.Cells(EmpresaEscolhida + 2, 4)
            Modelo.Range("B11") = Âncoras.Cells(EmpresaEscolhida + 2, 5)
            Modelo.Range("B13") = Âncoras.Cells(EmpresaEscolhida + 2, 6)
            Modelo.Range("B15") = Âncoras.Cells(EmpresaEscolhida + 2, 7)
            
            Modelo.Range("B17") = Âncoras.Cells(EmpresaEscolhida + 2, 8)
            If (Âncoras.Cells(EmpresaEscolhida + 2, 9) <> "") Then
                Modelo.Range("B17") = Modelo.Range("B17") & "," & Âncoras.Cells(EmpresaEscolhida + 2, 9)
            End If
            If (Âncoras.Cells(EmpresaEscolhida + 2, 10) <> "") Then
                Modelo.Range("B17") = Modelo.Range("B17") & "," & Âncoras.Cells(EmpresaEscolhida + 2, 10)
            End If
            
            Modelo.Range("B19") = Âncoras.Cells(EmpresaEscolhida + 2, 11)
            If (Âncoras.Cells(EmpresaEscolhida + 2, 12) <> "") Then
                Modelo.Range("B19") = Modelo.Range("B19") & "," & Âncoras.Cells(EmpresaEscolhida + 2, 12)
            End If
            If (Âncoras.Cells(EmpresaEscolhida + 2, 13) <> "") Then
                Modelo.Range("B19") = Modelo.Range("B19") & "," & Âncoras.Cells(EmpresaEscolhida + 2, 13)
            End If
            
            Modelo.Range("B25") = Âncoras.Cells(EmpresaEscolhida + 2, 14)
            Modelo.Range("B29") = Âncoras.Cells(EmpresaEscolhida + 2, 15)
            Modelo.Range("B31") = Âncoras.Cells(EmpresaEscolhida + 2, 16)
            Modelo.Range("B33") = Âncoras.Cells(EmpresaEscolhida + 2, 17)
            Modelo.Range("B35") = Âncoras.Cells(EmpresaEscolhida + 2, 18)
            Modelo.Range("B37") = Âncoras.Cells(EmpresaEscolhida + 2, 19)
            
            Modelo.Range("B45") = Âncoras.Cells(EmpresaEscolhida + 2, 20)
            Modelo.Range("B47") = Âncoras.Cells(EmpresaEscolhida + 2, 21)
            Modelo.Range("B49") = Âncoras.Cells(EmpresaEscolhida + 2, 22)
            Modelo.Range("B51") = Âncoras.Cells(EmpresaEscolhida + 2, 23)
            Modelo.Range("B53") = Âncoras.Cells(EmpresaEscolhida + 2, 24)
            Modelo.Range("B55") = Âncoras.Cells(EmpresaEscolhida + 2, 25)
            Modelo.Range("B57") = Âncoras.Cells(EmpresaEscolhida + 2, 26)
            Modelo.Range("B59") = Âncoras.Cells(EmpresaEscolhida + 2, 27)
            Modelo.Range("B61") = Âncoras.Cells(EmpresaEscolhida + 2, 28)
            Modelo.Range("B63") = Âncoras.Cells(EmpresaEscolhida + 2, 29)
            Modelo.Range("B65") = Âncoras.Cells(EmpresaEscolhida + 2, 30)
            Modelo.Range("B67") = Âncoras.Cells(EmpresaEscolhida + 2, 31)
            
            Modelo.Range("B71") = Âncoras.Cells(EmpresaEscolhida + 2, 32)
            Modelo.Range("B73") = Âncoras.Cells(EmpresaEscolhida + 2, 33)
            Modelo.Select
        
        Cells(2, 2).Select
        End If
    Else
        MsgBox "Não há empresas cadastradas!"
    End If

End Sub

Sub ExcluirAncora()

    Dim Âncoras As Worksheet
    Set Âncoras = Sheets("Âncoras")
        
    ultima = Âncoras.Range("A1").End(xlDown).Row - 2

    If (ultima > 0) Then
        UserFormAncora.Show
        
        If (EmpresaEscolhida <> 0) Then
            ID = Âncoras.Cells(EmpresaEscolhida + 2, 1)
            numero_de_subs = Sheets("Subcritérios").Range("A1").End(xlDown).Row - 2
            
            'Apagar a ancora do danco de dados (mover a lista para cima)
            Âncoras.Rows(EmpresaEscolhida + 2).Delete Shift:=xlUp
            
            'Apagar a ancora do banco de dados de Pesos
            ultima_linha_peso = Sheets("Pesos").Range("A1").End(xlDown).Row
            If ultima_linha_peso > 2 Then
                    
                L = 3
                While (Sheets("Pesos").Cells(L, 1) <> ID) And (L <= ultima_linha_peso) 'para não ficar em loop inf
                      L = L + 1
                Wend
                
                If (L <= ultima_linha_peso) Then
                    'jeito novo
                    Sheets("Pesos").Rows(L).Delete Shift:=xlUp
                End If
              
            End If
            
            'Apagar onde tem a ancora do banco de dados de Notas
            ultima_linha_notas = Sheets("Notas").Range("A1").End(xlDown).Row
            If ultima_linha_notas > 2 Then
            
                L = 3
                For v = 1 To ultima_linha_notas - 2
                    If (Sheets("Notas").Cells(L, 1) = ID) Then
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
Sub VoltarAncora()
    Anc.Show
End Sub


