Attribute VB_Name = "�ncora"
Public EmpresaEscolhida As Integer

Sub AddAncora()

    Sheets("Pr�-question�rio �ncora").Select
    Cells(2, 2).Select
    EmpresaEscolhida = 0
    
End Sub
Sub SalvarAncora()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pr�-question�rio �ncora")
    
    Dim �ncoras As Worksheet
    Set �ncoras = Sheets("�ncoras")
    
    Dim Nome_�ncora As String
    Nome_�ncora = UCase(Questionario.Range("B5"))
    
    '-------------------------------------------------------
    
    ultima = �ncoras.Range("A1").End(xlDown).Row - 2
    existe = 0
    
        If (EmpresaEscolhida = 0) Then 'Adicionando ancora
        
            'Verificar campos obrigatorios
            If (Nome_�ncora = "" Or Questionario.Range("B15") = "" Or Questionario.Range("B81") = "") Then
                MsgBox ("Preencha todos os campos obrigat�rios antes de salvar!")
            
            Else
                For i = 1 To ultima
                    If �ncoras.Cells(i + 2, 2) = Nome_�ncora Then
                        MsgBox ("O nome da empresa j� existe!")
                        existe = 1
                    End If
                Next i
                
                If (existe = 0) Then
                    nova_linha = �ncoras.Range("A1").End(xlDown).Row + 1
                    If (nova_linha = 3) Then
                        �ncoras.Cells(nova_linha, 1) = "A1"
                    Else
                        �ncoras.Cells(nova_linha, 1) = "A" & Right$(�ncoras.Range("A1").End(xlDown).Value, Len(�ncoras.Range("A1").End(xlDown)) - 1) + 1
                    End If
                    GravarCampos (nova_linha)
                    MsgBox "Empresa cadastrada!"
                    
                    LimparAncora
                    VoltarAncora
                End If
                
            End If
            
        Else 'Editando ancora
            'Verificar campos obrigatorios
            If (Nome_�ncora = "" Or Questionario.Range("B15") = "" Or Questionario.Range("B81") = "") Then
                MsgBox ("Preencha todos os campos obrigat�rios antes de salvar!")
            Else
                'pode salvar altera��es
                GravarCampos (EmpresaEscolhida + 2) 'GravarCampos est� recebendo a linha do banco de dados referente � empresa escolhida
                MsgBox "Altera��es salvas!"
                LimparAncora
                VoltarAncora
            End If
            
        End If

End Sub

Sub LimparAncora()

    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pr�-question�rio �ncora")
    
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

    Dim �ncoras As Worksheet
    Set �ncoras = Sheets("�ncoras")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pr�-question�rio �ncora")
    
    '-------------------------------------------------------
    
    ultima = �ncoras.Range("A1").End(xlDown).Row - 2
    
    If (ultima > 0) Then
        UserFormAncora.Show 'Altera EmpresaEscolhida
        If (EmpresaEscolhida <> 0) Then
            'Question�rio recebe os dados do banco de dados
            Questionario.Range("B5") = �ncoras.Cells(EmpresaEscolhida + 2, 2)
            Questionario.Range("B7") = �ncoras.Cells(EmpresaEscolhida + 2, 3)
            Questionario.Range("B9") = �ncoras.Cells(EmpresaEscolhida + 2, 4)
            Questionario.Range("B11") = �ncoras.Cells(EmpresaEscolhida + 2, 5)
            Questionario.Range("B13") = �ncoras.Cells(EmpresaEscolhida + 2, 6)
            Questionario.Range("B15") = �ncoras.Cells(EmpresaEscolhida + 2, 7)
            Questionario.Range("B18") = �ncoras.Cells(EmpresaEscolhida + 2, 8)
            Questionario.Range("B20") = �ncoras.Cells(EmpresaEscolhida + 2, 9)
            Questionario.Range("B22") = �ncoras.Cells(EmpresaEscolhida + 2, 10)
            Questionario.Range("B25") = �ncoras.Cells(EmpresaEscolhida + 2, 11)
            Questionario.Range("B27") = �ncoras.Cells(EmpresaEscolhida + 2, 12)
            Questionario.Range("B29") = �ncoras.Cells(EmpresaEscolhida + 2, 13)
            Questionario.Range("B35") = �ncoras.Cells(EmpresaEscolhida + 2, 14)
            Questionario.Range("B39") = �ncoras.Cells(EmpresaEscolhida + 2, 15)
            Questionario.Range("B41") = �ncoras.Cells(EmpresaEscolhida + 2, 16)
            Questionario.Range("B43") = �ncoras.Cells(EmpresaEscolhida + 2, 17)
            Questionario.Range("B45") = �ncoras.Cells(EmpresaEscolhida + 2, 18)
            Questionario.Range("B47") = �ncoras.Cells(EmpresaEscolhida + 2, 19)
            Questionario.Range("B55") = �ncoras.Cells(EmpresaEscolhida + 2, 20)
            Questionario.Range("B57") = �ncoras.Cells(EmpresaEscolhida + 2, 21)
            Questionario.Range("B59") = �ncoras.Cells(EmpresaEscolhida + 2, 22)
            Questionario.Range("B61") = �ncoras.Cells(EmpresaEscolhida + 2, 23)
            Questionario.Range("B63") = �ncoras.Cells(EmpresaEscolhida + 2, 24)
            Questionario.Range("B65") = �ncoras.Cells(EmpresaEscolhida + 2, 25)
            Questionario.Range("B67") = �ncoras.Cells(EmpresaEscolhida + 2, 26)
            Questionario.Range("B69") = �ncoras.Cells(EmpresaEscolhida + 2, 27)
            Questionario.Range("B71") = �ncoras.Cells(EmpresaEscolhida + 2, 28)
            Questionario.Range("B73") = �ncoras.Cells(EmpresaEscolhida + 2, 29)
            Questionario.Range("B75") = �ncoras.Cells(EmpresaEscolhida + 2, 30)
            Questionario.Range("B77") = �ncoras.Cells(EmpresaEscolhida + 2, 31)
            Questionario.Range("B81") = �ncoras.Cells(EmpresaEscolhida + 2, 32)
            Questionario.Range("B83") = �ncoras.Cells(EmpresaEscolhida + 2, 33)
        
            Questionario.Select
            Cells(2, 2).Select
        End If
    Else
        MsgBox "N�o h� empresas cadastradas!"
    End If
        
End Sub

'Fun��o usada no UserForm.'Retorna ID em empresa escolhida
Sub ReceberAncora(Empresa As String)

    ultima = Sheets("�ncoras").Range("A1").End(xlDown).Row - 2
    'Pegar a linha
    EmpresaEscolhida = 0
        For i = 1 To ultima
            If (Sheets("�ncoras").Cells(2 + i, 2) = Empresa) Then
                EmpresaEscolhida = i
            End If
        Next i
         
End Sub
Sub GravarCampos(Linha As Integer)

    Dim �ncoras As Worksheet
    Set �ncoras = Sheets("�ncoras")
    
    Dim Questionario As Worksheet
    Set Questionario = Sheets("Pr�-question�rio �ncora")
    
    '-------------------------------------------------------
    
        'Banco de dados recebe os dados do Question�rio
        
        
        'gravar maiuscula
        Nome = UCase(Questionario.Range("B5"))
        �ncoras.Cells(Linha, 2) = Nome
        �ncoras.Cells(Linha, 3) = Questionario.Range("B7")
        �ncoras.Cells(Linha, 4) = Questionario.Range("B9")
        �ncoras.Cells(Linha, 5) = Questionario.Range("B11")
        �ncoras.Cells(Linha, 6) = Questionario.Range("B13")
        �ncoras.Cells(Linha, 7) = Questionario.Range("B15")
        �ncoras.Cells(Linha, 8) = Questionario.Range("B18")
        �ncoras.Cells(Linha, 9) = Questionario.Range("B20")
        �ncoras.Cells(Linha, 10) = Questionario.Range("B22")
        �ncoras.Cells(Linha, 11) = Questionario.Range("B25")
        �ncoras.Cells(Linha, 12) = Questionario.Range("B27")
        �ncoras.Cells(Linha, 13) = Questionario.Range("B29")
        �ncoras.Cells(Linha, 14) = Questionario.Range("B35")
        �ncoras.Cells(Linha, 15) = Questionario.Range("B39")
        �ncoras.Cells(Linha, 16) = Questionario.Range("B41")
        �ncoras.Cells(Linha, 17) = Questionario.Range("B43")
        �ncoras.Cells(Linha, 18) = Questionario.Range("B45")
        �ncoras.Cells(Linha, 19) = Questionario.Range("B47")
        �ncoras.Cells(Linha, 20) = Questionario.Range("B55")
        �ncoras.Cells(Linha, 21) = Questionario.Range("B57")
        �ncoras.Cells(Linha, 22) = Questionario.Range("B59")
        �ncoras.Cells(Linha, 23) = Questionario.Range("B61")
        �ncoras.Cells(Linha, 24) = Questionario.Range("B63")
        �ncoras.Cells(Linha, 25) = Questionario.Range("B65")
        �ncoras.Cells(Linha, 26) = Questionario.Range("B67")
        �ncoras.Cells(Linha, 27) = Questionario.Range("B69")
        �ncoras.Cells(Linha, 28) = Questionario.Range("B71")
        �ncoras.Cells(Linha, 29) = Questionario.Range("B73")
        �ncoras.Cells(Linha, 30) = Questionario.Range("B75")
        �ncoras.Cells(Linha, 31) = Questionario.Range("B77")
        �ncoras.Cells(Linha, 32) = Questionario.Range("B81")
        �ncoras.Cells(Linha, 33) = Questionario.Range("B83")
        
 
End Sub
Sub VisualizarAncora()

    Dim �ncoras As Worksheet
    Set �ncoras = Sheets("�ncoras")
    
    Dim Modelo As Worksheet
    Set Modelo = Sheets("Visualiza��o �ncora")
    
    '-------------------------------------------------------
    
    ultima = Sheets("�ncoras").Range("A1").End(xlDown).Row - 2

    If (ultima > 0) Then
        UserFormAncora.Show
        
        If (EmpresaEscolhida <> 0) Then
            'Mostrar os dados do banco de dados para o usu�rio
            
            Modelo.Range("B5") = �ncoras.Cells(EmpresaEscolhida + 2, 2)
            Modelo.Range("B7") = �ncoras.Cells(EmpresaEscolhida + 2, 3)
            Modelo.Range("B9") = �ncoras.Cells(EmpresaEscolhida + 2, 4)
            Modelo.Range("B11") = �ncoras.Cells(EmpresaEscolhida + 2, 5)
            Modelo.Range("B13") = �ncoras.Cells(EmpresaEscolhida + 2, 6)
            Modelo.Range("B15") = �ncoras.Cells(EmpresaEscolhida + 2, 7)
            
            Modelo.Range("B17") = �ncoras.Cells(EmpresaEscolhida + 2, 8)
            If (�ncoras.Cells(EmpresaEscolhida + 2, 9) <> "") Then
                Modelo.Range("B17") = Modelo.Range("B17") & "," & �ncoras.Cells(EmpresaEscolhida + 2, 9)
            End If
            If (�ncoras.Cells(EmpresaEscolhida + 2, 10) <> "") Then
                Modelo.Range("B17") = Modelo.Range("B17") & "," & �ncoras.Cells(EmpresaEscolhida + 2, 10)
            End If
            
            Modelo.Range("B19") = �ncoras.Cells(EmpresaEscolhida + 2, 11)
            If (�ncoras.Cells(EmpresaEscolhida + 2, 12) <> "") Then
                Modelo.Range("B19") = Modelo.Range("B19") & "," & �ncoras.Cells(EmpresaEscolhida + 2, 12)
            End If
            If (�ncoras.Cells(EmpresaEscolhida + 2, 13) <> "") Then
                Modelo.Range("B19") = Modelo.Range("B19") & "," & �ncoras.Cells(EmpresaEscolhida + 2, 13)
            End If
            
            Modelo.Range("B25") = �ncoras.Cells(EmpresaEscolhida + 2, 14)
            Modelo.Range("B29") = �ncoras.Cells(EmpresaEscolhida + 2, 15)
            Modelo.Range("B31") = �ncoras.Cells(EmpresaEscolhida + 2, 16)
            Modelo.Range("B33") = �ncoras.Cells(EmpresaEscolhida + 2, 17)
            Modelo.Range("B35") = �ncoras.Cells(EmpresaEscolhida + 2, 18)
            Modelo.Range("B37") = �ncoras.Cells(EmpresaEscolhida + 2, 19)
            
            Modelo.Range("B45") = �ncoras.Cells(EmpresaEscolhida + 2, 20)
            Modelo.Range("B47") = �ncoras.Cells(EmpresaEscolhida + 2, 21)
            Modelo.Range("B49") = �ncoras.Cells(EmpresaEscolhida + 2, 22)
            Modelo.Range("B51") = �ncoras.Cells(EmpresaEscolhida + 2, 23)
            Modelo.Range("B53") = �ncoras.Cells(EmpresaEscolhida + 2, 24)
            Modelo.Range("B55") = �ncoras.Cells(EmpresaEscolhida + 2, 25)
            Modelo.Range("B57") = �ncoras.Cells(EmpresaEscolhida + 2, 26)
            Modelo.Range("B59") = �ncoras.Cells(EmpresaEscolhida + 2, 27)
            Modelo.Range("B61") = �ncoras.Cells(EmpresaEscolhida + 2, 28)
            Modelo.Range("B63") = �ncoras.Cells(EmpresaEscolhida + 2, 29)
            Modelo.Range("B65") = �ncoras.Cells(EmpresaEscolhida + 2, 30)
            Modelo.Range("B67") = �ncoras.Cells(EmpresaEscolhida + 2, 31)
            
            Modelo.Range("B71") = �ncoras.Cells(EmpresaEscolhida + 2, 32)
            Modelo.Range("B73") = �ncoras.Cells(EmpresaEscolhida + 2, 33)
            Modelo.Select
        
        Cells(2, 2).Select
        End If
    Else
        MsgBox "N�o h� empresas cadastradas!"
    End If

End Sub

Sub ExcluirAncora()

    Dim �ncoras As Worksheet
    Set �ncoras = Sheets("�ncoras")
        
    ultima = �ncoras.Range("A1").End(xlDown).Row - 2

    If (ultima > 0) Then
        UserFormAncora.Show
        
        If (EmpresaEscolhida <> 0) Then
            ID = �ncoras.Cells(EmpresaEscolhida + 2, 1)
            numero_de_subs = Sheets("Subcrit�rios").Range("A1").End(xlDown).Row - 2
            
            'Apagar a ancora do danco de dados (mover a lista para cima)
            �ncoras.Rows(EmpresaEscolhida + 2).Delete Shift:=xlUp
            
            'Apagar a ancora do banco de dados de Pesos
            ultima_linha_peso = Sheets("Pesos").Range("A1").End(xlDown).Row
            If ultima_linha_peso > 2 Then
                    
                L = 3
                While (Sheets("Pesos").Cells(L, 1) <> ID) And (L <= ultima_linha_peso) 'para n�o ficar em loop inf
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
        MsgBox "N�o h� empresas cadastradas!"
    End If
    
End Sub
Sub VoltarAncora()
    Anc.Show
End Sub


