Attribute VB_Name = "Notas"
Dim Pos As Integer 'Posição a ser colocado o criterio

Dim LinhaAncora As Integer
Public Sub AvaliarFornecedor()

Application.ScreenUpdating = False


    If Sheets("Âncoras").Range("A1").End(xlDown).Row - 2 > 0 Then 'verifica se há ancoras
        UserFormAncora.Show 'EmpresaEscolhida
        
        If EmpresaEscolhida <> 0 Then
        
            If (EncontraLinhaAncora <> 0) Then
                'verificva se ha pelo menos um peso <> de 0
                If (VerificaPesos(EncontraLinhaAncora)) Then
                
                    If Sheets("Fornecedores").Range("A1").End(xlDown).Row - 2 > 0 Then 'verifica se há forncedores
                        UserFormFornecedor.Show 'FornecedorEscolhido
            
                        If FornecedorEscolhido <> 0 Then
                            Sheets("Avaliação do fornecedor").Select
                            LimparNota
                            Application.ScreenUpdating = False
                            'insere nome do fornecedor
                            Sheets("Avaliação do fornecedor").Range("B5") = "EMPRESA AVALIADA: " & Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 2)
                    
                            '---------------
                            'Inicio Impacto financeiro
                            Sheets("Avaliação do fornecedor").Range("B7") = "IMPACTO FINANCEIRO"
                            ContR = 0 'Posição inicial da listagem temporaria de posiçoes dos criteriso de risco de fornecimento
                            Pos = 10
                
                
                            'Separa em I ou R, se I grava na planilha, se R guarda no vetor
                            For crite = 1 To Sheets("Critérios").Range("A1").End(xlDown).Row - 2 'i-esimo criterio
                                If (Sheets("Critérios").Cells(crite + 2, 4) = "I") Then
                                    Call EscreverCriterios(Pos, crite)
                                Else
                                    ContR = ContR + 1
                                    Sheets("Avaliação do fornecedor").Cells(ContR, 1) = crite
                                End If
                            Next crite
                    
                            'Inicio Risco de fornecimento
                
                                
                                If (Pos <> 11) Then 'há criteiros de impacto financeiro
                                    'formatação do titulo "Risco de fornecimento"
                                    Pos0R = Pos 'Posição onde começo a trabalhar com "risco de fornecimento"
                                    Pos = Pos + 1
                            
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2) = "RISCO DE FORNECIMENTO"
                                    Sheets("Avaliação do fornecedor").Range(Cells(Pos, 2), Cells(Pos, 3)).Merge
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).HorizontalAlignment = xlCenter
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).VerticalAlignment = xlCenter
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Interior.Color = RGB(117, 113, 113)
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Color = RGB(242, 242, 242)
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Size = 12
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Bold = True
            
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 2) = "Critérios"
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 2).Font.Bold = True
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 2).Font.Size = 12
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 3) = "Notas"
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 3).Font.Bold = True
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 3).Font.Size = 12
                       
                                    Pos = Pos + 4
                                Else
                                    Sheets("Avaliação do fornecedor").Range("B7") = "RISCO DE FORNECIMENTO"
                                End If
                        
                                For i = 1 To ContR
                                    crite = Sheets("Avaliação do fornecedor").Cells(i, 1)
                                    Sheets("Avaliação do fornecedor").Cells(i, 1) = ""
                                    Call EscreverCriterios(Pos, crite)
                               
                                Next i
                                
                                If (Pos0R + 5 = Pos) Then 'Não há pesos em nenhum criterio do eixo R
                                    Pos = Pos - 4
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2) = ""
                                    
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).HorizontalAlignment = xlLeft
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).VerticalAlignment = xlCenter
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Interior.Pattern = xlNone
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Color = RGB(0, 0, 0)
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Size = 12
                                    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Bold = False
                                    Sheets("Avaliação do fornecedor").Range(Cells(Pos, 2), Cells(Pos, 3)).UnMerge
            
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 2) = ""
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 2).Font.Bold = False
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 2).Font.Size = 12
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 3) = ""
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 3).Font.Bold = False
                                    Sheets("Avaliação do fornecedor").Cells(Pos + 2, 3).Font.Size = 12
                                    Pos = Pos - 1
                                End If
                            'bota no lugar os botoes
                            Sheets("Avaliação do fornecedor").Shapes("VoltarMenu").Top = Sheets("Avaliação do fornecedor").Cells(Pos, 2).Top
                            Sheets("Avaliação do fornecedor").Shapes("Salvar").Top = Sheets("Avaliação do fornecedor").Cells(Pos, 2).Top
            
                            'Preenche pesos ja preenchidos anteriormente
                            PreencherNotas
                            Application.ScreenUpdating = True
                            Sheets("Avaliação do fornecedor").Range("B3").Select
                    
                        
                                
                        End If
                    
                    Else
                        FornecedorEscolhido = 0
                        MsgBox "Não há empresas fornecedoras cadastradas!"
                    End If 'do quinto if
                
            Else
                EmpresaEscolhida = 0
                MsgBox "A empresa ainda não deu peso aos critérios!"

            End If 'do quarto If
        Else
            EmpresaEscolhida = 0
             MsgBox "A empresa ainda não deu peso aos critérios!"
        End If 'do terceiro If
        
    End If 'do segundo if
        
    Else 'do primeiro If
        EmpresaEscolhida = 0
        MsgBox "Não há empresas cadastradas!"
        
    End If
   
   
End Sub
Sub EscreverCriterios(ByRef Pos As Integer, ByVal crite As Integer)
    'Adiciona o criterio. Verifica no criterio foi dado peso . Caso sim grava, caso o contrario pula para o proximo.
    
    Sheets("Avaliação do fornecedor").Select
    
    LinhaAncora = EncontraLinhaAncora
    ID = Sheets("Critérios").Cells(crite + 2, 1)
    ColunaCrite = 2
    While (Sheets("Pesos").Cells(1, ColunaCrite) <> ID And Sheets("Pesos").Cells(1, ColunaCrite) <> "")
        ColunaCrite = ColunaCrite + 1
    Wend
    
    If (Sheets("Pesos").Cells(LinhaAncora, ColunaCrite) > 0 And Sheets("Pesos").Cells(1, ColunaCrite) <> "") Then
        'Adiciona o criterio e formatação
    Sheets("Avaliação do fornecedor").Cells(Pos, 2) = Sheets("Critérios").Cells(crite + 2, 2) 'nome do criterio
    'Formata o nome do criterio
        Call FormataçãocriterioN(Pos, crite) 'Insere caixa de texto da descrição do critério
        Call InserirCaixaDesc(Pos, crite)
            
            Sheets("Avaliação do fornecedor").TextBoxes("D" & Pos).Text = Sheets("Critérios").Cells(crite + 2, 3) 'Descrição do criterio
        
   
        
     
                    'Insere caixas de grupo
                        L = Sheets("Avaliação do fornecedor").Cells(Pos, 3).Left
                        T = Sheets("Avaliação do fornecedor").Cells(Pos, 3).Top
                        W = Sheets("Avaliação do fornecedor").Cells(Pos, 4).Left - L
                        h = Sheets("Avaliação do fornecedor").Cells(Pos + 1, 3).Top - T

                        Sheets("Avaliação do fornecedor").GroupBoxes.Add(L, T, W, h).Select
                        Selection.Name = ID
                        Selection.Visible = False
                        Selection.Caption = ""
                    
                    'Insere os optionsbuttons
                        For cont = 1 To 5
                            T = Sheets("Avaliação do fornecedor").Cells(Pos, 3).Top
                            E = Sheets("Avaliação do fornecedor").Cells(Pos, 3).Left
                            W = 15
                            h = 15
                            Sheets("Avaliação do fornecedor").OptionButtons.Add(E + cont * 25, T, W, h).Select
                            Selection.Name = ID & "O" & cont
                            Selection.Display3DShading = True
                            Selection.Caption = cont 'legenda d 1 a 5
                            
                            
                        Next cont
            
                    Pos = Pos + 6 ' linha em branco para separar criterios
                    

        
        
        End If
       
    
        
    
        
End Sub
'
'

Sub InserirCaixaDesc(ByRef Pos As Integer, ByVal crite As Integer)
            'Insere caixa de texto da descrição do critério
        
        T = Sheets("Avaliação do fornecedor").Cells(Pos + 1, 2).Top + 10.5
        E = Sheets("Avaliação do fornecedor").Cells(Pos + 1, 2).Left + 3.75
        W = 636
        h = 60.75
            
        Sheets("Avaliação do fornecedor").TextBoxes.Add(E, T, W, h).Select
        Selection.Name = "D" & Pos
        
        'formatação do fundo da caixa
       'Selection.Interior.Color = RGB(217, 225, 242)
        
        'formatação de tras da caixa
        If (Sheets("Critérios").Cells(crite + 2, 7) = "I") Then
            For Linha = Pos + 1 To Pos + 4
                Sheets("Avaliação do fornecedor").Cells(Linha, 2).Interior.Color = RGB(169, 208, 142)
                Sheets("Avaliação do fornecedor").Cells(Linha, 3).Interior.Color = RGB(169, 208, 142)
            Next
        End If
        If (Sheets("Critérios").Cells(crite + 2, 7) = "P") Then
        For Linha = Pos + 1 To Pos + 4
                Sheets("Avaliação do fornecedor").Cells(Linha, 2).Interior.Color = RGB(221, 235, 247)
                Sheets("Avaliação do fornecedor").Cells(Linha, 3).Interior.Color = RGB(221, 235, 247)
            Next
        End If
End Sub
'


'

Function EncontraLinhaAncora() As Integer
    'Encontra na planilha de pesos
    IDEmpresa = Sheets("Âncoras").Cells(EmpresaEscolhida + 2, 1)
    EncontraLinhaAncora = 3
        While (Sheets("Pesos").Cells(EncontraLinhaAncora, 1) <> IDEmpresa And EncontraLinhaAncora <= Sheets("Pesos").Range("A1").End(xlDown).Row)
            EncontraLinhaAncora = EncontraLinhaAncora + 1
        Wend
        If (EncontraLinhaAncora = Sheets("Pesos").Range("A1").End(xlDown).Row + 1) Then
            EncontraLinhaAncora = 0
        End If
End Function


Function EncontraColuna(ByVal ID As String) As Integer
    i = 1
    While (Sheets("Pesos").Cells(1, i + 1) <> "")
        If (Sheets("Pesos").Cells(1, i + 1) = ID) Then
            EncontraColuna = i + 1
        End If
        i = i + 1
    Wend
    If (EncontraColuna = 0) Then 'caso o criterio nao esteja sido avaliado em nenhuma ancora
        EncontraColuna = i + 1
    End If
    
End Function
Sub PreencherNotas()

    'Encotra a linha correta pelo ID da ancora, se nao existir o ID nao fazer nada
        LinhaAncora = EncontraLinhaAncora
        
        
        For i = 3 To Sheets("Notas").Range("A1").End(xlDown).Row
            
            If (Sheets("Notas").Cells(i, 1) = Sheets("Âncoras").Cells(EmpresaEscolhida + 2, 1) And _
                Sheets("Notas").Cells(i, 2) = Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 1)) Then 'Se a linha for da ancora e fornecedor
                
                Linha = i
                
                'Se o criterio tiver peso e nota entao grava no optionbutton
                    
                    Coluna = 3
                    While (Sheets("Notas").Cells(1, Coluna) <> "")
                        ID = Sheets("Notas").Cells(1, Coluna)
                        ColunaPeso = EncontraColuna(ID)
                        If (Sheets("Pesos").Cells(LinhaAncora, ColunaPeso) > 0 And Sheets("Notas").Cells(Linha, Coluna) > 0) Then
                            ValorNota = Sheets("Notas").Cells(Linha, Coluna)
                            Sheets("Avaliação do fornecedor").OptionButtons(ID & "O" & ValorNota).Value = True
                        End If
                        Coluna = Coluna + 1
                    Wend
            End If
        Next i


End Sub
Sub SalvarNotas()
    'Verifica se todos os criteriso escolhidos foram preenchidos
    If (Preenchido) Then
        
        'Encotra a linha correta pelo ID da ancora, se nao existir o ID cria nova linha
        
        LinhaNota = 1
        While (Sheets("Notas").Cells(LinhaNota, 1) <> Sheets("Âncoras").Cells(EmpresaEscolhida + 2, 1) Or _
                Sheets("Notas").Cells(LinhaNota, 2) <> Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 1))
                
            If (Sheets("Notas").Cells(LinhaNota, 1) <> "") Then
                LinhaNota = LinhaNota + 1
            Else 'Não foi criado ainda
                
                Sheets("Notas").Cells(LinhaNota, 1) = Sheets("Âncoras").Cells(EmpresaEscolhida + 2, 1)
                Sheets("Notas").Cells(LinhaNota, 2) = Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 1)
                Vazioigual0AncN (LinhaNota)
            End If
                
        Wend
        
        'procura a celula para gravar
        ColunaPeso = 2
        LinhaPeso = EncontraLinhaAncora
        
        While (Sheets("Pesos").Cells(1, ColunaPeso) <> "")
            If (Sheets("Pesos").Cells(LinhaPeso, ColunaPeso) <> 0) Then
                ID = Sheets("Pesos").Cells(1, ColunaPeso)
                ColunaNota = 3
            
                While (Sheets("Notas").Cells(1, ColunaNota) <> ID)
                    If (Sheets("Notas").Cells(1, ColunaNota) <> "") Then
                        ColunaNota = ColunaNota + 1
                    Else
                        Sheets("Notas").Cells(1, ColunaNota) = ID
                        Vazioigual0CritN (ColunaNota)
                    End If
                Wend
                
                'grava na celula
                Nota = 0
                For cont = 1 To 5
                    If (Sheets("Avaliação do fornecedor").OptionButtons(ID & "O" & cont).Value = 1) Then
                        Nota = cont
                    End If
                Next cont
                'se a nota estiver vazio desconsidera o peso
                If (Nota <> 0) Then
                    Sheets("Notas").Cells(LinhaNota, ColunaNota) = Nota
                End If
                
            End If
            ColunaPeso = ColunaPeso + 1
        Wend
        
        MsgBox "As notas foram salvas!"
        'LimparNota
        ' em vez de VoltarAvaliar ( que é retornar ao menu) pergunta se quer avaliar outro fornecedor
        ' não tem isso ainda Av1.Show
        VoltarAvaliar
        
    Else
    MsgBox "É obrigatorio a avaliação de todos os critérios!"
    End If
End Sub

Sub LimparNota()

Application.ScreenUpdating = False
Application.DisplayAlerts = False


    For i = 1 To Sheets("Avaliação do fornecedor").GroupBoxes.Count
        Sheets("Avaliação do fornecedor").GroupBoxes.Delete
    Next i
    For i = 1 To Sheets("Avaliação do fornecedor").OptionButtons.Count
        Sheets("Avaliação do fornecedor").OptionButtons.Delete
    Next i
    For i = 1 To Sheets("Avaliação do fornecedor").TextBoxes.Count
        Sheets("Avaliação do fornecedor").TextBoxes.Delete
    Next i
    
    Sheets("Avaliação do fornecedor").Select
    For i = 11 To Sheets("Avaliação do fornecedor").Shapes("VoltarMenu").Top

        'Desmesclar celulas
        Sheets("Avaliação do fornecedor").Range(Cells(i, 2), Cells(i, 3)).Merge
        Sheets("Avaliação do fornecedor").Range(Cells(i, 2), Cells(i, 3)).UnMerge
    Next i
    Plan21.Range("B9:C5000").ClearFormats
    Plan21.Range("B9:C5000").ClearContents
    Sheets("Avaliação do fornecedor").Shapes("VoltarMenu").Top = Sheets("Avaliação do fornecedor").Cells(11, 2).Top
    Sheets("Avaliação do fornecedor").Shapes("Salvar").Top = Sheets("Avaliação do fornecedor").Cells(11, 2).Top
    
 Application.DisplayAlerts = True
Application.ScreenUpdating = True
    
    
End Sub

Sub VoltarAvaliar()
    LimparNota
    'Vai para a planilha em branco antes de voltar ao menu
    Sheets("Projeto IEL").Select
    aval.Show vbModeless
    
End Sub


Function Preenchido() As Boolean
    'Verifica se ao menos um dos optionbuttons de cada grupo foi escolhido

    Preenchido = True
    For Grupo = 1 To Sheets("Avaliação do fornecedor").OptionButtons.Count / 5
        'Se todos desmarcados entao preenchido = false (-4146)
        If (Sheets("Avaliação do fornecedor").OptionButtons((Grupo - 1) * 5 + 1) = -4146 And _
         Sheets("Avaliação do fornecedor").OptionButtons((Grupo - 1) * 5 + 2) = -4146 And _
         Sheets("Avaliação do fornecedor").OptionButtons((Grupo - 1) * 5 + 3) = -4146 And _
         Sheets("Avaliação do fornecedor").OptionButtons((Grupo - 1) * 5 + 4) = -4146 And _
         Sheets("Avaliação do fornecedor").OptionButtons((Grupo - 1) * 5 + 5) = -4146) Then
            Preenchido = False
        End If
        
    Next Grupo
End Function

Sub Vazioigual0CritN(Coluna As Integer) 'por zero aodne teria vazio ao avaliar novo criteiro (notas)
    Linha = 3
    While (Sheets("Notas").Cells(Linha, 1) <> "")
        If (Sheets("Notas").Cells(Linha, Coluna) = "") Then
            Sheets("Notas").Cells(Linha, Coluna) = 0
        End If
        Linha = Linha + 1
    Wend
End Sub
Sub Vazioigual0AncN(Linha As Integer) ' por zero aodne teria vazio ao avaliar nova ancora (notas)
    Coluna = 3
    While (Sheets("Notas").Cells(1, Coluna) <> "")
        If (Sheets("Notas").Cells(Linha, Coluna) = "") Then
            Sheets("Notas").Cells(Linha, Coluna) = 0
        End If
        Coluna = Coluna + 1
    Wend
  
    
End Sub
Sub FormataçãocriterioN(Pos As Integer, ByVal crite As Integer)
        
    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Italic = True 'italico
    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Underline = False
    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Font.Size = 14
 

    Sheets("Avaliação do fornecedor").Cells(Pos, 2).Select
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 180
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    'Se inverso
    If (Sheets("Critérios").Cells(crite + 2, 7) = "I") Then
        With Selection.Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        End With
    End If
       'Se proporcional
    If (Sheets("Critérios").Cells(crite + 2, 7) = "P") Then
        With Selection.Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        End With
    End If
    
    Sheets("Avaliação do fornecedor").Cells(Pos, 3).Interior.Color = RGB(255, 255, 255)
End Sub

Function VerificaPesos(Linha As Integer) As Boolean

    VerificaPesos = False
    c = 2
    While ((Sheets("Pesos").Cells(Linha, c) <> "") And (VerificaPesos = False))
        If (Sheets("Pesos").Cells(Linha, c) = 0) Then
            c = c + 1
        Else
            VerificaPesos = True
        End If
    Wend
    
End Function



