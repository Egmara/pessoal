Attribute VB_Name = "Notas"
Dim Pos As Integer 'Posi��o a ser colocado o criterio/Subcriteiro

Dim LinhaAncora As Integer
Public Sub AvaliarFornecedor()

    If Sheets("�ncoras").Range("A1").End(xlDown).Row - 2 > 0 Then 'veficca se h� ancoras
        UserFormAncora.Show 'EmpresaEscolhida
        If EmpresaEscolhida <> 0 Then
        
            If (EncontraLinhaAncora <> 0) Then
                'verificva se ha pelo menos um peso <> de 0
                If (VerificaPesos(EncontraLinhaAncora)) Then
                
                    If Sheets("Fornecedores").Range("A1").End(xlDown).Row - 2 > 0 Then 'verifica se h� forncedores
                        UserFormFornecedor.Show 'FornecedorEscolhido
            
                        If FornecedorEscolhido <> 0 Then
                    
                            'insere nome do fornecedor
                            Sheets("Avalia��o do fornecedor").Range("B4") = "Empresa avaliada: " & Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 2)
                    
                            '---------------
                            'Inicio Impacto financeiro
                            Sheets("Avalia��o do fornecedor").Range("B6") = "Impacto Financeiro"
                            ContR = 0 'Posi��o inicial da listagem temporaria de posi�oes dos criteriso de risco de fornecimento
                            Pos = 10
                
                
                            'Separa em I ou R, se I grava na planilha, se R guarda no vetor
                            For crite = 1 To Sheets("Crit�rios").Range("A1").End(xlDown).Row - 2 'i-esimo criterio
                                If (Sheets("Crit�rios").Cells(crite + 2, 4) = "I") Then
                                    Call EscreverCriterios(Pos, crite)
                                Else
                                    ContR = ContR + 1
                                    Sheets("Avalia��o do fornecedor").Cells(ContR, 1) = crite
                                End If
                            Next crite
                    
                            'Inicio Risco de fornecimento
                
                                
                                If (Pos <> 10) Then 'h� criteiros de impacto financeiro
                                    'formata��o do titulo "Risco de fornecimento"
                                    Pos0R = Pos 'Posi��o onde come�o a trabalhar com "risco de fornecimento"
                                    Pos = Pos + 1
                            
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2) = "Risco de fornecimento"
                                    Sheets("Avalia��o do fornecedor").Range(Cells(Pos, 2), Cells(Pos, 3)).Merge
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).HorizontalAlignment = xlCenter
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).VerticalAlignment = xlCenter
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Interior.Color = RGB(117, 113, 113)
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Color = RGB(242, 242, 242)
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Size = 12
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Bold = True
            
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 2) = "Crit�rios"
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 2).Font.Bold = True
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 2).Font.Size = 12
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 3) = "Notas"
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 3).Font.Bold = True
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 3).Font.Size = 12
                       
                                    Pos = Pos + 4
                                Else
                                    Sheets("Avalia��o do fornecedor").Range("B6") = "Risco de fornecimento"
                                End If
                        
                                For i = 1 To ContR
                                    crite = Sheets("Avalia��o do fornecedor").Cells(i, 1)
                                    Sheets("Avalia��o do fornecedor").Cells(i, 1) = ""
                                    Call EscreverCriterios(Pos, crite)
                               
                                Next i
                                
                                If (Pos0R + 5 = Pos) Then 'N�o h� pesos em nenhum subcriterio do eixo R
                                    Pos = Pos - 4
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2) = ""
                                    
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).HorizontalAlignment = xlLeft
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).VerticalAlignment = xlCenter
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Interior.Pattern = xlNone
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Color = RGB(0, 0, 0)
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Size = 12
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Bold = False
                                    Sheets("Avalia��o do fornecedor").Range(Cells(Pos, 2), Cells(Pos, 3)).UnMerge
            
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 2) = ""
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 2).Font.Bold = False
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 2).Font.Size = 12
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 3) = ""
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 3).Font.Bold = False
                                    Sheets("Avalia��o do fornecedor").Cells(Pos + 2, 3).Font.Size = 12
                                    Pos = Pos - 1
                                End If
                            'bota no lugar os botoes
                            Sheets("Avalia��o do fornecedor").Shapes("VoltarMenu").Top = Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Top
                            Sheets("Avalia��o do fornecedor").Shapes("Salvar").Top = Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Top
            
                            'Preenche pesos ja preenchidos anteriormente
                            PreencherNotas
                
                            Sheets("Avalia��o do fornecedor").Range("B2").Select
                    
                        
                                
                        End If
                    
                    Else
                        MsgBox "N�o h� empresas fornecedoras cadastradas!"
                    End If 'do quinto if
                
            Else
                MsgBox "A empresa ainda n�o deu peso aos crit�rios!"
            End If 'do quarto If
        Else
             MsgBox "A empresa ainda n�o deu peso aos crit�rios!"
        End If 'do terceiro If
        
    End If 'do segundo if
        
    Else 'do primeiro If
        MsgBox "N�o h� empresas cadastradas!"
    End If
   
End Sub
Sub EscreverCriterios(ByRef Pos As Integer, ByVal crite As Integer)
    'Adiciona o criterio. Verifica no criterio foi dado peso para algum dos subcriterios. Caso sim grava, caso o contrario pula para o proximo.
    
    Sheets("Avalia��o do fornecedor").Select
    
    'Adiciona o criterio e formata��o
    Sheets("Avalia��o do fornecedor").Cells(Pos, 2) = Sheets("Crit�rios").Cells(crite + 2, 2) 'nome do criterio
    'Formata o nome do criterio
    Formata��ocriterioN (Pos)
    
    LinhaAncora = EncontraLinhaAncora
    
        'Insere caixa de texto da descri��o do crit�rio
            Call InserirCaixaDesc(Pos)
            
            Sheets("Avalia��o do fornecedor").TextBoxes("D" & Pos).Text = Sheets("Crit�rios").Cells(crite + 2, 3) 'Descri��o do criterio
        
        Pos = Pos + 5
    
    'Subcriterios
        QntSub = 0 ' quantidade de sub no criterio
        Ativo = False ' Indica se o criterio pussui ao menos um sub ativo (true) ou nao (false)
        While (Sheets("Crit�rios").Cells(crite + 2, QntSub + 7) <> "") ' recebe o ID do subc relacionado ao criterio
            QntSub = QntSub + 1
            ID = Sheets("Crit�rios").Cells(crite + 2, QntSub + 6)
            LSubc = EncontraLinhaLSubc(ID) 'i-esimo (linha-2) sub da lista de Subcriterios
            Csubc = EncontraColunaSubc(ID) 'i-esimo sub (coluna-1) da plan pesos
            
            'se o sub recebeu peso
                If (Sheets("Pesos").Cells(LinhaAncora, Csubc + 1) > 0) Then
                    Ativo = True
            
                    'Insere caixa de texto da descri��o do subcrit�rio
                        Call InserirCaixaDesc(Pos)
                        Sheets("Avalia��o do fornecedor").TextBoxes("D" & Pos).Text = Sheets("Subcrit�rios").Cells(LSubc + 2, 3) 'Descri��o do Subcriterio
                    'nome do subcriterio e formata��o
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2) = "Subcrit�rio " & QntSub & ": " & Sheets("Subcrit�rios").Cells(LSubc + 2, 2)
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Bold = False
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Italic = True
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Underline = True
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Size = 11
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Interior.Color = RGB(189, 215, 238)
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Interior.Color = RGB(189, 215, 238)
                    
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Borders(xlEdgeLeft).Weight = xlThick
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 4).Borders(xlEdgeLeft).Weight = xlThick
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Borders(xlEdgeLeft).Color = RGB(221, 235, 247)
                    Sheets("Avalia��o do fornecedor").Cells(Pos, 4).Borders(xlEdgeLeft).Color = RGB(221, 235, 247)
              
              
                    'Insere caixas de grupo
                        L = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Left
                        T = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Top
                        W = Sheets("Avalia��o do fornecedor").Cells(Pos, 4).Left - L
                        H = Sheets("Avalia��o do fornecedor").Cells(Pos + 1, 3).Top - T

                        Sheets("Avalia��o do fornecedor").GroupBoxes.Add(L, T, W, H).Select
                        Selection.Name = ID
                        Selection.Visible = False
                        Selection.Caption = ""
                    
                    'Insere os optionsbuttons
                        For cont = 1 To 5
                            T = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Top
                            E = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Left
                            W = 15
                            H = 15
                            Sheets("Avalia��o do fornecedor").OptionButtons.Add(E + cont * 25, T, W, H).Select
                            Selection.Name = ID & "O" & cont
                            Selection.Display3DShading = True
                            Selection.Caption = cont 'legenda d 1 a 5
                            
                            
                        Next cont
            
                    Pos = Pos + 5
                End If
                
        Wend
        
        Pos = Pos + 1 ' linha em branco para separar criterios
        
    'Apaga criterios inuteis
        If (Not Ativo) Then 'se nao h� subc entao apagar criterio
            Pos = Pos - 6
            For i = 0 To 6
                Sheets("Avalia��o do fornecedor").Range(Cells(Pos + i, 2), Cells(Pos + i, 3)).Interior.Pattern = xlNone
                Sheets("Avalia��o do fornecedor").Cells(Pos + i, 3).Borders(xlEdgeLeft).LineStyle = xlNone
                Sheets("Avalia��o do fornecedor").Cells(Pos + i, 4).Borders(xlEdgeLeft).LineStyle = xlNone
            Next i
            Sheets("Avalia��o do fornecedor").TextBoxes("D" & Pos).Delete
            Sheets("Avalia��o do fornecedor").Cells(Pos, 2) = ""
            
        End If
        
End Sub
'
'

Sub InserirCaixaDesc(ByRef Pos As Integer)
            'Insere caixa de texto da descri��o do crit�rio
        
        T = Sheets("Avalia��o do fornecedor").Cells(Pos + 1, 2).Top + 10.5
        E = Sheets("Avalia��o do fornecedor").Cells(Pos + 1, 2).Left + 3.75
        W = 636
        H = 60.75
            
        Sheets("Avalia��o do fornecedor").TextBoxes.Add(E, T, W, H).Select
        Selection.Name = "D" & Pos
        
        'formata��o do fundo da caixa
       'Selection.Interior.Color = RGB(217, 225, 242)
        
        'formata��o de tras da caixa
        For Linha = Pos + 1 To Pos + 4
            Sheets("Avalia��o do fornecedor").Cells(Linha, 2).Interior.Color = RGB(221, 235, 247)
            Sheets("Avalia��o do fornecedor").Cells(Linha, 3).Interior.Color = RGB(221, 235, 247)
        Next
        
End Sub
'

Function EncontraLinhaLSubc(ByVal ID As String) As Integer
    EncontraLinhaSub = 0
    While (Sheets("Subcrit�rios").Cells(EncontraLinhaLSubc + 2, 1) <> ID)
        EncontraLinhaLSubc = EncontraLinhaLSubc + 1
    Wend
End Function
'

Function EncontraColunaSubc(ByVal ID As String) As Integer
    i = 0
    While (Sheets("Pesos").Cells(1, i + 1) <> "")
        If (Sheets("Pesos").Cells(1, i + 1) = ID) Then
            EncontraColunaSubc = i
        End If
        i = i + 1
    Wend
    If (EncontraColunaSubc = 0) Then 'caso o subcriterio nao esteja sido avaliado em nenhuma ancora
        EncontraColunaSubc = i
    End If
    
End Function
'

Function EncontraLinhaAncora() As Integer
    'Encontra na planilha de pesos
    IDEmpresa = Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 1)
    EncontraLinhaAncora = 3
        While (Sheets("Pesos").Cells(EncontraLinhaAncora, 1) <> IDEmpresa And EncontraLinhaAncora <= Sheets("Pesos").Range("A1").End(xlDown).Row)
            EncontraLinhaAncora = EncontraLinhaAncora + 1
        Wend
        If (EncontraLinhaAncora = Sheets("Pesos").Range("A1").End(xlDown).Row + 1) Then
            EncontraLinhaAncora = 0
        End If
End Function


Sub PreencherNotas()

    'Encotra a linha correta pelo ID da ancora, se nao existir o ID nao fazer nada
        LinhaAncora = EncontraLinhaAncora
        
        
        For i = 3 To Sheets("Notas").Range("A1").End(xlDown).Row
            
            If (Sheets("Notas").Cells(i, 1) = Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 1) And _
                Sheets("Notas").Cells(i, 2) = Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 1)) Then 'Se a linha for da ancora e fornecedor
                
                Linha = i
                
                'Se o criterio tiver peso e nota entao grava no optionbutton
                    
                    Coluna = 3
                    While (Sheets("Notas").Cells(1, Coluna) <> "")
                        ID = Sheets("Notas").Cells(1, Coluna)
                        Csubc = EncontraColunaSubc(ID)
                        If (Sheets("Pesos").Cells(LinhaAncora, Csubc + 1) > 0 And Sheets("Notas").Cells(Linha, Coluna) > 0) Then
                            ValorNota = Sheets("Notas").Cells(Linha, Coluna)
                            Sheets("Avalia��o do fornecedor").OptionButtons(ID & "O" & ValorNota).Value = True
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
        While (Sheets("Notas").Cells(LinhaNota, 1) <> Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 1) Or _
                Sheets("Notas").Cells(LinhaNota, 2) <> Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 1))
                
            If (Sheets("Notas").Cells(LinhaNota, 1) <> "") Then
                LinhaNota = LinhaNota + 1
            Else 'N�o foi criado ainda
                
                Sheets("Notas").Cells(LinhaNota, 1) = Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 1)
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
                    If (Sheets("Avalia��o do fornecedor").OptionButtons(ID & "O" & cont).Value = 1) Then
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
        VoltarAvaliar
    Else
    MsgBox "� obrigatorio a avalia��o de todos os crit�rios!"
    End If
End Sub

Sub LimparNota()
'Pos = 100
    For i = 1 To Sheets("Avalia��o do fornecedor").GroupBoxes.Count
        Sheets("Avalia��o do fornecedor").GroupBoxes.Delete
    Next i
    For i = 1 To Sheets("Avalia��o do fornecedor").OptionButtons.Count
        Sheets("Avalia��o do fornecedor").OptionButtons.Delete
    Next i
    For i = 1 To Sheets("Avalia��o do fornecedor").TextBoxes.Count
        Sheets("Avalia��o do fornecedor").TextBoxes.Delete
    Next i
    For i = 10 To Pos
        Sheets("Avalia��o do fornecedor").Cells(i, 2) = ""
        Sheets("Avalia��o do fornecedor").Cells(i, 3) = ""
            Sheets("Avalia��o do fornecedor").Cells(i, 2).Font.Bold = False
            Sheets("Avalia��o do fornecedor").Cells(i, 2).Font.Italic = False
            Sheets("Avalia��o do fornecedor").Cells(i, 2).Font.Underline = False
            Sheets("Avalia��o do fornecedor").Cells(i, 2).Font.Size = 11
            Sheets("Avalia��o do fornecedor").Cells(i, 2).Font.Color = RGB(0, 0, 0)
            Sheets("Avalia��o do fornecedor").Cells(i, 3).Font.Bold = False
            Sheets("Avalia��o do fornecedor").Cells(i, 3).Font.Italic = False
            Sheets("Avalia��o do fornecedor").Cells(i, 3).Font.Underline = False
            Sheets("Avalia��o do fornecedor").Cells(i, 3).Font.Size = 11
            Sheets("Avalia��o do fornecedor").Cells(i, 3).Font.Color = RGB(0, 0, 0)
            Sheets("Avalia��o do fornecedor").Cells(i, 2).HorizontalAlignment = xlLeft
        'Sem preenchimento
        Sheets("Avalia��o do fornecedor").Range(Cells(i, 2), Cells(i, 3)).Interior.Pattern = xlNone
        Sheets("Avalia��o do fornecedor").Cells(i, 3).Borders(xlEdgeLeft).LineStyle = xlNone
        Sheets("Avalia��o do fornecedor").Cells(i, 4).Borders(xlEdgeLeft).LineStyle = xlNone
        'Desmesclar celulas
        Sheets("Avalia��o do fornecedor").Range(Cells(i, 2), Cells(i, 3)).UnMerge
    Next i
    Sheets("Avalia��o do fornecedor").Shapes("VoltarMenu").Top = Sheets("Avalia��o do fornecedor").Cells(10, 2).Top
    Sheets("Avalia��o do fornecedor").Shapes("Salvar").Top = Sheets("Avalia��o do fornecedor").Cells(10, 2).Top
    
    
End Sub

Sub VoltarAvaliar()
    LimparNota
    Av1.Show
    
End Sub


Function Preenchido() As Boolean
    'Verifica se ao menos um dos optionbuttons de cada grupo foi escolhido

    Preenchido = True
    For Grupo = 1 To Sheets("Avalia��o do fornecedor").OptionButtons.Count / 5
        'Se todos desmarcados entao preenchido = false (-4146)
        If (Sheets("Avalia��o do fornecedor").OptionButtons((Grupo - 1) * 5 + 1) = -4146 And _
         Sheets("Avalia��o do fornecedor").OptionButtons((Grupo - 1) * 5 + 2) = -4146 And _
         Sheets("Avalia��o do fornecedor").OptionButtons((Grupo - 1) * 5 + 3) = -4146 And _
         Sheets("Avalia��o do fornecedor").OptionButtons((Grupo - 1) * 5 + 4) = -4146 And _
         Sheets("Avalia��o do fornecedor").OptionButtons((Grupo - 1) * 5 + 5) = -4146) Then
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
Sub Formata��ocriterioN(Pos As Integer)
        
    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Italic = True 'italico
    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Underline = False
    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Font.Size = 14
 

    Sheets("Avalia��o do fornecedor").Cells(Pos, 2).Select
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 180
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
        With Selection.Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Interior.Color = RGB(255, 255, 255)
End Sub

Function VerificaPesos(Linha As Integer) As Boolean

    VerificaPesos = False
    C = 2
    While ((Sheets("Pesos").Cells(Linha, C) <> "") And (VerificaPesos = False))
        If (Sheets("Pesos").Cells(Linha, C) = 0) Then
            C = C + 1
        Else
            VerificaPesos = True
        End If
    Wend
    
End Function



