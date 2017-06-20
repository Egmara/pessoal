Attribute VB_Name = "Notas"
Dim Pos As Integer 'Posi��o a ser colocado o criterio

Dim LinhaAncora As Integer
Public Sub AvaliarFornecedor()

Application.ScreenUpdating = False


    If Sheets("�ncoras").Range("A1").End(xlDown).Row - 2 > 0 Then 'verifica se h� ancoras
        UserFormAncora.Show 'EmpresaEscolhida
        
        If EmpresaEscolhida <> 0 Then
        
            If (EncontraLinhaAncora <> 0) Then
                'verificva se ha pelo menos um peso <> de 0
                If (VerificaPesos(EncontraLinhaAncora)) Then
                
                    If Sheets("Fornecedores").Range("A1").End(xlDown).Row - 2 > 0 Then 'verifica se h� forncedores
                        UserFormFornecedor.Show 'FornecedorEscolhido
            
                        If FornecedorEscolhido <> 0 Then
                            Sheets("Avalia��o do fornecedor").Select
                            LimparNota
                            Application.ScreenUpdating = False
                            'insere nome do fornecedor
                            Sheets("Avalia��o do fornecedor").Range("B5") = "EMPRESA AVALIADA: " & Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 2)
                    
                            '---------------
                            'Inicio Impacto financeiro
                            Sheets("Avalia��o do fornecedor").Range("B7") = "IMPACTO FINANCEIRO"
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
                
                                
                                If (Pos <> 11) Then 'h� criteiros de impacto financeiro
                                    'formata��o do titulo "Risco de fornecimento"
                                    Pos0R = Pos 'Posi��o onde come�o a trabalhar com "risco de fornecimento"
                                    Pos = Pos + 1
                            
                                    Sheets("Avalia��o do fornecedor").Cells(Pos, 2) = "RISCO DE FORNECIMENTO"
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
                                    Sheets("Avalia��o do fornecedor").Range("B7") = "RISCO DE FORNECIMENTO"
                                End If
                        
                                For i = 1 To ContR
                                    crite = Sheets("Avalia��o do fornecedor").Cells(i, 1)
                                    Sheets("Avalia��o do fornecedor").Cells(i, 1) = ""
                                    Call EscreverCriterios(Pos, crite)
                               
                                Next i
                                
                                If (Pos0R + 5 = Pos) Then 'N�o h� pesos em nenhum criterio do eixo R
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
                            Application.ScreenUpdating = True
                            Sheets("Avalia��o do fornecedor").Range("B3").Select
                    
                        
                                
                        End If
                    
                    Else
                        FornecedorEscolhido = 0
                        MsgBox "N�o h� empresas fornecedoras cadastradas!"
                    End If 'do quinto if
                
            Else
                EmpresaEscolhida = 0
                MsgBox "A empresa ainda n�o deu peso aos crit�rios!"

            End If 'do quarto If
        Else
            EmpresaEscolhida = 0
             MsgBox "A empresa ainda n�o deu peso aos crit�rios!"
        End If 'do terceiro If
        
    End If 'do segundo if
        
    Else 'do primeiro If
        EmpresaEscolhida = 0
        MsgBox "N�o h� empresas cadastradas!"
        
    End If
   
   
End Sub
Sub EscreverCriterios(ByRef Pos As Integer, ByVal crite As Integer)
    'Adiciona o criterio. Verifica no criterio foi dado peso . Caso sim grava, caso o contrario pula para o proximo.
    
    Sheets("Avalia��o do fornecedor").Select
    
    LinhaAncora = EncontraLinhaAncora
    ID = Sheets("Crit�rios").Cells(crite + 2, 1)
    ColunaCrite = 2
    While (Sheets("Pesos").Cells(1, ColunaCrite) <> ID And Sheets("Pesos").Cells(1, ColunaCrite) <> "")
        ColunaCrite = ColunaCrite + 1
    Wend
    
    If (Sheets("Pesos").Cells(LinhaAncora, ColunaCrite) > 0 And Sheets("Pesos").Cells(1, ColunaCrite) <> "") Then
        'Adiciona o criterio e formata��o
    Sheets("Avalia��o do fornecedor").Cells(Pos, 2) = Sheets("Crit�rios").Cells(crite + 2, 2) 'nome do criterio
    'Formata o nome do criterio
        Call Formata��ocriterioN(Pos, crite) 'Insere caixa de texto da descri��o do crit�rio
        Call InserirCaixaDesc(Pos, crite)
            
            Sheets("Avalia��o do fornecedor").TextBoxes("D" & Pos).Text = Sheets("Crit�rios").Cells(crite + 2, 3) 'Descri��o do criterio
        
   
        
     
                    'Insere caixas de grupo
                        L = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Left
                        T = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Top
                        W = Sheets("Avalia��o do fornecedor").Cells(Pos, 4).Left - L
                        h = Sheets("Avalia��o do fornecedor").Cells(Pos + 1, 3).Top - T

                        Sheets("Avalia��o do fornecedor").GroupBoxes.Add(L, T, W, h).Select
                        Selection.Name = ID
                        Selection.Visible = False
                        Selection.Caption = ""
                    
                    'Insere os optionsbuttons
                        For cont = 1 To 5
                            T = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Top
                            E = Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Left
                            W = 15
                            h = 15
                            Sheets("Avalia��o do fornecedor").OptionButtons.Add(E + cont * 25, T, W, h).Select
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
            'Insere caixa de texto da descri��o do crit�rio
        
        T = Sheets("Avalia��o do fornecedor").Cells(Pos + 1, 2).Top + 10.5
        E = Sheets("Avalia��o do fornecedor").Cells(Pos + 1, 2).Left + 3.75
        W = 636
        h = 60.75
            
        Sheets("Avalia��o do fornecedor").TextBoxes.Add(E, T, W, h).Select
        Selection.Name = "D" & Pos
        
        'formata��o do fundo da caixa
       'Selection.Interior.Color = RGB(217, 225, 242)
        
        'formata��o de tras da caixa
        If (Sheets("Crit�rios").Cells(crite + 2, 7) = "I") Then
            For Linha = Pos + 1 To Pos + 4
                Sheets("Avalia��o do fornecedor").Cells(Linha, 2).Interior.Color = RGB(169, 208, 142)
                Sheets("Avalia��o do fornecedor").Cells(Linha, 3).Interior.Color = RGB(169, 208, 142)
            Next
        End If
        If (Sheets("Crit�rios").Cells(crite + 2, 7) = "P") Then
        For Linha = Pos + 1 To Pos + 4
                Sheets("Avalia��o do fornecedor").Cells(Linha, 2).Interior.Color = RGB(221, 235, 247)
                Sheets("Avalia��o do fornecedor").Cells(Linha, 3).Interior.Color = RGB(221, 235, 247)
            Next
        End If
End Sub
'


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
            
            If (Sheets("Notas").Cells(i, 1) = Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 1) And _
                Sheets("Notas").Cells(i, 2) = Sheets("Fornecedores").Cells(FornecedorEscolhido + 2, 1)) Then 'Se a linha for da ancora e fornecedor
                
                Linha = i
                
                'Se o criterio tiver peso e nota entao grava no optionbutton
                    
                    Coluna = 3
                    While (Sheets("Notas").Cells(1, Coluna) <> "")
                        ID = Sheets("Notas").Cells(1, Coluna)
                        ColunaPeso = EncontraColuna(ID)
                        If (Sheets("Pesos").Cells(LinhaAncora, ColunaPeso) > 0 And Sheets("Notas").Cells(Linha, Coluna) > 0) Then
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
        'LimparNota
        ' em vez de VoltarAvaliar ( que � retornar ao menu) pergunta se quer avaliar outro fornecedor
        ' n�o tem isso ainda Av1.Show
        VoltarAvaliar
        
    Else
    MsgBox "� obrigatorio a avalia��o de todos os crit�rios!"
    End If
End Sub

Sub LimparNota()

Application.ScreenUpdating = False
Application.DisplayAlerts = False


    For i = 1 To Sheets("Avalia��o do fornecedor").GroupBoxes.Count
        Sheets("Avalia��o do fornecedor").GroupBoxes.Delete
    Next i
    For i = 1 To Sheets("Avalia��o do fornecedor").OptionButtons.Count
        Sheets("Avalia��o do fornecedor").OptionButtons.Delete
    Next i
    For i = 1 To Sheets("Avalia��o do fornecedor").TextBoxes.Count
        Sheets("Avalia��o do fornecedor").TextBoxes.Delete
    Next i
    
    Sheets("Avalia��o do fornecedor").Select
    For i = 11 To Sheets("Avalia��o do fornecedor").Shapes("VoltarMenu").Top

        'Desmesclar celulas
        Sheets("Avalia��o do fornecedor").Range(Cells(i, 2), Cells(i, 3)).Merge
        Sheets("Avalia��o do fornecedor").Range(Cells(i, 2), Cells(i, 3)).UnMerge
    Next i
    Plan21.Range("B9:C5000").ClearFormats
    Plan21.Range("B9:C5000").ClearContents
    Sheets("Avalia��o do fornecedor").Shapes("VoltarMenu").Top = Sheets("Avalia��o do fornecedor").Cells(11, 2).Top
    Sheets("Avalia��o do fornecedor").Shapes("Salvar").Top = Sheets("Avalia��o do fornecedor").Cells(11, 2).Top
    
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
Sub Formata��ocriterioN(Pos As Integer, ByVal crite As Integer)
        
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

    'Se inverso
    If (Sheets("Crit�rios").Cells(crite + 2, 7) = "I") Then
        With Selection.Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        End With
    End If
       'Se proporcional
    If (Sheets("Crit�rios").Cells(crite + 2, 7) = "P") Then
        With Selection.Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        End With
    End If
    
    Sheets("Avalia��o do fornecedor").Cells(Pos, 3).Interior.Color = RGB(255, 255, 255)
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



