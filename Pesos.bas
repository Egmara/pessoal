Attribute VB_Name = "Pesos"
Dim Pos As Integer 'Posi��o a ser colocado o criterio/subcriteiro

Dim LinhaAncora As Integer

Sub EscolherCriterios()
    'Crit � a posi��o do criteiro na lista de criterios
    If (ExisteSubcriterios) Then 'Existe ao menos um subcriterios
        If (Sheets("�ncoras").Range("A1").End(xlDown).Row - 2 > 0) Then
            UserFormAncora.Show
            If (EmpresaEscolhida <> 0) Then
            Sheets("Escolha dos crit�rios").Select
        
                'Inicio Impacto financeiro
                Sheets("Escolha dos crit�rios").Range("B6") = "Impacto Financeiro"
            
                'formata��o da linha do nome da empresa ***
                Sheets("Escolha dos crit�rios").Range("B4") = "Empresa avaliadora: " & Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 2) 'Nome da Ancora
                Sheets("Escolha dos crit�rios").Cells(4, 2).Interior.Color = RGB(91, 155, 213)
                Sheets("Escolha dos crit�rios").Cells(4, 3).Interior.Color = RGB(91, 155, 213)
            
                    ContR = 0 'Posi��o inicial da listagem temporaria de posi�oes dos criteriso de risco de fornecimento
                    Pos = 10
                
                
                    'Separa em I ou R, se I grava na planilha, se R guarda no vetor
                        For crite = 1 To Sheets("Crit�rios").Range("A1").End(xlDown).Row - 2 'i-esimo criterio
                            If (Sheets("Crit�rios").Cells(crite + 2, 4) = "I") Then
                                Call EscreverCriterios(Pos, crite)
                            Else
                                ContR = ContR + 1
                                Sheets("Escolha dos crit�rios").Cells(ContR, 1) = crite
                            End If
                        Next crite
                
                    
                'Inicio Risco de fornecimento
                        If (Pos <> 10) Then 'h� criteiros de impacto financeiro
                        'formata��o do titulo "Risco de fornecimento"
                            Pos0R = Pos
                            Pos = Pos + 1
                
                        'formata��o do titulo "Risco de fornecimento"
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2) = "Risco de fornecimento"
                            Sheets("Escolha dos crit�rios").Range(Cells(Pos, 2), Cells(Pos, 3)).Merge
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).HorizontalAlignment = xlCenter
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).VerticalAlignment = xlCenter
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Interior.Color = RGB(117, 113, 113)
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Color = RGB(242, 242, 242)
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Size = 12
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Bold = True
            
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 2) = "Crit�rios"
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 2).Font.Bold = True
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 2).Font.Size = 12
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 3) = "Pesos"
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 3).Font.Bold = True
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 3).Font.Size = 12
                        
                            Pos = Pos + 4
                        Else
                            Sheets("Escolha dos crit�rios").Range("B6") = "Risco de fornecimento"
                        End If
                
                    For i = 1 To ContR
                        crite = Sheets("Escolha dos crit�rios").Cells(i, 1)
                        Sheets("Escolha dos crit�rios").Cells(i, 1) = ""
                        Call EscreverCriterios(Pos, crite)
                               
                    Next i
                    
                    If (Pos0R + 5 = Pos) Then 'N�o h� pesos em nenhum subcriterio do eixo R
                        Pos = Pos - 4
                        Sheets("Escolha dos crit�rios").Cells(Pos, 2) = ""
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).HorizontalAlignment = xlLeft
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).VerticalAlignment = xlCenter
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Interior.Pattern = xlNone
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Color = RGB(0, 0, 0)
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Size = 12
                            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Bold = False
                            Sheets("Escolha dos crit�rios").Range(Cells(Pos, 2), Cells(Pos, 3)).UnMerge
            
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 2) = ""
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 2).Font.Bold = False
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 2).Font.Size = 12
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 3) = ""
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 3).Font.Bold = False
                            Sheets("Escolha dos crit�rios").Cells(Pos + 2, 3).Font.Size = 12
                            Pos = Pos - 1
                    End If
                'bota no lugar os botoes
                    Sheets("Escolha dos crit�rios").Shapes("VoltarMenu").Top = Sheets("Escolha dos crit�rios").Cells(Pos, 2).Top
                    Sheets("Escolha dos crit�rios").Shapes("Salvar").Top = Sheets("Escolha dos crit�rios").Cells(Pos, 2).Top
                
                'Preenche pesos ja preenchidos anteriormente
                    PreencherPeso
                Sheets("Escolha dos crit�rios").Range("B2").Select
            End If
        
        Else
            MsgBox "N�o h� empresas cadastradas!"
        End If
    Else
        MsgBox "N�o h� crit�rios cadastrados!"
    End If
    
End Sub
'
'
Function ExisteSubcriterios() As Boolean
    Linha = 3
    ExisteSubcriterios = False
    While (Sheets("Crit�rios").Cells(Linha, 1) <> "" And ExisteSubcriterios = False)
        If (Sheets("Crit�rios").Cells(Linha, 7) = "") Then
            Linha = Linha + 1
        Else
            ExisteSubcriterios = True
        End If
    Wend
End Function
Sub EscreverCriterios(ByRef Pos As Integer, ByVal crite As Integer)
    Sheets("Escolha dos crit�rios").Select
    
    'Adiciona o criterio e formata��o
    Sheets("Escolha dos crit�rios").Cells(Pos, 2) = Sheets("Crit�rios").Cells(crite + 2, 2) 'nome do criterio
        
    'Formata o nome criterio
    Formata��ocriterioP (Pos)
    
        'Insere caixa de texto da descri��o do crit�rio
            Call InserirCaixaDesc(Pos)
            
            Sheets("Escolha dos crit�rios").TextBoxes("D" & Pos).Text = Sheets("Crit�rios").Cells(crite + 2, 3) 'Descri��o do criterio
        
        Pos = Pos + 5

    'Adiciona Subcriterios
        QntSub = 0 ' quantidade de sub no criterio
        While (Sheets("Crit�rios").Cells(crite + 2, QntSub + 7) <> "")
            QntSub = QntSub + 1
            ID = Sheets("Crit�rios").Cells(crite + 2, QntSub + 6)
            Subc = EncontraSubc(Sheets("Crit�rios").Cells(crite + 2, QntSub + 6)) 'i-esimo sub da lista de subcriterios
            
            'Insere caixa de texto da descri��o do subcrit�rio
                Call InserirCaixaDesc(Pos)
                Sheets("Escolha dos crit�rios").TextBoxes("D" & Pos).Text = Sheets("Subcrit�rios").Cells(Subc + 2, 3) 'Descri��o do subcriterio
            
            'nome do subcriterio e formata��o
            Sheets("Escolha dos crit�rios").Cells(Pos, 2) = "Subcrit�rio " & QntSub & ": " & Sheets("Subcrit�rios").Cells(Subc + 2, 2)
            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Bold = False
            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Italic = True
            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Underline = True
            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Size = 11
            Sheets("Escolha dos crit�rios").Cells(Pos, 2).Interior.Color = RGB(189, 215, 238)
            Sheets("Escolha dos crit�rios").Cells(Pos, 3).Interior.Color = RGB(189, 215, 238)
            
            Sheets("Escolha dos crit�rios").Cells(Pos, 3).Borders(xlEdgeLeft).Weight = xlThick
            Sheets("Escolha dos crit�rios").Cells(Pos, 4).Borders(xlEdgeLeft).Weight = xlThick
            Sheets("Escolha dos crit�rios").Cells(Pos, 3).Borders(xlEdgeLeft).Color = RGB(221, 235, 247)
            Sheets("Escolha dos crit�rios").Cells(Pos, 4).Borders(xlEdgeLeft).Color = RGB(221, 235, 247)
                         
            'Insere comboboxes
                T = Sheets("Escolha dos crit�rios").Cells(Pos, 3).Top + 2.0454
                E = Sheets("Escolha dos crit�rios").Cells(Pos, 3).Left + 3.4091
                W = 55.29551
                H = 15
                 
                Sheets("Escolha dos crit�rios").DropDowns.Add(E, T, W, H).Select
                Selection.Name = ID
                Sheets("Escolha dos crit�rios").Shapes(ID).ControlFormat.AddItem Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
            
            Pos = Pos + 5
        Wend
        Pos = Pos + 1 ' linha em branco para separar criterios
        
       
    
End Sub
Sub InserirCaixaDesc(ByRef Pos As Integer)
            'Insere caixa de texto da descri��o do crit�rio
        
        T = Sheets("Escolha dos crit�rios").Cells(Pos + 1, 2).Top + 10.5
        E = Sheets("Escolha dos crit�rios").Cells(Pos + 1, 2).Left + 3.75
        W = 589.5
        H = 60.75
            
        Sheets("Escolha dos crit�rios").TextBoxes.Add(E, T, W, H).Select
        Selection.Name = "D" & Pos
        
        'formata��o do fundo da caixa
        
        
        'formata��o de tras da caixa
        For Linha = Pos + 1 To Pos + 4
            Sheets("Escolha dos crit�rios").Cells(Linha, 2).Interior.Color = RGB(221, 235, 247)
            Sheets("Escolha dos crit�rios").Cells(Linha, 3).Interior.Color = RGB(221, 235, 247)
        Next
        
End Sub
Function EncontraSubc(ID As String) As Integer
    EncontraSubc = 0
    While (Sheets("Subcrit�rios").Cells(EncontraSubc + 2, 1) <> ID)
        EncontraSubc = EncontraSubc + 1
    Wend
End Function

Sub SalvarPeso()

        IDEmpresa = Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 1)
        Linha = 3
        While (Sheets("Pesos").Cells(Linha, 1) <> IDEmpresa)
                    If (Sheets("Pesos").Cells(Linha, 1) = "") Then
                        Sheets("Pesos").Cells(Linha, 1) = IDEmpresa
                        Vazioigual0AncP (Linha)
                    Else
                        Linha = Linha + 1
                    End If
        
        Wend
    'Olhando todos os combobox, se tiver valor, procure pra ve se ja tem o criterio listado e caso nao tenha liste-o, em seguida salve o valor
        For i = 1 To Sheets("Escolha dos crit�rios").DropDowns.Count
            If Sheets("Escolha dos crit�rios").DropDowns(i) <> 0 Then
                Coluna = 2
                While (Sheets("Pesos").Cells(1, Coluna) <> Sheets("Escolha dos crit�rios").DropDowns(i).Name)
                    If (Sheets("Pesos").Cells(1, Coluna) = "") Then
                        Sheets("Pesos").Cells(1, Coluna) = Sheets("Escolha dos crit�rios").DropDowns(i).Name
                        Vazioigual0CritP (Coluna)
                    Else
                        Coluna = Coluna + 1
                    End If
                Wend
            
            
                Sheets("Pesos").Cells(Linha, Coluna) = Sheets("Escolha dos crit�rios").DropDowns(i).Value - 1 'pois o dropdown trabalha com o indice nao o valor em si
                If Sheets("Pesos").Cells(Linha, Coluna) = -1 Then
                    Sheets("Pesos").Cells(Linha, Coluna) = 0
                End If
                    
            End If
        Next i
        MsgBox "Os pesos foram salvos!"
        VoltarEscolhaCriterios
        
End Sub
Sub PreencherPeso()

    'Encotra a linha correta pelo ID da ancora, se nao existir o ID nao fazer nada
        
        For i = 3 To Sheets("Pesos").Range("A1").End(xlDown).Row
            If (Sheets("Pesos").Cells(i, 1) = Sheets("�ncoras").Cells(EmpresaEscolhida + 2, 1)) Then
                Linha = i
                'Se o criterio tiver peso entao grava no combobox
                    Coluna = 2
                    While (Sheets("Pesos").Cells(1, Coluna) <> "")
                        ID = Sheets("Pesos").Cells(1, Coluna)
                        If (Sheets("Pesos").Cells(Linha, Coluna) > 0) Then
                            Sheets("Escolha dos crit�rios").DropDowns(ID).Value = Sheets("Pesos").Cells(Linha, Coluna) + 1
                        End If
                        Coluna = Coluna + 1
                    Wend
            End If
        Next i

End Sub

Sub LimparPeso()
Pos = 100 'descomentar para limpar planilha caso aja bug
    For i = 1 To Sheets("Escolha dos crit�rios").DropDowns.Count
        Sheets("Escolha dos crit�rios").DropDowns.Delete
    Next i
    For i = 1 To Sheets("Escolha dos crit�rios").TextBoxes.Count
        Sheets("Escolha dos crit�rios").TextBoxes.Delete
    Next i
    For i = 10 To Pos
        Sheets("Escolha dos crit�rios").Cells(i, 2) = ""
        Sheets("Escolha dos crit�rios").Cells(i, 3) = ""
     
            Sheets("Escolha dos crit�rios").Cells(i, 2).Font.Bold = False
            Sheets("Escolha dos crit�rios").Cells(i, 2).Font.Italic = False
            Sheets("Escolha dos crit�rios").Cells(i, 2).Font.Underline = False
            Sheets("Escolha dos crit�rios").Cells(i, 2).Font.Size = 11
            Sheets("Escolha dos crit�rios").Cells(i, 2).Font.Color = RGB(0, 0, 0)
            Sheets("Escolha dos crit�rios").Cells(i, 3).Font.Bold = False
            Sheets("Escolha dos crit�rios").Cells(i, 3).Font.Italic = False
            Sheets("Escolha dos crit�rios").Cells(i, 3).Font.Underline = False
            Sheets("Escolha dos crit�rios").Cells(i, 3).Font.Size = 11
            Sheets("Escolha dos crit�rios").Cells(i, 3).Font.Color = RGB(0, 0, 0)
            Sheets("Escolha dos crit�rios").Cells(i, 2).HorizontalAlignment = xlLeft
        'Sem preenchimento
        Sheets("Escolha dos crit�rios").Range(Cells(i, 2), Cells(i, 3)).Interior.Pattern = xlNone
        Sheets("Escolha dos crit�rios").Cells(i, 3).Borders(xlEdgeLeft).LineStyle = xlNone
        Sheets("Escolha dos crit�rios").Cells(i, 4).Borders(xlEdgeLeft).LineStyle = xlNone
        'Desmesclar celulas
        Sheets("Escolha dos crit�rios").Range(Cells(i, 2), Cells(i, 3)).UnMerge
    Next i
    Sheets("Escolha dos crit�rios").Shapes("VoltarMenu").Top = Sheets("Escolha dos crit�rios").Cells(10, 2).Top
    Sheets("Escolha dos crit�rios").Shapes("Salvar").Top = Sheets("Escolha dos crit�rios").Cells(10, 2).Top
    
    
End Sub

Sub VoltarEscolhaCriterios()
    LimparPeso
    Sheets("Menu temporario").Select
    
End Sub

Sub Vazioigual0CritP(Coluna As Integer) 'por zero aodne teria vazio ao avaliar novo criteiro (peso)
    Linha = 3
    While (Sheets("Pesos").Cells(Linha, 1) <> "")
        If (Sheets("Pesos").Cells(Linha, Coluna) = "") Then
            Sheets("Pesos").Cells(Linha, Coluna) = 0
        End If
        Linha = Linha + 1
    Wend
End Sub
Sub Vazioigual0AncP(Linha As Integer) 'por zero aodne teria vazio ao avaliar nova ancora (peso)
    Coluna = 2
    While (Sheets("Pesos").Cells(1, Coluna) <> "")
        If (Sheets("Pesos").Cells(Linha, Coluna) = "") Then
            Sheets("Pesos").Cells(Linha, Coluna) = 0
        End If
        Coluna = Coluna + 1
    Wend
  
    
End Sub
Sub Formata��ocriterioP(Pos As Integer)
        
    Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Italic = True 'italico
    Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Underline = False
    Sheets("Escolha dos crit�rios").Cells(Pos, 2).Font.Size = 14
 

    Sheets("Escolha dos crit�rios").Cells(Pos, 2).Select
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
    Sheets("Escolha dos crit�rios").Cells(Pos, 3).Interior.Color = RGB(255, 255, 255)
End Sub
'
