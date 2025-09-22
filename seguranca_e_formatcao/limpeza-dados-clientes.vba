Option Explicit

'Rotina para limpeza de dados dos clientes
Sub sbLimpaDados()
    
    'Declaração de variável
    Dim lContador As Long

    'Inicializa a variável de linha
    lContador = 2
    
    'Cria uma cópia da planilha
    ActiveSheet.Copy After:=Sheets(1)
    ActiveSheet.Name = "Revisada-" & Format(Now(), "HH-mm-ss")
    
    'Repetição para cada uma das linhas da planilha
    Do While Trim(Cells(lContador, 1)) <> vbNullString
    
        'COLUNA A: Ajustando o ID do cliente
        If Left(Cells(lContador, 1), 5) <> "byte_" Then
            Cells(lContador, 1) = "byte_" & Cells(lContador, 1)
        End If
        
        'COLUNA B: Limpando caracteres estranhos no nome do cliente
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "#", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "$", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "*", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "%", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "&", "")
            
        'COLUNA C: Ajustando o valor moeda
        Cells(lContador, 3) = Replace(Cells(lContador, 3), "R$", "")
        Cells(lContador, 3) = Replace(Cells(lContador, 3), ",", "")
        Cells(lContador, 3) = Replace(Cells(lContador, 3), ".", ",")
        Cells(lContador, 3).NumberFormat = "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"
        
        'COLUNA D: Criando o e-mail interno do cliente
        Cells(lContador, 4) = Cells(lContador, 1) & "@bytebank.com.br"
        
        lContador = lContador + 1
    Loop
    
    'Formata como tabela
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$D$10"), , xlYes).Name = _
        "Tabela1"
    
    
End Sub