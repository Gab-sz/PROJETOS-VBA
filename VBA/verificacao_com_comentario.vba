' =================================================================================================
' EVENTO PRINCIPAL: Executa sempre que uma célula da planilha for alterada.
' É como um "assistente robô" que fica vigiando a planilha.
' =================================================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- Bloco de Tratamento de Erros ---
    ' Se qualquer erro inesperado acontecer, o código pula para a linha "TratarErro".
    ' É o "guarda-costas" do código, evitando que o Excel trave.
    On Error GoTo TratarErro

    ' --- Variáveis ---
    ' São como "caixas" para guardar informações que usaremos no código.
    Dim celula As Range
    Dim numeroPR As String
    Dim pastaRaiz As String
    Dim fso As Object
    Dim arquivoEncontrado As Boolean
    Dim contemCredito As Boolean
    Dim subpastasPermitidas As Variant
    Dim subpasta As Variant
    Dim linha As Long

    ' --- PROTEÇÃO ANTI-TRAVAMENTO ---
    ' Impede loops ao excluir linhas/colunas. Se mais de 1000 células mudam, ele para.
    If Target.Cells.CountLarge > 1000 Then
        Debug.Print "AVISO: Alteração massiva detectada. Execução interrompida para evitar travamento."
        Exit Sub
    End If

    ' --- 1. VERIFICAÇÃO INICIAL (O GATILHO) ---
    ' O robô só continua se a célula alterada estiver na coluna "I" (Crédito) ou "J" (Nº da PR).
    If Not Intersect(Target, Union(Me.Columns("I"), Me.Columns("J"))) Is Nothing Then
        ' Desativa os eventos do Excel temporariamente para evitar que o robô entre em um loop infinito.
        Application.EnableEvents = False

        ' --- 2. CONFIGURAÇÃO DA BUSCA (A MISSÃO) ---
        pastaRaiz = Environ("USERPROFILE") & "\MerckGroup\ORCAMENTOS - General\"
        Set fso = CreateObject("Scripting.FileSystemObject")

        ' Verifica se a pasta principal existe. Se não, avisa o usuário e para.
        If Not fso.FolderExists(pastaRaiz) Then
            MsgBox "ERRO: A pasta principal de orçamentos não foi encontrada. Verifique o caminho e seu acesso.", vbCritical, "Pasta Não Localizada"
            GoTo Fim
        End If

        subpastasPermitidas = Array("2 - OT - DESPESA", "3 - CAPEX - PROJETOS NOVOS")

        ' --- 3. PROCESSAMENTO DAS CÉLULAS ALTERADAS (MÃOS À OBRA) ---
        For Each celula In Target
            linha = celula.Row
            numeroPR = CStr(Trim(Me.Cells(linha, "J").Value))

            Debug.Print "-------------------------------------------------------"
            Debug.Print "Processando Linha " & linha & " | PR: '" & numeroPR & "'"

            ' Limpa a formatação antiga ANTES de qualquer nova lógica.
            ' .ColorIndex = 0 restaura a cor automática da Tabela (linhas alternadas).
            Me.Cells(linha, "I").Interior.ColorIndex = 0
            Me.Cells(linha, "J").Interior.ColorIndex = 0

            ' Só faz a busca se a célula da PR não estiver vazia.
            If numeroPR <> "" Then
                arquivoEncontrado = False
                contemCredito = False

                ' Loop para procurar nas subpastas permitidas.
                For Each subpasta In subpastasPermitidas
                    If BuscarArquivoComCredito(fso.GetFolder(pastaRaiz & subpasta), numeroPR, contemCredito) Then
                        arquivoEncontrado = True
                        Exit For
                    End If
                Next subpasta

                ' --- 4. ATUALIZAÇÃO DA PLANILHA (O FEEDBACK VISUAL) ---
                ' Agora, o robô colore as células para mostrar o resultado.
                If arquivoEncontrado Then
                    ' Pinta de amarelo (sucesso)
                    Me.Cells(linha, "J").Interior.Color = RGB(255, 242, 204)
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 242, 204)
                    Debug.Print "-> RESULTADO: Arquivo encontrado."

                    ' Se o nome do arquivo continha "crédito" (Verificação Inteligente)...
                    If contemCredito Then
                        Me.Cells(linha, "I").Value = "X" ' ...ele marca a coluna "I" com "X" sozinho.
                        Debug.Print "-> DETALHE: 'crédito' encontrado. Coluna I marcada com 'X'."
                    End If
                Else
                    ' Pinta de vermelho (erro)
                    Me.Cells(linha, "J").Interior.Color = RGB(255, 99, 71)
                    Debug.Print "-> RESULTADO: Arquivo NÃO encontrado."
                End If

                ' --- 5. VALIDAÇÃO DE ERRO HUMANO (O "POLICIAL DE ERROS") ---
                ' Esta parte verifica se o usuário cometeu algum engano.
                Dim valorColunaI As String
                valorColunaI = UCase(Me.Cells(linha, "I").Value)

                ' CASO 1: Usuário marcou "X", mas o arquivo NÃO foi encontrado.
                If valorColunaI = "X" And Not arquivoEncontrado Then
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 99, 71)
                    Debug.Print "-> ALERTA: 'X' manual, mas arquivo não localizado."
                End If

                ' CASO 2: Usuário marcou "X", o arquivo FOI encontrado, mas NÃO contém "crédito".
                If valorColunaI = "X" And arquivoEncontrado And Not contemCredito Then
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 99, 71)
                    Me.Cells(linha, "J").Interior.Color = RGB(255, 99, 71)
                    Debug.Print "-> ALERTA: 'X' manual, mas arquivo não contém 'crédito'."
                End If

            Else
                ' Se a célula da PR ficou vazia, o código só precisa garantir que a marcação "X" seja limpa.
                ' A cor já foi resetada no início do loop.
                If UCase(Me.Cells(linha, "I").Value) = "X" Then
                    Me.Cells(linha, "I").ClearContents
                End If
                Debug.Print "-> IGNORADO: Célula da PR vazia."
            End If
        Next celula
    End If

' --- Ponto Final da Execução Normal ---
Fim:
    ' Reativa os eventos do Excel. ESSENCIAL para a planilha voltar ao normal.
    Application.EnableEvents = True
    Exit Sub

' --- Bloco de Tratamento de Erros (O "GUARDA-COSTAS") ---
' O código só chega aqui se um erro inesperado ocorreu.
TratarErro:
    Debug.Print "!!! ERRO INESPERADO: " & Err.Description & " !!!"
    MsgBox "Ocorreu um erro inesperado na automação.", vbCritical, "Erro na Macro"
    Resume Fim
End Sub

' =================================================================================================
' FUNÇÃO AUXILIAR 1: Busca o arquivo recursivamente.
' "Recursivamente" significa que ela procura em uma pasta e em todas as subpastas dentro dela.
' =================================================================================================
Private Function BuscarArquivoComCredito(pasta As Object, prefixo As String, ByRef contemCredito As Boolean) As Boolean
    Dim arquivo As Object
    Dim subpasta As Object
    Dim nomeSubpasta As String
    Dim ano As Long
    Dim nomeArquivoSemExtensao As String

    ' 1. Procura nos arquivos da pasta atual.
    For Each arquivo In pasta.Files
        nomeArquivoSemExtensao = Left(arquivo.Name, InStrRev(arquivo.Name, ".") - 1)

        If VerificarCodigoEmNome(nomeArquivoSemExtensao, prefixo) Then
            ' Verifica se o nome do arquivo contém "crédito"
            If InStr(1, nomeArquivoSemExtensao, "crédito", vbTextCompare) > 0 Then
                contemCredito = True
            End If
            BuscarArquivoComCredito = True
            Exit Function
        End If
    Next arquivo

    ' 2. Se não achou, procura nas subpastas.
    For Each subpasta In pasta.SubFolders
        nomeSubpasta = Trim(subpasta.Name)

        ' REGRA DE NEGÓCIO: Ignora pastas com nome de ano anterior a 2025 para otimizar a busca.
        If IsNumeric(nomeSubpasta) Then
            ano = CLng(nomeSubpasta)
            If ano < 2025 Then
                GoTo ProximaSubpasta ' Pula para a próxima subpasta.
            End If
        End If

        ' Chama a si mesma (recursão) para buscar dentro da subpasta.
        If BuscarArquivoComCredito(subpasta, prefixo, contemCredito) Then
            BuscarArquivoComCredito = True
            Exit Function
        End If

ProximaSubpasta:
    Next subpasta

    BuscarArquivoComCredito = False
End Function

' =================================================================================================
' FUNÇÃO AUXILIAR 2: Usa Expressão Regular (Regex) para uma busca inteligente.
' Garante que a PR "123" seja encontrada em "PR-123", mas não em "ABC-51234".
' =================================================================================================
Private Function VerificarCodigoEmNome(nomeArquivo As String, codigo As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Define o padrão de busca para encontrar o código como uma palavra isolada.
    re.Pattern = "(^|[\s\-])" & codigo & "($|[\s\-])"
    re.IgnoreCase = True ' Não diferencia maiúsculas de minúsculas.
    re.Global = False    ' Para na primeira vez que encontrar.

    VerificarCodigoEmNome = re.Test(nomeArquivo)
End Function
