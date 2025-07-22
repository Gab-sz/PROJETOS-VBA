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
    Dim celula As Range             ' Guarda a célula que foi modificada.
    Dim numeroPR As String          ' Guarda o número da PR digitado.
    Dim pastaRaiz As String         ' Guarda o caminho da pasta principal onde os arquivos estão.
    Dim fso As Object               ' Objeto para manipular pastas e arquivos do sistema.
    Dim arquivoEncontrado As Boolean ' Verdadeiro ou Falso: indica se o arquivo da PR foi achado.
    Dim contemCredito As Boolean    ' Verdadeiro ou Falso: indica se a palavra "crédito" está no nome do arquivo.
    Dim subpastasPermitidas As Variant ' Lista de pastas onde a busca é autorizada.
    Dim subpasta As Variant         ' Guarda o nome de cada subpasta durante o loop.
    Dim linha As Long               ' Guarda o número da linha que está sendo processada.

    Debug.Print "======================================================="
    Debug.Print "INÍCIO: Evento Worksheet_Change disparado."
    Debug.Print "Célula(s) alterada(s): " & Target.Address

    ' --- 1. VERIFICAÇÃO INICIAL (O GATILHO) ---
    ' O robô só continua se a célula alterada estiver na coluna "H" (Crédito) ou "I" (Nº da PR).
    If Not Intersect(Target, Union(Me.Columns("H"), Me.Columns("I"))) Is Nothing Then
        Debug.Print "-> Gatilho VÁLIDO: Alteração detectada nas colunas H ou I."

        ' Desativa os eventos do Excel temporariamente para evitar que o robô entre em um loop infinito.
        Application.EnableEvents = False
        Debug.Print "-> AÇÃO: Eventos do Excel DESATIVADOS para evitar loops."

        ' --- 2. CONFIGURAÇÃO DA BUSCA ---
        ' Define o caminho da pasta principal, funcionando em diferentes computadores.
        pastaRaiz = Environ("USERPROFILE") & "\tkinGroup\ORCAMENTOS - General\"
        Set fso = CreateObject("Scripting.FileSystemObject") ' Cria a ferramenta para mexer com pastas.
        Debug.Print "-> INFO: Pasta raiz definida como: " & pastaRaiz

        ' Define exatamente em quais subpastas o robô deve procurar.
        subpastasPermitidas = Array("2 - OT - DESPESA", "3 - CAPEX - PROJETOS NOVOS")
        Debug.Print "-> INFO: Subpastas permitidas para busca: " & Join(subpastasPermitidas, ", ")

        ' Verifica se a pasta principal existe. Se não, avisa o usuário e para.
        If Not fso.FolderExists(pastaRaiz) Then
            Debug.Print "-> ERRO CRÍTICO: A pasta raiz não existe. Abortando execução."
            MsgBox "ERRO: A pasta principal de orçamentos não foi encontrada. Verifique se o caminho está correto e se você tem acesso." & vbNewLine & vbNewLine & "Caminho procurado: " & pastaRaiz, vbCritical, "Pasta Não Localizada"
            GoTo Fim ' Pula para o final do código.
        End If

        ' --- 3. PROCESSAMENTO DAS CÉLULAS ALTERADAS  ---
        ' Garante que o robô verifique cada uma das células que foram alteradas de uma vez.
        For Each celula In Target
            linha = celula.Row ' Pega o número da linha da célula atual.
            numeroPR = CStr(Trim(Me.Cells(linha, "I").Value)) ' Pega o valor da PR na coluna I.

            Debug.Print "-------------------------------------------------------"
            Debug.Print "Processando Linha " & linha & " | Célula alterada: " & celula.Address

            ' Só faz a busca se a célula da PR não estiver vazia.
            If numeroPR <> "" Then
                Debug.Print "-> VALOR ENCONTRADO: PR '" & numeroPR & "' na coluna I. Iniciando busca..."
                arquivoEncontrado = False ' Reseta as "caixas" para a nova busca.
                contemCredito = False

                ' Loop para procurar nas subpastas permitidas.
                For Each subpasta In subpastasPermitidas
                    Debug.Print "--> Buscando na subpasta: '" & subpasta & "'"
                    ' Chama a função "BuscarArquivo" para fazer a busca.
                    If BuscarArquivo(fso.GetFolder(pastaRaiz & subpasta), numeroPR, contemCredito) Then
                        Debug.Print "--> SUCESSO: Arquivo da PR '" & numeroPR & "' encontrado nesta pasta."
                        arquivoEncontrado = True ' Marca que achou...
                        Exit For                 ' ...e para de procurar.
                    Else
                        Debug.Print "--> INFO: Arquivo da PR '" & numeroPR & "' não encontrado nesta subpasta."
                    End If
                Next subpasta

                ' --- 4. ATUALIZAÇÃO DA PLANILHA (FEEDBACK VISUAL) ---
                ' Agora, o robô colore as células para mostrar o resultado.
                If arquivoEncontrado Then
                    Debug.Print "-> RESULTADO: ARQUIVO ENCONTRADO para a PR '" & numeroPR & "'."
                    ' Pinta as células de amarelo para indicar "Sucesso".
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 242, 204) ' Amarelo claro
                    Me.Cells(linha, "H").Interior.Color = RGB(255, 242, 204)
                    Debug.Print "-> AÇÃO: Células H" & linha & " e I" & linha & " coloridas de AMARELO (Sucesso)."

                    ' Se o nome do arquivo continha "crédito" (Verificação Inteligente)...
                    If contemCredito Then
                        Debug.Print "-> DETALHE: A palavra 'crédito' foi encontrada no nome do arquivo."
                        Me.Cells(linha, "H").Value = "X" ' ...ele marca a coluna "H" com "X" sozinho.
                        Debug.Print "-> AÇÃO: Coluna H" & linha & " marcada com 'X' automaticamente."
                    Else
                        Debug.Print "-> DETALHE: A palavra 'crédito' NÃO foi encontrada no nome do arquivo."
                    End If
                Else
                    Debug.Print "-> RESULTADO: ARQUIVO NÃO ENCONTRADO para a PR '" & numeroPR & "'."
                    ' Pinta a célula da PR de vermelho para indicar "Erro".
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 99, 71) ' Vermelho tomate
                    Debug.Print "-> AÇÃO: Célula I" & linha & " colorida de VERMELHO (Erro)."
                End If

                ' --- 5. VALIDAÇÃO DE ERRO HUMANO ("POLICIAL DE ERROS") ---
                ' Esta parte verifica se o usuário cometeu algum engano.
                Dim valorColunaH As String
                valorColunaH = UCase(Me.Cells(linha, "H").Value)

                ' CASO 1: Usuário marcou "X", mas o arquivo NÃO foi encontrado.
                If valorColunaH = "X" And Not arquivoEncontrado Then
                    Me.Cells(linha, "H").Interior.Color = RGB(255, 99, 71)
                    Debug.Print "-> ALERTA DE INCONSISTÊNCIA: 'X' manual na coluna H, mas o arquivo da PR não foi localizado. Célula H" & linha & " colorida de VERMELHO."
                End If

                ' CASO 2: Usuário marcou "X", o arquivo FOI encontrado, mas NÃO contém "crédito".
                If valorColunaH = "X" And arquivoEncontrado And Not contemCredito Then
                    Me.Cells(linha, "H").Interior.Color = RGB(255, 99, 71)
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 99, 71)
                    Debug.Print "-> ALERTA DE INCONSISTÊNCIA: 'X' manual, arquivo localizado, mas o nome NÃO contém 'crédito'. Células H" & linha & " e I" & linha & " coloridas de VERMELHO."
                End If

            Else
                ' Se a célula da PR estiver vazia, limpa a formatação da linha.
                Debug.Print "-> IGNORADO: Célula da PR na linha " & linha & " está vazia. Limpando formatação."
                Me.Cells(linha, "I").Interior.ColorIndex = xlNone ' Sem cor
                Me.Cells(linha, "H").Interior.ColorIndex = xlNone ' Sem cor
            End If
        Next celula
    Else
        Debug.Print "-> Gatilho IGNORADO: A alteração não ocorreu nas colunas H ou I."
    End If

' --- Ponto Final da Execução Normal ---
Fim:
    ' Reativa os eventos do Excel. ESSENCIAL para a planilha voltar ao normal.
    Application.EnableEvents = True
    Debug.Print "-> AÇÃO: Eventos do Excel REATIVADOS."
    Debug.Print "FIM: Execução do evento finalizada."
    Debug.Print "=======================================================" & vbNewLine
    Exit Sub ' Encerra o código.

' --- Bloco de Tratamento de Erros ("GUARDA-COSTAS") ---
' O código só chega aqui se um erro inesperado ocorreu.
TratarErro:
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Debug.Print "-> ERRO INESPERADO CAPTURADO!"
    Debug.Print "-> Descrição do Erro: " & Err.Description
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    ' Exibe uma caixa de mensagem amigável com a descrição do erro.
    MsgBox "Ocorreu um erro inesperado na automação." & vbNewLine & vbNewLine & _
           "Erro: " & Err.Description & vbNewLine & vbNewLine & _
           "Por favor, contate o suporte ou tente novamente.", vbCritical, "Erro na Macro"
    ' Mesmo após um erro, pula para a linha "Fim" para garantir que os eventos sejam reativados.
    Resume Fim
End Sub


' =================================================================================================
' FUNÇÃO AUXILIAR 1: Busca o arquivo recursivamente.
' "Recursivamente" significa que ela procura em uma pasta e em todas as subpastas dentro dela.
' =================================================================================================
Private Function BuscarArquivo(pasta As Object, ByVal prefixo As String, ByRef contemCredito As Boolean) As Boolean
    Dim arquivo As Object
    Dim subpasta As Object
    Dim nomeSubpasta As String
    Dim ano As Long
    Dim nomeArquivoSemExtensao As String

    ' 1. Procura nos arquivos da pasta atual.
    For Each arquivo In pasta.Files
        nomeArquivoSemExtensao = Left(arquivo.Name, InStrRev(arquivo.Name, ".") - 1)

        If VerificarCodigoNoNome(nomeArquivoSemExtensao, prefixo) Then
            Debug.Print "    [Busca] MATCH ENCONTRADO! Arquivo: '" & arquivo.Name & "'"
            If InStr(1, nomeArquivoSemExtensao, "crédito", vbTextCompare) > 0 Then
                contemCredito = True
                Debug.Print "    [Busca] CONFIRMADO: Nome do arquivo contém 'crédito'."
            Else
                Debug.Print "    [Busca] INFO: Nome do arquivo NÃO contém 'crédito'."
            End If
            BuscarArquivo = True
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
                Debug.Print "    [Busca] IGNORANDO pasta de ano antigo: '" & subpasta.Path & "'"
                GoTo ProximaSubpasta ' Pula para a próxima subpasta.
            End If
        End If

        ' Chama a si mesma (recursão) para buscar dentro da subpasta.
        If BuscarArquivo(subpasta, prefixo, contemCredito) Then
            BuscarArquivo = True
            Exit Function
        End If
ProximaSubpasta:
    Next subpasta

    BuscarArquivo = False
End Function


' =================================================================================================
' FUNÇÃO AUXILIAR 2: Usa Expressão Regular (Regex) para uma busca inteligente.
' Garante que a PR "123" seja encontrada em "PR-123", mas não em "ABC-51234".
' =================================================================================================
Private Function VerificarCodigoNoNome(ByVal nomeArquivo As String, ByVal codigo As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Define o padrão de busca para encontrar o código como uma palavra isolada.
    re.Pattern = "(^|[\s\-_])" & codigo & "($|[\s\-_])"
    re.IgnoreCase = True ' Não diferencia maiúsculas de minúsculas.
    re.Global = False    ' Para na primeira vez que encontrar.

    VerificarCodigoNoNome = re.Test(nomeArquivo)
    
    If VerificarCodigoNoNome Then
        Debug.Print "      [Regex] PASSOU: Código '" & codigo & "' encontrado de forma isolada em '" & nomeArquivo & "'"
    End If
End Function
```
