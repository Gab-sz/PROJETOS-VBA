' Evento que dispara quando qualquer célula da planilha é alterada
Private Sub Worksheet_Change(ByVal Target As Range)
    Debug.Print "Evento Worksheet_Change ativado." & vbNewLine

    Dim celula As Range
    Dim valorBusca As String
    Dim pastaRaiz As String
    Dim fso As Object
    Dim encontrado As Boolean
    Dim contemCredito As Boolean
    Dim subpastasPermitidas As Variant
    Dim subpasta As Variant
    Dim linha As Long

    ' Verifica se a alteração ocorreu nas colunas H (crédito) ou I (nº da PR)
    If Not Intersect(Target, Union(Me.Columns("H"), Me.Columns("I"))) Is Nothing Then
        Debug.Print "Alteração detectada nas colunas H ou I." & vbNewLine

        Application.EnableEvents = False ' Impede que o código cause recursividade

        ' Define a pasta raiz onde os arquivos estão armazenados (OneDrive corporativo)
        pastaRaiz = Environ("USERPROFILE") & "\tkinGroup\ORCAMENTOS - General\"
        Set fso = CreateObject("Scripting.FileSystemObject")

        ' Define as subpastas autorizadas para busca
        subpastasPermitidas = Array("2 - OT - DESPESA", "3 - CAPEX - PROJETOS NOVOS")

        ' Se a pasta raiz não existir, aborta a execução
        If Not fso.FolderExists(pastaRaiz) Then
            MsgBox "A pasta especificada não existe: " & pastaRaiz, vbExclamation
            Debug.Print "ERRO: Pasta raiz não encontrada: " & pastaRaiz & vbNewLine
            GoTo Fim
        End If

        ' Percorre todas as células afetadas na alteração
        For Each celula In Intersect(Target, Union(Me.Columns("H"), Me.Columns("I")))
            linha = celula.Row
            valorBusca = CStr(Trim(Me.Cells(linha, "I").Value)) ' Extrai valor da PR na linha correspondente

            Debug.Print "Linha: " & linha & " | PR informada: '" & valorBusca & "'"

            ' Ignora linhas sem PR
            If valorBusca <> "" Then
                encontrado = False
                contemCredito = False

                ' Busca o arquivo nas subpastas permitidas
                For Each subpasta In subpastasPermitidas
                    Debug.Print "Procurando na subpasta: " & subpasta
                    If BuscarArquivoComCredito(fso.GetFolder(pastaRaiz & subpasta), valorBusca, contemCredito) Then
                        encontrado = True
                        Debug.Print "Arquivo encontrado em: " & subpasta
                        Exit For
                    End If
                Next subpasta

                ' Se arquivo for encontrado:
                If encontrado Then
                    ' Colore as células de amarelo (arquivo encontrado)
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 242, 204)
                    Me.Cells(linha, "H").Interior.Color = RGB(255, 242, 204)
                    Debug.Print "STATUS: Arquivo correspondente à PR foi localizado."

                    ' Se o nome do arquivo contiver "crédito", marca automaticamente com "X"
                    If contemCredito Then
                        Me.Cells(linha, "H").Value = "X"
                        Debug.Print "AÇÃO: 'Crédito' detectado no nome do arquivo. Coluna H marcada com 'X'."
                    Else
                        Debug.Print "INFO: Arquivo não contém a palavra 'crédito'. Nenhuma marcação na Coluna H."
                    End If
                Else
                    ' Arquivo não encontrado: destaca a PR de vermelho
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 99, 71)
                    Debug.Print "ERRO: Arquivo da PR não encontrado em nenhuma subpasta."

                    ' Se houver um X manual, sinaliza erro com vermelho na coluna H
                    If UCase(Me.Cells(linha, "H").Value) = "X" Then
                        Me.Cells(linha, "H").Interior.Color = RGB(255, 99, 71)
                        Debug.Print "ERRO: X manual na Coluna H, mas arquivo da PR não localizado."
                    End If
                End If

                ' Caso o arquivo exista, mas não contenha "crédito" e o usuário tiver marcado "X"
                If UCase(Me.Cells(linha, "H").Value) = "X" And encontrado And Not contemCredito Then
                    Me.Cells(linha, "H").Interior.Color = RGB(255, 99, 71)
                    Me.Cells(linha, "I").Interior.Color = RGB(255, 99, 71)
                    Debug.Print "ERRO: X marcado manualmente, mas arquivo localizado NÃO contém a palavra 'crédito'."
                End If
            Else
                Debug.Print "IGNORADO: Linha " & linha & " vazia na coluna I."
            End If

            Debug.Print vbNewLine ' Separador visual no log
        Next celula

Fim:
        Application.EnableEvents = True ' Restaura o evento após a execução
        Debug.Print "Execução do evento Worksheet_Change finalizada." & vbNewLine
    End If
End Sub

' Função que busca arquivos por PR e identifica se contêm "crédito" no nome
Private Function BuscarArquivoComCredito(pasta As Object, prefixo As String, ByRef contemCredito As Boolean) As Boolean
    Dim arquivo As Object
    Dim subpasta As Object
    Dim nomeSubpasta As String
    Dim ano As Long
    Dim nomeArquivoSemExtensao As String

    Debug.Print "Iniciando busca na pasta: " & pasta.Path

    ' Verifica todos os arquivos da pasta atual
    For Each arquivo In pasta.Files
        nomeArquivoSemExtensao = Left(arquivo.Name, InStrRev(arquivo.Name, ".") - 1)

        If VerificarCodigoEmNome(nomeArquivoSemExtensao, prefixo) Then
            Debug.Print "MATCH: Arquivo com PR encontrada -> " & arquivo.Name

            ' Verifica se o nome do arquivo contém "crédito"
            If InStr(1, nomeArquivoSemExtensao, "crédito", vbTextCompare) > 0 Then
                contemCredito = True
                Debug.Print "CONFIRMADO: Nome do arquivo contém a palavra 'crédito'."
            Else
                Debug.Print "ALERTA: Nome do arquivo NÃO contém a palavra 'crédito'."
            End If

            Debug.Print vbNewLine
            BuscarArquivoComCredito = True
            Exit Function
        End If
    Next arquivo

    ' Busca recursiva em subpastas (ignora pastas com nomes de anos anteriores a 2025)
    For Each subpasta In pasta.SubFolders
        nomeSubpasta = Trim(subpasta.Name)

        If IsNumeric(nomeSubpasta) Then
            ano = CLng(nomeSubpasta)
            If ano < 2025 Then
                Debug.Print "Subpasta ignorada por ser anterior a 2025: " & subpasta.Path
                GoTo ProximaSubpasta
            End If
        End If

        ' Chamada recursiva
        If BuscarArquivoComCredito(subpasta, prefixo, contemCredito) Then
            BuscarArquivoComCredito = True
            Exit Function
        End If

ProximaSubpasta:
    Next subpasta

    Debug.Print "Nenhum arquivo com a PR foi encontrado nesta pasta: " & pasta.Path
    Debug.Print vbNewLine
    BuscarArquivoComCredito = False
End Function

' Função que usa expressão regular para verificar se a PR está isolada no nome do arquivo
Private Function VerificarCodigoEmNome(nomeArquivo As String, codigo As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Regex: confere se a PR aparece como palavra isolada (com espaço, hífen ou underline)
    re.Pattern = "(^|[\s\-_])" & codigo & "($|[\s\-_])"
    re.IgnoreCase = True
    re.Global = False

    VerificarCodigoEmNome = re.Test(nomeArquivo)
    If VerificarCodigoEmNome Then
        Debug.Print "Regex PASSOU: Código isolado '" & codigo & "' encontrado em '" & nomeArquivo & "'"
        Debug.Print vbNewLine
    End If
End Function
