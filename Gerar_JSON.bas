Attribute VB_Name = "Gerar_JSON"
Sub GerarJSON()
    Dim wsExp As Worksheet, wsHab As Worksheet, wsForm As Worksheet
    Dim wsCert As Worksheet, wsVaga As Worksheet, wsAest As Worksheet, wsAav As Worksheet
    Dim json, config, contexto, agentes, empresas, habilidades, formacao, certificacoes As Object
    Dim situacao, atividadeDict, empresaDict As Object
    Dim ultimaLinha, ultimaLinha_1, i As Integer
    Dim jsonString, filePath As String
    Dim idiomaEscolhido As String
    Dim perfilOtimizacao As String
    Dim resposta As VbMsgBoxResult
    Dim respostaPerfil As VbMsgBoxResult
    
    ' ========== SELEÇÃO DE IDIOMA ==========
    resposta = MsgBox("Escolha o idioma de saída:" & vbCrLf & vbCrLf & _
                      "Sim = Português" & vbCrLf & _
                      "Não = Inglês", _
                      vbQuestion + vbYesNo, "Idioma de Saída")
    
    If resposta = vbYes Then
        idiomaEscolhido = "pt-BR"
    Else
        idiomaEscolhido = "en-US"
    End If
    
    ' ========== SELEÇÃO DE PERFIL DE OTIMIZAÇÃO ==========
    Dim msgPerfil As String
    msgPerfil = "Escolha o perfil de otimização dos bullets:" & vbCrLf & vbCrLf & _
                "============================" & vbCrLf & _
                "SIM = ATS-Friendly" & vbCrLf & _
                "   • Foco em palavras-chave (40%)" & vbCrLf & _
                "   • Ideal para passar por sistemas automáticos" & vbCrLf & vbCrLf & _
                "NÃO = Balanceado" & vbCrLf & _
                "   • Equilíbrio entre especificidade (40%) e keywords (20%)" & vbCrLf & _
                "   • Recomendado para a maioria dos casos" & vbCrLf & vbCrLf & _
                "CANCELAR = Premium" & vbCrLf & _
                "   • Máxima diferenciação (50% especificidade)" & vbCrLf & _
                "   • Ideal para posições sênior/executivas"
    
    respostaPerfil = MsgBox(msgPerfil, vbYesNoCancel + vbQuestion, "Perfil de Otimização")
    
    Select Case respostaPerfil
        Case vbYes
            perfilOtimizacao = "ats_friendly"
        Case vbNo
            perfilOtimizacao = "balanceado"
        Case vbCancel
            perfilOtimizacao = "premium"
        Case Else
            perfilOtimizacao = "balanceado"
    End Select
      
    ' Definir planilhas
    Set wsExp = ThisWorkbook.Sheets("Experiencias")
    Set wsHab = ThisWorkbook.Sheets("Habilidades")
    Set wsForm = ThisWorkbook.Sheets("Formacao")
    Set wsCert = ThisWorkbook.Sheets("Certificacoes")
    Set wsVaga = ThisWorkbook.Sheets("VagaPretendida")
    Set wsAest = ThisWorkbook.Sheets("Aest")
    Set wsAav = ThisWorkbook.Sheets("Aav")
    
    ' ========== ESTRUTURA RAIZ DO JSON ==========
    Set json = CreateObject("Scripting.Dictionary")
    
    ' README para compatibilidade universal com outras IAs
    json("readme") = "INSTRUCOES DE PROCESSAMENTO: Este JSON contem dados estruturados para analise de compatibilidade candidato-vaga. " & _
                     "FLUXO: (1) Processe os agentes na ordem definida em 'config.ordem_execucao'. " & _
                     "(2) Agentes com 'output_destination=chat_response' devem ter seu output exibido na conversa. " & _
                     "DEPENDENCIAS: Respeite o campo 'depende_de' para executar agentes apos suas dependencias."
    
    ' Schema customizado (identificador de versão e formato)
    json("schema") = "analise-vaga-candidato-v1.0"
    json("version") = "1.0.0"
    
    ' Configurações globais
    Set config = CreateObject("Scripting.Dictionary")
    config("idioma_saida") = idiomaEscolhido
    config("processar_agentes_em_ordem") = True
    config("ordem_execucao") = Array("especialista_storytelling_curriculo", "sugestao_melhoria_storytelling", "analista_compatibilidade_vaga")
    config("perfil_otimizacao") = perfilOtimizacao
    Set json("config") = config
    
    ' ========== CONTEXTO (Dados do candidato e vaga) ==========
    Set contexto = CreateObject("Scripting.Dictionary")
    
    ' Vaga Pretendida
    Dim vagaPretendida As Object
    Dim requitosArray As Variant
    Dim competenciaArray As Variant
    
    Set vagaPretendida = CreateObject("Scripting.Dictionary")
    vagaPretendida("vaga_cargo") = wsVaga.Cells(2, 1).Value
    
    ' Gerar arrays separadamente
    requitosArray = GerarArrayColuna(wsVaga, 2)
    competenciaArray = GerarArrayColuna(wsVaga, 3)
    
    ' Atribuir aos dicionários SEM Set
    vagaPretendida("vaga_requisitos") = requitosArray
    vagaPretendida("vaga_responsabilidades") = competenciaArray
    
    Set contexto("vaga_pretendida") = vagaPretendida
    
    ' Profissional
    Dim profissional As Object
    Dim historicoArray As Variant
    Dim formacaoArray As Variant
    Dim certificacoesArray As Variant
    
    Set profissional = CreateObject("Scripting.Dictionary")
    
    ' Gerar arrays separadamente
    historicoArray = GerarHistoricoProfissional(wsExp)
    formacaoArray = GerarFormacao(wsForm)
    certificacoesArray = GerarCertificacoes(wsCert)
    
    ' Atribuir SEM Set (são arrays)
    profissional("candidato_experiencia") = historicoArray
    
    ' Habilidades
    Set habilidades = CreateObject("Scripting.Dictionary")
    habilidades("hard_skills") = Split(wsHab.Range("A2").Value, ", ")
    habilidades("soft_skills") = Split(wsHab.Range("B2").Value, ", ")
    Set profissional("candidato_competencias") = habilidades
    
    ' Formação e Certificações
    profissional("candidato_formacao") = formacaoArray
    profissional("candidato_certificacoes") = certificacoesArray
    
    Set contexto("profissional") = profissional
    
    Set json("contexto") = contexto
    
    ' ========== AGENTES (Instruções de processamento) ==========
    Set agentes = CreateObject("Scripting.Dictionary")
    
    ' Agente 1: Especialista Storytelling
    Set agentes("especialista_storytelling_curriculo") = GerarAgenteStorytelling(wsAest, perfilOtimizacao)
    
    ' Agente 2: Agente de Melhoria Contínua
    Set agentes("sugestao_melhoria_storytelling") = GerarAgenteSugestaoMelhoria()
    
    ' Agente 3: Analista Compatibilidade
    Set agentes("analista_compatibilidade_vaga") = GerarAgenteCompatibilidade(wsAav)
    
    Set json("agentes") = agentes
    
    ' ========== CONVERSÃO E SALVAMENTO ==========
    jsonString = JsonConverter.ConvertToJson(json, Whitespace:=2)
    jsonString = RemoveUnicodeEscapes(jsonString)
    
    ' Salvar como UTF-8 sem BOM
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    filePath = ThisWorkbook.Path & "\historico_profissional.json"
    
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText jsonString
        .SaveToFile filePath, 2
        .Close
    End With
    
    ' Mensagem personalizada com o perfil escolhido
    Dim msgPerfilEscolhido As String
    Select Case perfilOtimizacao
        Case "ats_friendly"
            msgPerfilEscolhido = "ATS-Friendly (40% keywords)"
        Case "balanceado"
            msgPerfilEscolhido = "Balanceado (40% especificidade, 20% keywords)"
        Case "premium"
            msgPerfilEscolhido = "Premium (50% especificidade)"
    End Select
    
    MsgBox "JSON gerado com sucesso!" & vbCrLf & vbCrLf & _
           "Perfil: " & msgPerfilEscolhido & vbCrLf & _
           "Arquivo: " & filePath, vbInformation, "Sucesso!"
End Sub

Private Function GerarAgenteStorytelling(ws As Worksheet, perfilOtimizacao As String) As Object
    Dim agente, regras As Object
    Set agente = CreateObject("Scripting.Dictionary")
    
    agente("prioridade") = 1
    agente("output_format") = "bullets"
    agente("output_destination") = "chat_response"
    agente("tarefa") = ws.Cells(2, 1).Value
    
    ' Definir pesos baseado no perfil
    Dim pesoEspecifico As Integer, pesoGenerico As Integer, pesoKeywords As Integer
    
    Select Case perfilOtimizacao
        Case "ats_friendly"
            pesoEspecifico = 30
            pesoGenerico = 30
            pesoKeywords = 40
        Case "balanceado"
            pesoEspecifico = 40
            pesoGenerico = 40
            pesoKeywords = 20
        Case "premium"
            pesoEspecifico = 50
            pesoGenerico = 30
            pesoKeywords = 20
        Case Else
            pesoEspecifico = 40
            pesoGenerico = 40
            pesoKeywords = 20
    End Select
    
    ' Regras agrupadas
    Set regras = CreateObject("Scripting.Dictionary")
    Dim criterios As Variant
    Dim exemplosArray As Variant
    
    criterios = GerarCriteriosAgrupados(ws, 2)
    exemplosArray = GerarArrayColuna(ws, 4)
    
    ' Extrair e organizar por tipo de critério
    Dim i As Integer
    For i = 0 To UBound(criterios)
        Dim criterioNome As String
        criterioNome = criterios(i)("criterio")
        
        Dim regraObj As Object
        Set regraObj = CreateObject("Scripting.Dictionary")
        regraObj("instrucoes") = criterios(i)("instrucoes")
        
        ' Normalizar nome do critério e atribuir peso
        Dim criterioNormalizado As String
        Select Case criterioNome
            Case "Criação de bullets Especifico", "Criacao de bullets Especifico"
                criterioNormalizado = "criacao_bullets_especifico"
                regraObj("peso") = pesoEspecifico
            Case "Criação de bullets Genérico", "Criacao de bullets Generico"
                criterioNormalizado = "criacao_bullets_generico"
                regraObj("peso") = pesoGenerico
            Case "Palavras-Chaves", "Palavras-chaves"
                criterioNormalizado = "palavras_chave"
                regraObj("peso") = pesoKeywords
            Case Else
                ' Normalização genérica: remove acentos, lowercase, espaços->underscore
                criterioNormalizado = NormalizarNome(criterioNome)
                regraObj("peso") = 20
        End Select
        
        Set regras(criterioNormalizado) = regraObj
    Next i
    
    Set agente("regras") = regras
    agente("exemplos_saida") = exemplosArray
    
    Set GerarAgenteStorytelling = agente
End Function

Private Function GerarAgenteCompatibilidade(ws As Worksheet) As Object
    Dim agente, criterios As Object
    Set agente = CreateObject("Scripting.Dictionary")
    
    agente("prioridade") = 2
    agente("depende_de") = Array("especialista_storytelling_curriculo")
    agente("output_destination") = "chat_response"
    agente("tarefa") = ws.Cells(2, 1).Value
    
    ' Critérios com pesos
    Set criterios = CreateObject("Scripting.Dictionary")
    Dim criteriosArray As Variant
    criteriosArray = GerarCriteriosAgrupados(ws, 2)
    
    Dim i As Integer
    For i = 0 To UBound(criteriosArray)
        Dim criterioObj As Object
        Set criterioObj = CreateObject("Scripting.Dictionary")
        criterioObj("instrucoes") = criteriosArray(i)("instrucoes")
        
        Dim criterioOriginal As String
        Dim criterioNormalizado As String
        criterioOriginal = criteriosArray(i)("criterio")
        
        ' Atribuir pesos baseado no critério original
        Select Case True
            Case InStr(criterioOriginal, "Pontos Fortes") > 0 Or InStr(criterioOriginal, "pontos fortes") > 0
                criterioNormalizado = "pontos_fortes"
                criterioObj("peso") = 30
            Case InStr(criterioOriginal, "Lacunas") > 0 Or InStr(criterioOriginal, "lacunas") > 0
                criterioNormalizado = "lacunas"
                criterioObj("peso") = 25
            Case InStr(criterioOriginal, "Pontuação") > 0 Or InStr(criterioOriginal, "pontuacao") > 0 Or InStr(criterioOriginal, "Compatibilidade") > 0
                criterioNormalizado = "pontuacao_compatibilidade"
                criterioObj("peso") = 20
            Case InStr(criterioOriginal, "Recomendações") > 0 Or InStr(criterioOriginal, "recomendacoes") > 0
                criterioNormalizado = "recomendacoes"
                criterioObj("peso") = 25
            Case Else
                criterioNormalizado = NormalizarNome(criterioOriginal)
                criterioObj("peso") = 20
        End Select
        
        Set criterios(criterioNormalizado) = criterioObj
    Next i
    
    Set agente("criterios_avaliacao") = criterios
    
    ' Formato de saída estruturado
    Dim formatoSaida As Object
    Dim recomendacoes As Object
    
    Set formatoSaida = CreateObject("Scripting.Dictionary")
    formatoSaida("pontos_fortes") = "array"
    formatoSaida("lacunas") = "array"
    formatoSaida("pontuacao") = "number"
    formatoSaida("justificativa") = "string"
    
    Set recomendacoes = CreateObject("Scripting.Dictionary")
    recomendacoes("curto_prazo") = "array"
    recomendacoes("medio_prazo") = "array"
    recomendacoes("longo_prazo") = "array"
    
    Set formatoSaida("recomendacoes") = recomendacoes
    Set agente("formato_saida_estruturado") = formatoSaida
    
    Set GerarAgenteCompatibilidade = agente
End Function

Private Function GerarCriteriosAgrupados(ws As Worksheet, colCriterio As Integer) As Variant
    Dim criterioList() As Object
    Dim criterioIndex As Integer
    Dim ultimaLinha, i As Integer
    Dim currentCriterio As String
    Dim currentInstrucoes() As String
    Dim instrucaoIndex As Integer
    Dim criterioObj As Object
    Dim lastCriterio As String
    
    criterioIndex = -1
    ultimaLinha = ws.Cells(Rows.Count, colCriterio).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If ws.Cells(i, colCriterio).Value <> "" Then
            If ws.Cells(i, colCriterio).Value <> lastCriterio Then
                If lastCriterio <> "" Then
                    Set criterioObj = CreateObject("Scripting.Dictionary")
                    criterioObj("criterio") = lastCriterio
                    criterioObj("instrucoes") = currentInstrucoes
                    criterioIndex = criterioIndex + 1
                    ReDim Preserve criterioList(criterioIndex)
                    Set criterioList(criterioIndex) = criterioObj
                End If
                
                lastCriterio = ws.Cells(i, colCriterio).Value
                ReDim currentInstrucoes(0)
                instrucaoIndex = 0
            End If
        End If
        
        If ws.Cells(i, colCriterio + 1).Value <> "" Then
            If instrucaoIndex > 0 Then
                ReDim Preserve currentInstrucoes(instrucaoIndex)
            End If
            currentInstrucoes(instrucaoIndex) = ws.Cells(i, colCriterio + 1).Value
            instrucaoIndex = instrucaoIndex + 1
        End If
        
        If i = ultimaLinha Then
            Set criterioObj = CreateObject("Scripting.Dictionary")
            criterioObj("criterio") = lastCriterio
            criterioObj("instrucoes") = currentInstrucoes
            criterioIndex = criterioIndex + 1
            ReDim Preserve criterioList(criterioIndex)
            Set criterioList(criterioIndex) = criterioObj
        End If
    Next i
    
    GerarCriteriosAgrupados = criterioList
End Function

Private Function GerarArrayColuna(ws As Worksheet, coluna As Integer) As Variant
    Dim lista() As String
    Dim ultimaLinha, i, Index As Integer
    
    Index = 0
    ultimaLinha = ws.Cells(Rows.Count, coluna).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If Index = 0 Then
            ReDim lista(0)
        Else
            ReDim Preserve lista(Index)
        End If
        lista(Index) = ws.Cells(i, coluna).Value
        Index = Index + 1
    Next i
    
    GerarArrayColuna = lista
End Function

Private Function ParsearPeriodo(ByVal periodo As String) As Object
    Dim periodoObj As Object
    Dim partes() As String
    Dim mesAnoInicio() As String
    Dim mesAnoFim() As String
    Dim dataInicio As Date
    Dim dataFim As Date
    Dim duracaoMeses As Integer
    
    Set periodoObj = CreateObject("Scripting.Dictionary")
    
    ' Verificar se tem "Atual" ou é um período completo
    If InStr(periodo, "Atual") > 0 Or InStr(periodo, "atual") > 0 Then
        partes = Split(periodo, " - ")
        mesAnoInicio = Split(Trim(partes(0)), "/")
        
        periodoObj("inicio") = mesAnoInicio(1) & "-" & Format(mesAnoInicio(0), "00")
        periodoObj("fim") = "presente"
        
        ' Calcular duração até hoje
        dataInicio = DateSerial(CInt(mesAnoInicio(1)), CInt(mesAnoInicio(0)), 1)
        duracaoMeses = DateDiff("m", dataInicio, Date)
        periodoObj("duracao_meses") = duracaoMeses
        periodoObj("texto_original") = periodo
        periodoObj("em_andamento") = True
        
    ElseIf InStr(periodo, " - ") > 0 Then
        ' Período completo (ex: "04/2024 - 12/2025")
        partes = Split(periodo, " - ")
        mesAnoInicio = Split(Trim(partes(0)), "/")
        mesAnoFim = Split(Trim(partes(1)), "/")
        
        periodoObj("inicio") = mesAnoInicio(1) & "-" & Format(mesAnoInicio(0), "00")
        periodoObj("fim") = mesAnoFim(1) & "-" & Format(mesAnoFim(0), "00")
        
        ' Calcular duração
        dataInicio = DateSerial(CInt(mesAnoInicio(1)), CInt(mesAnoInicio(0)), 1)
        dataFim = DateSerial(CInt(mesAnoFim(1)), CInt(mesAnoFim(0)), 1)
        duracaoMeses = DateDiff("m", dataInicio, dataFim)
        periodoObj("duracao_meses") = duracaoMeses
        periodoObj("texto_original") = periodo
        periodoObj("em_andamento") = False
    Else
        ' Caso inválido ou formato diferente - retornar só o texto
        periodoObj("texto_original") = periodo
        periodoObj("inicio") = ""
        periodoObj("fim") = ""
        periodoObj("duracao_meses") = 0
        periodoObj("em_andamento") = False
    End If
    
    Set ParsearPeriodo = periodoObj
End Function

Private Function GerarHistoricoProfissional(ws As Worksheet) As Variant
    Dim empresas, empresaDict, atividadeDict, atividadeDetalhes As Object
    Dim expList() As Object
    Dim ultimaLinha, i, expIndex As Integer
    Dim empresa, cargo, periodo, atividade As String
    Dim descricaoResumida, acoesPrincipais, ferramentas, impacto As String
    Dim incluirNoHistorico As String
    Dim periodoObjTemp As Object
    
    Set empresas = CreateObject("Scripting.Dictionary")
    ultimaLinha = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Agrupar por empresa/cargo/atividade
    For i = 2 To ultimaLinha
        empresa = ws.Cells(i, 1).Value
        cargo = ws.Cells(i, 2).Value
        periodo = ws.Cells(i, 3).Value
        atividade = ws.Cells(i, 4).Value
        descricaoResumida = ws.Cells(i, 5).Value
        acoesPrincipais = ws.Cells(i, 6).Value
        ferramentas = ws.Cells(i, 7).Value
        impacto = ws.Cells(i, 8).Value
        
        incluirNoHistorico = ws.Cells(i, 9).Value

        ' Filtrar apenas experiencias marcadas como "SIM" na coluna I
        If UCase(Trim(incluirNoHistorico)) <> "SIM" Then
            GoTo NextIteration
        End If
        ' Validar se período não está vazio
        If Trim(periodo) = "" Then
            periodo = "Não informado"
        End If
        
        If Not empresas.Exists(empresa) Then
            Set empresaDict = CreateObject("Scripting.Dictionary")
            empresas.Add empresa, empresaDict
        Else
            Set empresaDict = empresas(empresa)
        End If
        
        If Not empresaDict.Exists(cargo) Then
            Set atividadeDict = CreateObject("Scripting.Dictionary")
            Set periodoObjTemp = ParsearPeriodo(periodo)
            Set atividadeDict("periodo_cargo") = periodoObjTemp
            Set atividadeDict("areas_atuacao") = CreateObject("Scripting.Dictionary")
            empresaDict.Add cargo, atividadeDict
        Else
            Set atividadeDict = empresaDict(cargo)
        End If
        
        If Not atividadeDict("areas_atuacao").Exists(atividade) Then
            Set atividadeDetalhes = CreateObject("Scripting.Dictionary")
            atividadeDetalhes("area") = atividade
            atividadeDetalhes("responsabilidades") = Array(CriarResponsabilidadeEstruturada(descricaoResumida, acoesPrincipais, ferramentas, impacto))
            atividadeDict("areas_atuacao").Add atividade, atividadeDetalhes
        Else
            Set atividadeDetalhes = atividadeDict("areas_atuacao")(atividade)
            atividadeDetalhes("responsabilidades") = AppendToArrayObj(atividadeDetalhes("responsabilidades"), CriarResponsabilidadeEstruturada(descricaoResumida, acoesPrincipais, ferramentas, impacto))
        End If
NextIteration:
    Next i
    
    ' Converter para array
    expIndex = 0
    For Each empresaName In empresas.Keys
        Set empresaDict = empresas(empresaName)
        Dim empresaObj As Object
        Set empresaObj = CreateObject("Scripting.Dictionary")
        empresaObj("nome_empresa") = empresaName
        
        Dim cargosList() As Object
        Dim cargoIndex As Integer
        cargoIndex = 0
        
        For Each cargoName In empresaDict.Keys
            Set atividadeDict = empresaDict(cargoName)
            Dim cargoObj As Object
            Set cargoObj = CreateObject("Scripting.Dictionary")
            cargoObj("cargo_nome") = cargoName
            Set cargoObj("periodo_cargo") = atividadeDict("periodo_cargo")
            
            Dim atividadesList() As Object
            Dim atividadeIndex As Integer
            atividadeIndex = 0
            
            For Each atividadeName In atividadeDict("areas_atuacao").Keys
                Set atividadeDetalhes = atividadeDict("areas_atuacao")(atividadeName)
                Dim atividadeObj As Object
                Set atividadeObj = CreateObject("Scripting.Dictionary")
                atividadeObj("area") = atividadeDetalhes("area")
                atividadeObj("responsabilidades") = atividadeDetalhes("responsabilidades")
                
                ReDim Preserve atividadesList(atividadeIndex)
                Set atividadesList(atividadeIndex) = atividadeObj
                atividadeIndex = atividadeIndex + 1
            Next atividadeName
            
            cargoObj("areas_atuacao") = atividadesList
            ReDim Preserve cargosList(cargoIndex)
            Set cargosList(cargoIndex) = cargoObj
            cargoIndex = cargoIndex + 1
        Next cargoName
        
        Call SortByPeriodo(cargosList)
        empresaObj("cargos") = cargosList
        
        ReDim Preserve expList(expIndex)
        Set expList(expIndex) = empresaObj
        expIndex = expIndex + 1
    Next empresaName
    
    GerarHistoricoProfissional = expList
End Function

Private Function GerarFormacao(ws As Worksheet) As Variant
    Dim formacao, situacao As Object
    Dim ultimaLinha, i As Integer
    
    Set formacao = CreateObject("Scripting.Dictionary")
    ultimaLinha = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        Set situacao = CreateObject("Scripting.Dictionary")
        situacao("nome_instituicao") = ws.Cells(i, 1).Value
        situacao("nome_curso") = ws.Cells(i, 2).Value
        situacao("mes_ano_inicio") = ws.Cells(i, 3).Value
        situacao("mes_ano_fim") = ws.Cells(i, 4).Value
        Set formacao(formacao.Count + 1) = situacao
    Next i
    
    GerarFormacao = formacao.Items()
End Function

Private Function GerarCertificacoes(ws As Worksheet) As Variant
    Dim certificacoes, situacao As Object
    Dim ultimaLinha, i As Integer
    
    Set certificacoes = CreateObject("Scripting.Dictionary")
    ultimaLinha = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        Set situacao = CreateObject("Scripting.Dictionary")
        situacao("nome_instituicao") = ws.Cells(i, 1).Value
        situacao("nome_titulo") = ws.Cells(i, 2).Value
        situacao("mes_ano_fim") = ws.Cells(i, 3).Value
        Set certificacoes(certificacoes.Count + 1) = situacao
    Next i
    
    GerarCertificacoes = certificacoes.Items()
End Function

Private Function RemoveUnicodeEscapes(ByVal jsonStr As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\\u([0-9a-fA-F]{4})"
    
    Dim matches As Object
    Set matches = regex.Execute(jsonStr)
    Dim match As Object
    Dim charCode As Long
    
    For Each match In matches
        charCode = CLng("&H" & match.SubMatches(0))
        jsonStr = Replace(jsonStr, match.Value, ChrW(charCode))
    Next
    
    RemoveUnicodeEscapes = jsonStr
End Function

Sub SortByPeriodo(ByRef cargos() As Object)
    Dim i As Integer, j As Integer
    Dim temp As Object
    Dim periodo1 As Date, periodo2 As Date
    
    For i = LBound(cargos) To UBound(cargos) - 1
        For j = i + 1 To UBound(cargos)
            periodo1 = GetPeriodoData(cargos(i)("periodo_cargo"))
            periodo2 = GetPeriodoData(cargos(j)("periodo_cargo"))
            
            If periodo1 < periodo2 Then
                Set temp = cargos(i)
                Set cargos(i) = cargos(j)
                Set cargos(j) = temp
            End If
        Next j
    Next i
End Sub

Function GetPeriodoData(periodoObj As Object) As Date
    Dim mesAno() As String
    Dim dataInicio As String
    
    ' Verificar se é um objeto Dictionary ou string
    On Error Resume Next
    dataInicio = periodoObj("inicio")
    
    If Err.Number <> 0 Then
        ' Se falhou, é string antiga - manter compatibilidade
        On Error GoTo 0
        If InStr(CStr(periodoObj), "Atual") Then
            GetPeriodoData = Date
        Else
            Dim partes() As String
            partes = Split(CStr(periodoObj), " - ")
            mesAno = Split(partes(1), "/")
            GetPeriodoData = DateSerial(CInt(mesAno(1)), CInt(mesAno(0)), 1)
        End If
    Else
        On Error GoTo 0
        ' É um objeto Dictionary
        If periodoObj("em_andamento") Then
            GetPeriodoData = Date
        Else
            mesAno = Split(periodoObj("fim"), "-")
            GetPeriodoData = DateSerial(CInt(mesAno(0)), CInt(mesAno(1)), 1)
        End If
    End If
End Function

Function AppendToArray(ByVal arr As Variant, ByVal item As String) As Variant
    Dim result() As String
    Dim i As Integer
    Dim arrLength As Integer
    
    arrLength = UBound(arr) - LBound(arr) + 1
    ReDim result(arrLength)
    
    For i = LBound(arr) To UBound(arr)
        result(i) = arr(i)
    Next i
    
    result(arrLength) = item
    AppendToArray = result
End Function

Private Function GerarAgenteSugestaoMelhoria() As Object
    Dim agente As Object
    Set agente = CreateObject("Scripting.Dictionary")
    
    ' === Metadados ===
    agente("prioridade") = 1.5
    agente("depende_de") = Array("especialista_storytelling_curriculo")
    agente("output_format") = "markdown"
    agente("output_destination") = "chat_response"
    
    ' === TAREFA FIXA ===
    agente("tarefa") = "Voce e o agente de MELHORIA CONTINUA do especialista em storytelling de curriculos. " & _
                       "Analise com olhar critico o trabalho que o agente 'especialista_storytelling_curriculo' acabou de realizar " & _
                       "nesta execucao especifica. " & _
                       "Identifique exatamente UM ponto de melhoria (o mais impactante) que ele poderia adotar " & _
                       "nas proximas execucoes para gerar bullets ainda mais fortes, quantificados, " & _
                       "orientados a resultados, ATS-friendly e alinhados com as melhores praticas de recrutamento. " & _
                       "Sua saida deve ser uma sugestao concreta que possa ser implementada diretamente na aba 'Aest' " & _
                       "como uma nova regra ou refinamento de regra existente (sem precisar de novas colunas)."

    ' === Critérios de qualidade da sugestão ===
    Dim regras As Object
    Set regras = CreateObject("Scripting.Dictionary")
    
    With regras
        .Add "impacto", CriarRegra("A melhoria deve elevar significativamente a qualidade dos bullets (ex: mais quantificacao, verbos fortes, estrutura STAR, etc.)", 40)
        .Add "especificidade", CriarRegra("A regra sugerida deve ser clara, objetiva e diretamente aplicavel como instrucao de prompt", 30)
        .Add "inovacao", CriarRegra("Priorize melhorias que ainda nao foram aplicadas em execucoes anteriores", 20)
        .Add "simplicidade", CriarRegra("Quanto mais simples de implementar como texto, melhor", 10)
    End With
    
    Set agente("regras_avaliacao") = regras

    ' === Formato de saída estruturado ===
    Dim formato As Object
    Set formato = CreateObject("Scripting.Dictionary")
    
    formato("titulo") = "string"
    formato("problema_identificado") = "string"
    formato("exemplo_antes") = "string"
    formato("exemplo_depois") = "string"
    formato("impacto_esperado") = "string"
    formato("nivel_de_prioridade") = "string"
    
    Set agente("formato_saida_estruturado") = formato

    agente("exemplos_saida") = Array( _
        "{""titulo"":""Melhoria [#1 e continua] – Verbo no imperativo para responsabilidades atuais""," & _
        """problema_identificado"":""Bullets de cargo atual estao no gerundio ('Estou desenvolvendo') ou presente simples fraco.""," & _
        """exemplo_antes"":""Estou responsavel pela gestao de equipe"",""exemplo_depois"":""Lidero equipe de 12 analistas, reduzindo prazo de entrega em 30%""," & _
        """impacto_esperado"":""Aumenta forca e proatividade percebida pelo recrutador"",""nivel_de_prioridade"":""Alta""}" _
    )

    Set GerarAgenteSugestaoMelhoria = agente
End Function

Private Function CriarRegra(descricao As String, peso As Integer) As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("descricao") = descricao
    r("peso") = peso
    Set CriarRegra = r
End Function

Private Function NormalizarNome(ByVal nome As String) As String
    ' Remove acentos comuns
    nome = Replace(nome, "á", "a")
    nome = Replace(nome, "à", "a")
    nome = Replace(nome, "â", "a")
    nome = Replace(nome, "ã", "a")
    nome = Replace(nome, "é", "e")
    nome = Replace(nome, "ê", "e")
    nome = Replace(nome, "í", "i")
    nome = Replace(nome, "ó", "o")
    nome = Replace(nome, "ô", "o")
    nome = Replace(nome, "õ", "o")
    nome = Replace(nome, "ú", "u")
    nome = Replace(nome, "ü", "u")
    nome = Replace(nome, "ç", "c")
    nome = Replace(nome, "Á", "A")
    nome = Replace(nome, "À", "A")
    nome = Replace(nome, "Â", "A")
    nome = Replace(nome, "Ã", "A")
    nome = Replace(nome, "É", "E")
    nome = Replace(nome, "Ê", "E")
    nome = Replace(nome, "Í", "I")
    nome = Replace(nome, "Ó", "O")
    nome = Replace(nome, "Ô", "O")
    nome = Replace(nome, "Õ", "O")
    nome = Replace(nome, "Ú", "U")
    nome = Replace(nome, "Ü", "U")
    nome = Replace(nome, "Ç", "C")
    
    ' Remove parênteses e conteúdo
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\([^)]*\)"
    nome = regex.Replace(nome, "")
    
    ' Trim espaços extras
    nome = Trim(nome)
    
    ' Converte para lowercase
    nome = LCase(nome)
    
    ' Substitui espaços e caracteres especiais por underscore
    nome = Replace(nome, " ", "_")
    nome = Replace(nome, "-", "_")
    nome = Replace(nome, "/", "_")
    nome = Replace(nome, "\", "_")
    nome = Replace(nome, ":", "_")
    
    ' Remove underscores duplicados
    Do While InStr(nome, "__") > 0
        nome = Replace(nome, "__", "_")
    Loop
    
    ' Remove underscore no início e fim
    If Left(nome, 1) = "_" Then nome = Mid(nome, 2)
    If Right(nome, 1) = "_" Then nome = Left(nome, Len(nome) - 1)
    
    NormalizarNome = nome
End Function

Private Function CriarResponsabilidadeEstruturada(ByVal descricao As String, ByVal acoes As String, ByVal ferramentas As String, ByVal impacto As String) As Object
    Dim respObj As Object
    Set respObj = CreateObject("Scripting.Dictionary")
    
    ' Descrição resumida
    respObj("descricao_resumida") = Trim(descricao)
    
    ' Ações principais (split por pipe |)
    If Trim(acoes) <> "" Then
        respObj("acoes_principais") = Split(acoes, "|")
        ' Trim em cada ação
        Dim i As Integer
        Dim acoesArray() As String
        acoesArray = Split(acoes, "|")
        For i = LBound(acoesArray) To UBound(acoesArray)
            acoesArray(i) = Trim(acoesArray(i))
        Next i
        respObj("acoes_principais") = acoesArray
    Else
        respObj("acoes_principais") = Array()
    End If
    
    ' Ferramentas/Sistemas (split por pipe |)
    If Trim(ferramentas) <> "" Then
        Dim ferramentasArray() As String
        ferramentasArray = Split(ferramentas, "|")
        For i = LBound(ferramentasArray) To UBound(ferramentasArray)
            ferramentasArray(i) = Trim(ferramentasArray(i))
        Next i
        respObj("ferramentas_sistemas") = ferramentasArray
    Else
        respObj("ferramentas_sistemas") = Array()
    End If
    
    ' Impacto/Resultado
    respObj("impacto") = Trim(impacto)
    
    Set CriarResponsabilidadeEstruturada = respObj
End Function

Function AppendToArrayObj(ByVal arr As Variant, ByVal item As Object) As Variant
    Dim result() As Object
    Dim i As Integer
    Dim arrLength As Integer
    
    arrLength = UBound(arr) - LBound(arr) + 1
    ReDim result(arrLength)
    
    For i = LBound(arr) To UBound(arr)
        Set result(i) = arr(i)
    Next i
    
    Set result(arrLength) = item
    AppendToArrayObj = result
End Function

