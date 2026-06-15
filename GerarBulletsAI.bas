Attribute VB_Name = "GerarBulletsAI"
Option Explicit

Const API_URL As String = "https://integrate.api.nvidia.com/v1/chat/completions"
Const API_KEY As String = "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"
Const MODEL_NAME As String = "minimaxai/minimax-m3"
Const CAMINHO_JSON As String = "C:\Users\tiago\OneDrive\Desktop\BackUp\Tiago\historico_profissional.json"

Sub GerarBulletsComIA()
    Dim ws As Worksheet
    Dim jsonContent As String
    Dim jsonParsed As Object
    Dim experiencias As Collection
    Dim expItem As Object
    Dim cargos As Collection
    Dim cargoItem As Object
    Dim areas As Collection
    Dim areaItem As Object
    Dim responsabilidades As Collection
    Dim respItem As Object
    Dim acoesPrincipais As Collection
    Dim acaoItem As Variant
    
    ' Config e Idioma Dinâmico
    Dim configObj As Object
    Dim idiomaSaida As String
    Dim instrucaoIA As String
    
    ' Dados do agente especialista
    Dim agenteStorytelling As Object
    
    ' Dados da vaga pretendida
    Dim vagaCargo As String
    Dim vagaRequisitos As Collection
    Dim vagaResponsabilidades As Collection
    Dim vagaReqItem As Variant
    Dim vagaRespItem As Variant
    Dim listaRequisitos As Collection
    Dim listaResponsabilidades As Collection
    Dim vagaDict As Object
    
    Dim empresaNome As String
    Dim cargoNome As String
    Dim areaNome As String
    Dim acaoOriginal As String
    
    Dim promptData As Object
    Dim dataDict As Object
    Dim areasDict As Object
    Dim acoesDaArea As Collection
    Dim jsonString As String
    Dim responseText As String
    
    Dim jsonResponse As Object
    Dim choices As Collection
    Dim msgContent As String
    Dim bulletsJson As Object
    Dim bulletsOtimizados As Collection
    Dim bulletItem As Variant
    
    Dim linhaDestino As Long
    Dim totalEmpresas As Long
    Dim totalAcoes As Long
    Dim startTime As Double
    
    On Error GoTo ErroGeral
    
    startTime = Timer
    Debug.Print "=== INICIANDO PROCESSAMENTO ==="
    
    ' 1. Ler o arquivo JSON
    jsonContent = LerArquivoJSON(CAMINHO_JSON)
    If jsonContent = "" Then
        MsgBox "Falha ao ler o arquivo JSON.", vbCritical
        Exit Sub
    End If
    
    ' 2. Parsear o JSON
    Set jsonParsed = JsonConverter.ParseJson(jsonContent)
    
    ' 3. Extrair CONFIG e IDIOMA DINÂMICO
    If Not jsonParsed.Exists("config") Then Err.Raise vbObjectError + 1, , "Chave 'config' não encontrada."
    Set configObj = jsonParsed("config")
    idiomaSaida = CStr(configObj("idioma_saida"))
    Debug.Print "Idioma de saída dinâmico extraído: " & idiomaSaida
    
    ' Monta a instrução dinâmica baseada no idioma do config
    instrucaoIA = "Você é o agente especialista em storytelling de currículo. Receberá um JSON com: (1) o 'config' do sistema, (2) suas próprias regras no campo 'agente', (3) dados da VAGA PRETENDIDA, (4) dados do CARGO do candidato. " & _
                  "Execute sua tarefa conforme 'agente.tarefa' e 'agente.regras'. " & _
                  "REGRAS OBRIGATÓRIAS: " & _
                  "1. O idioma de saída DEVE ser estritamente: " & idiomaSaida & ". Respeite o 'config.idioma_saida' e não use outros idiomas. " & _
                  "2. Inicie cada bullet com verbo no infinitivo ou substantivo de ação. " & _
                  "3. Mantenha cada bullet entre 35-50 palavras. " & _
                  "4. Foque em resultados e impactos mensuráveis. " & _
                  "RETORNE APENAS um JSON válido no formato: {""bullets_otimizados"": [""bullet 1"", ""bullet 2"", ...]}. Não inclua markdown, explicações ou texto fora do JSON."
    
    ' 4. Extrair dados do AGENTE especialista_storytelling_curriculo
    If Not jsonParsed.Exists("agentes") Then Err.Raise vbObjectError + 2, , "Chave 'agentes' não encontrada."
    If Not jsonParsed("agentes").Exists("especialista_storytelling_curriculo") Then Err.Raise vbObjectError + 3, , "Agente 'especialista_storytelling_curriculo' não encontrado."
    Set agenteStorytelling = jsonParsed("agentes")("especialista_storytelling_curriculo")
    Debug.Print "Agente carregado: " & agenteStorytelling("tarefa")
    
    ' 5. Extrair dados da VAGA PRETENDIDA
    If Not jsonParsed.Exists("contexto") Then Err.Raise vbObjectError + 4, , "Chave 'contexto' não encontrada."
    If Not jsonParsed("contexto").Exists("profissional") Then Err.Raise vbObjectError + 5, , "Chave 'profissional' não encontrada."
    If Not jsonParsed("contexto")("profissional").Exists("candidato_experiencia") Then Err.Raise vbObjectError + 6, , "Chave 'candidato_experiencia' não encontrada."
    
    If jsonParsed("contexto").Exists("vaga_pretendida") Then
        vagaCargo = jsonParsed("contexto")("vaga_pretendida")("vaga_cargo")
        Set vagaRequisitos = jsonParsed("contexto")("vaga_pretendida")("vaga_requisitos")
        Set vagaResponsabilidades = jsonParsed("contexto")("vaga_pretendida")("vaga_responsabilidades")
        
        Debug.Print "Vaga pretendida: " & vagaCargo
        
        Set listaRequisitos = New Collection
        For Each vagaReqItem In vagaRequisitos
            listaRequisitos.Add CStr(vagaReqItem)
        Next vagaReqItem
        
        Set listaResponsabilidades = New Collection
        For Each vagaRespItem In vagaResponsabilidades
            listaResponsabilidades.Add CStr(vagaRespItem)
        Next vagaRespItem
    Else
        Err.Raise vbObjectError + 7, , "Dados da vaga pretendida não encontrados."
    End If
    
    Set experiencias = jsonParsed("contexto")("profissional")("candidato_experiencia")
    totalEmpresas = experiencias.count
    Debug.Print "Total de empresas: " & totalEmpresas
    
    ' 6. Preparar a Planilha
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Bullets_Gerados")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "Bullets_Gerados"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErroGeral
    
    ws.Range("A1:C1").Value = Array("Empresa", "Cargo", "Bullet Otimizado (IA)")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(200, 220, 255)
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 100
    linhaDestino = 2
    
    ' 7. Loop pelas EMPRESAS e CARGOS
    For Each expItem In experiencias
        empresaNome = expItem("nome_empresa")
        Debug.Print vbCrLf & ">>> EMPRESA: " & empresaNome
        
        If expItem.Exists("cargos") Then
            Set cargos = expItem("cargos")
            Debug.Print "  Total de cargos: " & cargos.count
            
            For Each cargoItem In cargos
                cargoNome = cargoItem("cargo_nome")
                Debug.Print "  -> CARGO: " & cargoNome
                
                Set areasDict = CreateObject("Scripting.Dictionary")
                totalAcoes = 0
                
                If cargoItem.Exists("areas_atuacao") Then
                    Set areas = cargoItem("areas_atuacao")
                    
                    For Each areaItem In areas
                        areaNome = areaItem("area")
                        Debug.Print "     (Área: " & areaNome & ")"
                        
                        Set acoesDaArea = New Collection
                        
                        If areaItem.Exists("responsabilidades") Then
                            Set responsabilidades = areaItem("responsabilidades")
                            
                            For Each respItem In responsabilidades
                                If respItem.Exists("acoes_principais") Then
                                    Set acoesPrincipais = respItem("acoes_principais")
                                    
                                    For Each acaoItem In acoesPrincipais
                                        acaoOriginal = CStr(acaoItem)
                                        If Trim(acaoOriginal) <> "" Then
                                            acoesDaArea.Add acaoOriginal
                                            totalAcoes = totalAcoes + 1
                                        End If
                                    Next acaoItem
                                End If
                            Next respItem
                        End If
                        
                        If acoesDaArea.count > 0 Then
                            areasDict.Add areaNome, acoesDaArea
                        End If
                    Next areaItem
                End If
                
                Debug.Print "     Total de ações coletadas para este cargo: " & totalAcoes
                
                If areasDict.count > 0 Then
                    Debug.Print "     -> Enviando cargo completo para IA..."
                    
                    Set promptData = CreateObject("Scripting.Dictionary")
                    promptData.Add "instruction", instrucaoIA
                    
                    Set dataDict = CreateObject("Scripting.Dictionary")
                    
                    ' Envia o config completo (com idioma_saida dinâmico)
                    dataDict.Add "config", configObj
                    
                    ' Dados do AGENTE
                    dataDict.Add "agente", agenteStorytelling
                    
                    ' Dados da VAGA
                    Set vagaDict = CreateObject("Scripting.Dictionary")
                    vagaDict.Add "cargo_pretendido", vagaCargo
                    vagaDict.Add "requisitos", listaRequisitos
                    vagaDict.Add "responsabilidades", listaResponsabilidades
                    dataDict.Add "vaga_pretendida", vagaDict
                    
                    ' Dados do CANDIDATO
                    dataDict.Add "empresa", empresaNome
                    dataDict.Add "cargo", cargoNome
                    dataDict.Add "areas_atuacao", areasDict
                    
                    promptData.Add "data", dataDict
                    
                    jsonString = JsonConverter.ConvertToJson(promptData)
                    
                    responseText = ChamarAPINvidia(jsonString)
                    
                    If responseText <> "" Then
                        Debug.Print "     -> Resposta recebida, processando..."
                        
                        Set jsonResponse = JsonConverter.ParseJson(responseText)
                        
                        If jsonResponse.Exists("choices") Then
                            Set choices = jsonResponse("choices")
                            If choices.count > 0 Then
                                msgContent = choices(1)("message")("content")
                                
                                msgContent = Replace(msgContent, "```json", "")
                                msgContent = Replace(msgContent, "```", "")
                                msgContent = Trim(msgContent)
                                
                                Set bulletsJson = JsonConverter.ParseJson(msgContent)
                                
                                If bulletsJson.Exists("bullets_otimizados") Then
                                    Set bulletsOtimizados = bulletsJson("bullets_otimizados")
                                    Debug.Print "     -> Bullets gerados: " & bulletsOtimizados.count
                                    
                                    For Each bulletItem In bulletsOtimizados
                                        ws.Cells(linhaDestino, 1).Value = empresaNome
                                        ws.Cells(linhaDestino, 2).Value = cargoNome
                                        ws.Cells(linhaDestino, 3).Value = CStr(bulletItem)
                                        linhaDestino = linhaDestino + 1
                                    Next bulletItem
                                Else
                                    Debug.Print "     -> ERRO: 'bullets_otimizados' não encontrado na resposta."
                                End If
                            End If
                        End If
                    Else
                        Debug.Print "     -> ERRO: Resposta vazia da API."
                    End If
                End If
            Next cargoItem
        End If
    Next expItem
    
    Dim tempoTotal As Double
    tempoTotal = Timer - startTime
    
    Debug.Print vbCrLf & "=== PROCESSAMENTO CONCLUÍDO ==="
    Debug.Print "Total de bullets gerados: " & (linhaDestino - 2)
    Debug.Print "Tempo total: " & Format(tempoTotal / 60, "0.00") & " minutos"
    
    MsgBox "Processo concluído!" & vbCrLf & _
           (linhaDestino - 2) & " bullets gerados na aba 'Bullets_Gerados'.", vbInformation
    Exit Sub

ErroGeral:
    Debug.Print "ERRO CRÍTICO: " & Err.Description
    MsgBox "Erro crítico: " & Err.Description, vbCritical, "Falha no Processo"
End Sub

' --- Função Auxiliar: Ler Arquivo JSON (UTF-8 com remoção de BOM) ---
Function LerArquivoJSON(caminhoArquivo As String) As String
    Dim stream As Object
    Dim conteudo As String
    
    On Error GoTo ErroLeitura
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile caminhoArquivo
    conteudo = stream.ReadText(-1)
    stream.Close
    Set stream = Nothing
    
    If Len(conteudo) > 0 Then
        If AscW(Left(conteudo, 1)) = 65279 Then conteudo = Mid(conteudo, 2)
    End If
    
    LerArquivoJSON = Trim(conteudo)
    Exit Function

ErroLeitura:
    MsgBox "Erro ao ler arquivo: " & Err.Description, vbCritical
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    LerArquivoJSON = ""
End Function

' --- Função Auxiliar: Chamar API NVIDIA ---
Function ChamarAPINvidia(promptJson As String) As String
    Dim http As Object
    Dim finalBody As String
    Dim messagesColl As Collection
    Dim msgDict As Object
    Dim bodyDict As Object
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Set messagesColl = New Collection
    Set msgDict = CreateObject("Scripting.Dictionary")
    msgDict.Add "role", "user"
    msgDict.Add "content", promptJson
    messagesColl.Add msgDict
    
    Set bodyDict = CreateObject("Scripting.Dictionary")
    bodyDict.Add "model", MODEL_NAME
    bodyDict.Add "messages", messagesColl
    bodyDict.Add "temperature", 0.7
    bodyDict.Add "max_tokens", 2048
    
    finalBody = JsonConverter.ConvertToJson(bodyDict)
    
    On Error GoTo ErroAPI
    
    http.Open "POST", API_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & API_KEY
    http.setRequestHeader "Accept", "application/json"
    
    http.send finalBody
    
    If http.Status = 200 Then
        ChamarAPINvidia = http.responseText
    Else
        Debug.Print "Erro API: " & http.Status & " - " & http.statusText
        ChamarAPINvidia = ""
    End If
    
    Set http = Nothing
    Exit Function

ErroAPI:
    Debug.Print "Erro HTTP: " & Err.Description
    Set http = Nothing
    ChamarAPINvidia = ""
End Function

