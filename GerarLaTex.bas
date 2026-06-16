Attribute VB_Name = "GerarLaTex"
Option Explicit

' ============================================================================
' Módulo: GerarLaTex
' Descrição: Gera arquivo LaTeX de currículo baseado nos bullets gerados pela IA
' e no template definido (PT ou EN conforme idioma_saida do JSON)
' ============================================================================

' ----------------------------------------------------------------------------
' Função Principal: Gera o arquivo LaTeX completo
' ----------------------------------------------------------------------------
Public Sub GerarCurriculoLaTeX()
    On Error GoTo ErrorHandler
    
    Dim wsBullets As Worksheet
    Dim jsonStr As String
    Dim json As Object
    Dim idioma As String
    Dim template As String
    Dim outputContent As String
    Dim lastRow As Long
    Dim i As Long
    Dim cargoNome As String
    Dim bullets As String
    Dim experienciaBlock As String
    Dim formacaoBlock As String
    
    ' Definir planilha de bullets
    Set wsBullets = ThisWorkbook.Sheets("Bullets_Gerados")
    
    ' Ler JSON de configuração
    jsonStr = LerJSONConfig()
    If jsonStr = "" Then
        MsgBox "Erro: Não foi possível ler o JSON historico_profissional", vbCritical
        Exit Sub
    End If
    
    ' Parsear JSON
    Set json = JsonConverter.ParseJson(jsonStr)
    
    ' Obter idioma de saída
    idioma = UCase$(json("config")("idioma_saida"))
    
    ' Carregar template adequado
    template = CarregarTemplate(idioma)
    If template = "" Then
        MsgBox "Erro: Template não encontrado para idioma " & idioma, vbCritical
        Exit Sub
    End If
    
    ' Processar experiências profissionais
    experienciaBlock = GerarBlocoExperiencias(wsBullets, json)
    
    ' Processar formação acadêmica
    formacaoBlock = GerarBlocoFormacao(json)
    
    ' Substituir placeholders no template
    outputContent = template
    
    ' Substituir bloco de experiências
    outputContent = SubstituirExperiencias(outputContent, experienciaBlock)
    
    ' Substituir bloco de formação
    outputContent = SubstituirFormacao(outputContent, formacaoBlock)
    
    ' Salvar arquivo LaTeX
    SalvarArquivoLaTeX outputContent, idioma
    
    MsgBox "Currículo LaTeX gerado com sucesso!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao gerar currículo LaTeX: " & Err.Description, vbCritical
End Sub

' ----------------------------------------------------------------------------
' Função: Lê o JSON historico_profissional
' ----------------------------------------------------------------------------
Private Function LerJSONConfig() As String
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Dim jsonCell As Range
    Dim jsonStr As String
    
    ' Procurar JSON na aba config ou onde estiver armazenado
    For Each wsConfig In ThisWorkbook.Worksheets
        If InStr(1, wsConfig.Name, "config", vbTextCompare) > 0 Or _
           InStr(1, wsConfig.Name, "json", vbTextCompare) > 0 Then
            Set jsonCell = wsConfig.Cells.Find("historico_profissional", LookIn:=xlValues)
            If Not jsonCell Is Nothing Then
                jsonStr = jsonCell.Offset(0, 1).Value
                If Len(jsonStr) > 100 Then
                    LerJSONConfig = jsonStr
                    Exit Function
                End If
            End If
        End If
    Next wsConfig
    
    ' Tentar encontrar em qualquer célula da workbook
    For Each wsConfig In ThisWorkbook.Worksheets
        For Each jsonCell In wsConfig.UsedRange
            If InStr(1, jsonCell.Value, """historico_profissional""", vbTextCompare) > 0 Then
                jsonStr = jsonCell.Value
                If Len(jsonStr) > 100 Then
                    LerJSONConfig = jsonStr
                    Exit Function
                End If
            End If
        Next jsonCell
    Next wsConfig
    
    LerJSONConfig = ""
End Function

' ----------------------------------------------------------------------------
' Função: Carrega o template LaTeX baseado no idioma
' ----------------------------------------------------------------------------
Private Function CarregarTemplate(idioma As String) As String
    Dim templatePath As String
    Dim fso As Object
    Dim ts As Object
    Dim templateContent As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Definir caminho do template
    If idioma = "PT-BR" Or idioma = "PT" Then
        templatePath = ThisWorkbook.Path & "\templates\template_pt.tex"
    ElseIf idioma = "EN-US" Or idioma = "EN" Then
        templatePath = ThisWorkbook.Path & "\templates\template_en.tex"
    Else
        ' Default para PT-BR
        templatePath = ThisWorkbook.Path & "\templates\template_pt.tex"
    End If
    
    ' Tentar ler arquivo de template
    On Error Resume Next
    If fso.FileExists(templatePath) Then
        Set ts = fso.OpenTextFile(templatePath, 1)
        templateContent = ts.ReadAll
        ts.Close
        CarregarTemplate = templateContent
        Exit Function
    End If
    On Error GoTo 0
    
    ' Se não encontrar arquivo, usar templates embutidos
    If idioma = "PT-BR" Or idioma = "PT" Then
        CarregarTemplate = ObterTemplatePT()
    ElseIf idioma = "EN-US" Or idioma = "EN" Then
        CarregarTemplate = ObterTemplateEN()
    Else
        CarregarTemplate = ObterTemplatePT()
    End If
End Function

' ----------------------------------------------------------------------------
' Função: Gera o bloco de experiências em LaTeX
' ----------------------------------------------------------------------------
Private Function GerarBlocoExperiencias(ws As Worksheet, json As Object) As String
    Dim experiencias As Object
    Dim exp As Object
    Dim cargos As Object
    Dim cargo As Object
    Dim areas As Object
    Dim area As Object
    Dim responsabilidades As Object
    Dim resp As Object
    Dim bloqueio As String
    Dim cargoNome As String
    Dim nomeEmpresa As String
    Dim periodoInicio As String
    Dim periodoFim As String
    Dim duracaoMeses As Long
    Dim emAndamento As Boolean
    Dim bulletsColunaC As String
    Dim localizacao As String
    Dim i As Long, j As Long, k As Long, l As Long
    
    bloqueio = ""
    
    ' Iterar sobre candidato_experiencia
    Set experiencias = json("contexto")("profissional")("candidato_experiencia")
    
    For Each exp In experiencias
        nomeEmpresa = exp("nome_empresa")
        
        ' Iterar sobre cargos da empresa
        Set cargos = exp("cargos")
        
        For Each cargo In cargos
            cargoNome = cargo("cargo_nome")
            periodoInicio = cargo("periodo_cargo")("inicio")
            periodoFim = cargo("periodo_cargo")("fim")
            emAndamento = cargo("periodo_cargo")("em_andamento")
            
            ' Formatar período
            If emAndamento Then
                periodoFim = FormatPeriodoAtual(json("config")("idioma_saida"))
            End If
            
            ' Extrair localização das áreas de atuação (pegar a primeira)
            localizacao = ExtrairLocalizacao(cargo, nomeEmpresa)
            
            ' Buscar bullets na planilha Bullets_Gerados
            bulletsColunaC = BuscarBulletsCargo(ws, cargoNome)
            
            ' Se não encontrou bullets, tentar gerar a partir das responsabilidades do JSON
            If bulletsColunaC = "" Then
                bulletsColunaC = GerarBulletsDoJSON(cargo, json("config")("idioma_saida"))
            End If
            
            ' Criar bloco LaTeX para este cargo
            bloqueio = bloqueio & CriarItemExperiencia(cargoNome, nomeEmpresa, localizacao, _
                                                        periodoInicio, periodoFim, _
                                                        bulletsColunaC, json("config")("idioma_saida"))
        Next cargo
    Next exp
    
    GerarBlocoExperiencias = bloqueio
End Function

' ----------------------------------------------------------------------------
' Função: Extrai localização do cargo ou usa padrão da empresa
' ----------------------------------------------------------------------------
Private Function ExtrairLocalizacao(cargo As Object, nomeEmpresa As String) As String
    On Error Resume Next
    
    Dim areas As Object
    Dim area As Object
    Dim local As String
    
    ' Tentar extrair das áreas de atuação
    If cargo.Exists("areas_atuacao") Then
        Set areas = cargo("areas_atuacao")
        If areas.Count > 0 Then
            ' Procurar por padrões de localização nas responsabilidades
            For Each area In areas
                If area.Exists("responsabilidades") Then
                    ' Verificar se há menção a localização
                    ' Por enquanto, retornar vazio e usar padrão
                End If
            Next area
        End If
    End If
    
    ' Mapeamento básico de empresas para localidades
    Select Case True
        Case InStr(1, nomeEmpresa, "Agrobiotech", vbTextCompare) > 0
            local = "Ribeirão Preto/SP"
        Case InStr(1, nomeEmpresa, "BTZ", vbTextCompare) > 0
            local = "Londrina/PR"
        Case InStr(1, nomeEmpresa, "Schindler", vbTextCompare) > 0 Or _
             InStr(1, nomeEmpresa, "Atlas Schindler", vbTextCompare) > 0
            local = "Londrina/PR"
        Case InStr(1, nomeEmpresa, "Greenwich", vbTextCompare) > 0
            local = "São Paulo/SP"
        Case Else
            local = ""
    End Select
    
    ExtrairLocalizacao = local
End Function

' ----------------------------------------------------------------------------
' Função: Busca bullets na planilha Bullets_Gerados coluna C baseado no cargo (coluna B)
' ----------------------------------------------------------------------------
Private Function BuscarBulletsCargo(ws As Worksheet, cargoNome As String) As String
    Dim lastRow As Long
    Dim i As Long
    Dim bullets As String
    Dim cargoPlanilha As String
    
    bullets = ""
    
    ' Encontrar última linha com dados na coluna B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Iterar pelas linhas buscando o cargo
    For i = 2 To lastRow
        cargoPlanilha = Trim$(ws.Cells(i, "B").Value)
        
        ' Comparação case-insensitive e parcial
        If InStr(1, cargoPlanilha, cargoNome, vbTextCompare) > 0 Or _
           InStr(1, cargoNome, cargoPlanilha, vbTextCompare) > 0 Then
            
            ' Adicionar bullet da coluna C
            If ws.Cells(i, "C").Value <> "" Then
                If bullets <> "" Then bullets = bullets & vbCrLf
                bullets = bullets & Trim$(ws.Cells(i, "C").Value)
            End If
        End If
    Next i
    
    BuscarBulletsCargo = bullets
End Function

' ----------------------------------------------------------------------------
' Função: Gera bullets a partir das responsabilidades do JSON (fallback)
' ----------------------------------------------------------------------------
Private Function GerarBulletsDoJSON(cargo As Object, idioma As String) As String
    On Error Resume Next
    
    Dim areas As Object
    Dim area As Object
    Dim responsabilidades As Object
    Dim resp As Object
    Dim bullets As String
    Dim descricao As String
    
    bullets = ""
    
    If Not cargo.Exists("areas_atuacao") Then
        GerarBulletsDoJSON = ""
        Exit Function
    End If
    
    Set areas = cargo("areas_atuacao")
    
    For Each area In areas
        If area.Exists("responsabilidades") Then
            Set responsabilidades = area("responsabilidades")
            
            For Each resp In responsabilidades
                If resp.Exists("descricao_resumida") Then
                    descricao = resp("descricao_resumida")
                    If bullets <> "" Then bullets = bullets & vbCrLf
                    bullets = bullets & "- " & descricao
                End If
            Next resp
        End If
    Next area
    
    GerarBulletsDoJSON = bullets
End Function

' ----------------------------------------------------------------------------
' Função: Cria item LaTeX para uma experiência
' ----------------------------------------------------------------------------
Private Function CriarItemExperiencia(cargoNome As String, nomeEmpresa As String, localizacao As String, _
                                      periodoInicio As String, periodoFim As String, _
                                      bullets As String, idioma As String) As String
    Dim item As String
    Dim bulletLines As Variant
    Dim bulletLine As Variant
    Dim bulletStr As String
    Dim i As Long
    
    ' Formatar períodos
    periodoInicio = FormatarPeriodo(periodoInicio)
    periodoFim = FormatarPeriodo(periodoFim)
    
    ' Traduzir "Atual" se necessário
    If UCase$(idioma) = "EN-US" Or UCase$(idioma) = "EN" Then
        If periodoFim = "Atual" Then periodoFim = "Present"
    End If
    
    ' Construir cabeçalho do item
    item = "    \item[] % Empty item to control spacing" & vbCrLf
    item = item & "    \begin{minipage}[t]{\textwidth}" & vbCrLf
    item = item & "        \textbf{\Large " & cargoNome & "} \, | \, \textbf{" & nomeEmpresa
    If localizacao <> "" Then
        item = item & " - " & localizacao
    End If
    item = item & "} \hfill \textbf{\small " & periodoInicio & " - " & periodoFim & "}" & vbCrLf
    item = item & "    \end{minipage}" & vbCrLf
    item = item & "    \begin{itemize}" & vbCrLf
    
    ' Adicionar bullets
    If bullets <> "" Then
        bulletLines = Split(bullets, vbCrLf)
        For Each bulletLine In bulletLines
            bulletStr = Trim$(bulletLine)
            ' Remover prefixos como "-", "•", etc.
            bulletStr = LimparPrefixoBullet(bulletStr)
            If bulletStr <> "" Then
                item = item & "        \item " & bulletStr & vbCrLf
            End If
        Next bulletLine
    End If
    
    item = item & "    \end{itemize}" & vbCrLf
    item = item & vbCrLf
    
    CriarItemExperiencia = item
End Function

' ----------------------------------------------------------------------------
' Função: Limpa prefixos de bullet (-, •, *, etc.)
' ----------------------------------------------------------------------------
Private Function LimparPrefixoBullet(texto As String) As String
    Dim result As String
    result = Trim$(texto)
    
    ' Remover prefixos comuns
    If Left$(result, 2) = "- " Then result = Mid$(result, 3)
    If Left$(result, 1) = "-" Then result = Mid$(result, 2)
    If Left$(result, 2) = "• " Then result = Mid$(result, 3)
    If Left$(result, 1) = "•" Then result = Mid$(result, 2)
    If Left$(result, 2) = "* " Then result = Mid$(result, 3)
    If Left$(result, 1) = "*" Then result = Mid$(result, 2)
    
    LimparPrefixoBullet = Trim$(result)
End Function

' ----------------------------------------------------------------------------
' Função: Gera o bloco de formação acadêmica em LaTeX
' ----------------------------------------------------------------------------
Private Function GerarBlocoFormacao(json As Object) As String
    Dim formacoes As Object
    Dim formacao As Object
    Dim bloqueio As String
    Dim nomeInstituicao As String
    Dim nomeCurso As String
    Dim inicio As String
    Dim fim As String
    Dim idioma As String
    
    bloqueio = ""
    idioma = json("config")("idioma_saida")
    
    If Not json("contexto").Exists("profissional") Then
        GerarBlocoFormacao = ""
        Exit Function
    End If
    
    If Not json("contexto")("profissional").Exists("candidato_formacao") Then
        GerarBlocoFormacao = ""
        Exit Function
    End If
    
    Set formacoes = json("contexto")("profissional")("candidato_formacao")
    
    For Each formacao In formacoes
        nomeInstituicao = formacao("nome_instituicao")
        nomeCurso = formacao("nome_curso")
        inicio = FormatarPeriodo(formacao("mes_ano_inicio"))
        fim = FormatarPeriodo(formacao("mes_ano_fim"))
        
        ' Traduzir nomes de cursos se necessário
        If UCase$(idioma) = "EN-US" Or UCase$(idioma) = "EN" Then
            nomeCurso = TraduzirCurso(nomeCurso, idioma)
        End If
        
        bloqueio = bloqueio & "    \item[] % Empty item to control spacing" & vbCrLf
        bloqueio = bloqueio & "    \begin{minipage}[t]{\textwidth}" & vbCrLf
        bloqueio = bloqueio & "        \textbf{" & nomeCurso & "} \" & vbCrLf
        bloqueio = bloqueio & "        \textbf{" & nomeInstituicao & "} {\color{mainblue}$\bullet$} " & inicio & " - " & fim & vbCrLf
        bloqueio = bloqueio & "    \end{minipage}" & vbCrLf
        bloqueio = bloqueio & vbCrLf
        bloqueio = bloqueio & "    \vspace{0.5em} % Adding separation between entries" & vbCrLf
        bloqueio = bloqueio & vbCrLf
    Next formacao
    
    GerarBlocoFormacao = bloqueio
End Function

' ----------------------------------------------------------------------------
' Função: Traduz nomes de cursos para inglês
' ----------------------------------------------------------------------------
Private Function TraduzirCurso(curso As String, idioma As String) As String
    Dim cursoTraduzido As String
    
    Select Case True
        Case InStr(1, curso, "MBA", vbTextCompare) > 0 And InStr(1, curso, "Comércio Exterior", vbTextCompare) > 0
            cursoTraduzido = "MBA in International Trade"
        Case InStr(1, curso, "Superior", vbTextCompare) > 0 And InStr(1, curso, "Comércio Exterior", vbTextCompare) > 0
            cursoTraduzido = "Bachelor's Degree in International Trade"
        Case InStr(1, curso, "Engenharia", vbTextCompare) > 0
            cursoTraduzido = "Engineering Degree"
        Case InStr(1, curso, "Administração", vbTextCompare) > 0
            cursoTraduzido = "Business Administration"
        Case Else
            cursoTraduzido = curso
    End Select
    
    TraduzirCurso = cursoTraduzido
End Function

' ----------------------------------------------------------------------------
' Função: Formata período (MM/YYYY)
' ----------------------------------------------------------------------------
Private Function FormatarPeriodo(periodo As String) As String
    ' Já assume formato MM/YYYY ou similar
    ' Pode ser ajustado para validar/formatar conforme necessário
    FormatarPeriodo = periodo
End Function

' ----------------------------------------------------------------------------
' Função: Retorna texto para período atual baseado no idioma
' ----------------------------------------------------------------------------
Private Function FormatPeriodoAtual(idioma As String) As String
    If UCase$(idioma) = "EN-US" Or UCase$(idioma) = "EN" Then
        FormatPeriodoAtual = "Present"
    Else
        FormatPeriodoAtual = "Atual"
    End If
End Function

' ----------------------------------------------------------------------------
' Função: Substitui bloco de experiências no template
' ----------------------------------------------------------------------------
Private Function SubstituirExperiencias(template As String, experienciaBlock As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim sectionStart As String
    Dim sectionEnd As String
    
    ' Procurar seção EXPERIÊNCIA / PROFESSIONAL EXPERIENCE
    sectionStart = "\section*{EXPERIÊNCIA}"
    If InStr(1, template, "PROFESSIONAL EXPERIENCE", vbTextCompare) > 0 Then
        sectionStart = "\section*{PROFESSIONAL EXPERIENCE}"
    End If
    
    startPos = InStr(1, template, sectionStart, vbTextCompare)
    If startPos = 0 Then
        SubstituirExperiencias = template
        Exit Function
    End If
    
    ' Encontrar início do environment itemize após a seção
    startPos = InStr(startPos, template, "\begin{itemize}")
    If startPos = 0 Then
        SubstituirExperiencias = template
        Exit Function
    End If
    
    ' Encontrar fim do environment itemize
    endPos = InStr(startPos, template, "\end{itemize}")
    If endPos = 0 Then
        SubstituirExperiencias = template
        Exit Function
    End If
    
    ' Substituir conteúdo
    SubstituirExperiencias = Left$(template, startPos - 1) & vbCrLf & _
                             "\item[] % Empty item to control spacing" & vbCrLf & _
                             experienciaBlock & _
                             Mid$(template, endPos)
End Function

' ----------------------------------------------------------------------------
' Função: Substitui bloco de formação no template
' ----------------------------------------------------------------------------
Private Function SubstituirFormacao(template As String, formacaoBlock As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim sectionStart As String
    
    ' Procurar seção FORMAÇÃO ACADÊMICA / EDUCATION
    sectionStart = "\section*{FORMAÇÃO ACADÊMICA}"
    If InStr(1, template, "EDUCATION", vbTextCompare) > 0 Then
        sectionStart = "\section*{EDUCATION}"
    End If
    
    startPos = InStr(1, template, sectionStart, vbTextCompare)
    If startPos = 0 Then
        SubstituirFormacao = template
        Exit Function
    End If
    
    ' Encontrar início do environment itemize após a seção
    startPos = InStr(startPos, template, "\begin{itemize}")
    If startPos = 0 Then
        SubstituirFormacao = template
        Exit Function
    End If
    
    ' Encontrar fim do environment itemize
    endPos = InStr(startPos, template, "\end{itemize}")
    If endPos = 0 Then
        SubstituirFormacao = template
        Exit Function
    End If
    
    ' Substituir conteúdo
    SubstituirFormacao = Left$(template, startPos - 1) & vbCrLf & _
                         formacaoBlock & _
                         Mid$(template, endPos)
End Function

' ----------------------------------------------------------------------------
' Sub: Salva o arquivo LaTeX
' ----------------------------------------------------------------------------
Private Sub SalvarArquivoLaTeX(content As String, idioma As String)
    Dim fso As Object
    Dim ts As Object
    Dim filePath As String
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Definir nome do arquivo
    If UCase$(idioma) = "EN-US" Or UCase$(idioma) = "EN" Then
        fileName = "curriculo_en.tex"
    Else
        fileName = "curriculo_pt.tex"
    End If
    
    filePath = ThisWorkbook.Path & "\" & fileName
    
    ' Salvar arquivo
    Set ts = fso.CreateTextFile(filePath, True, True) ' True = overwrite, True = Unicode
    ts.Write content
    ts.Close
    
    MsgBox "Arquivo salvo em: " & filePath, vbInformation
End Sub

' ----------------------------------------------------------------------------
' Função: Template LaTeX em Português (embutido)
' ----------------------------------------------------------------------------
Private Function ObterTemplatePT() As String
    ObterTemplatePT = "\documentclass[11pt]{article}" & vbCrLf & vbCrLf & _
"% Including necessary packages" & vbCrLf & _
"\usepackage[utf8]{inputenc}" & vbCrLf & _
"\usepackage[T1]{fontenc}" & vbCrLf & _
"\usepackage{microtype}" & vbCrLf & _
"\usepackage{garamondlibre}" & vbCrLf & _
"\usepackage[a4paper, left=0.6in, right=0.6in, top=0.5in, bottom=0.5in]{geometry}" & vbCrLf & _
"\usepackage{enumitem}" & vbCrLf & _
"\usepackage{xcolor}" & vbCrLf & _
"\usepackage[hidelinks]{hyperref}" & vbCrLf & _
"\usepackage{fontawesome5}" & vbCrLf & _
"\usepackage{pifont}" & vbCrLf & vbCrLf & _
"% Defining custom colors" & vbCrLf & _
"\definecolor{mainblue}{RGB}{0,70,127}" & vbCrLf & vbCrLf & _
"% Customizing section headers" & vbCrLf & _
"\usepackage{titlesec}" & vbCrLf & _
"\titleformat{\section}{\large\bfseries\scshape\color{mainblue}}{\thesection}{1em}{}[\color{mainblue}\titlerule]" & vbCrLf & _
"\titlespacing*{\section}{0pt}{1em}{0.5em}" & vbCrLf & vbCrLf & _
"% Customizing itemize for tighter spacing and ensuring bullet symbol" & vbCrLf & _
"\setlist[itemize]{leftmargin=2.0em, itemsep=0.2em, topsep=0.2em, label=\textcolor{mainblue}{\raisebox{0.3ex}{\large\textbullet}}}" & vbCrLf & vbCrLf & _
"% Removing page numbers" & vbCrLf & _
"\pagestyle{empty}" & vbCrLf & vbCrLf & _
"\begin{document}" & vbCrLf & vbCrLf & _
"% Header with name and desired position aligned to the left with horizontal lines" & vbCrLf & _
"\begin{flushleft}" & vbCrLf & _
"    {\Huge \textbf{\color{mainblue}TIAGO JÚLIO DE MORAES}} \\[0.5em]" & vbCrLf & _
"    {\LARGE \textbf{\color{mainblue}Coordenador de Importação e Exportação}} \\[0.5em]" & vbCrLf & _
"    \small" & vbCrLf & _
"    \begin{tabular}{@{}l@{\hspace{1em}}l@{\hspace{1em}}l@{\hspace{1em}}l@{\hspace{1em}}l@{\hspace{1em}}l@{\hspace{1em}}l@{\hspace{1em}}l@{}}" & vbCrLf & _
"            \color{mainblue}\faWhatsapp & \color{mainblue}(43) 99804-9453 &" & vbCrLf & _
"            \color{mainblue}\faEnvelope & \color{mainblue}\href{mailto:tiagoj.moraes@gmail.com}{tiagoj.moraes@gmail.com} &" & vbCrLf & _
"            \color{mainblue}\faLinkedin & \color{mainblue}\href{https://www.linkedin.com/in/tiagojmoraes/ }{LinkedIn} &" & vbCrLf & _
"            \color{mainblue}\faMapMarker* & \color{mainblue}Ribeirão Preto, SP (disponível p/ mudança)\\" & vbCrLf & _
"    \end{tabular}" & vbCrLf & _
"    \vspace{0.5em}" & vbCrLf & _
"    {\color{mainblue}\hrule height 0.4pt depth 0pt}" & vbCrLf & _
"\end{flushleft}" & vbCrLf & vbCrLf & _
"% Professional Summary Section" & vbCrLf & _
"\section*{PERFIL PROFISSIONAL}" & vbCrLf & _
"Profissional de Comércio Exterior com mais de 15 anos de experiência na gestão de operações de importação, exportação e logística internacional, com atuação na liderança de equipes e estruturação de processos. Experiência na condução de iniciativas de melhoria contínua, automação e gestão de indicadores (KPIs), com foco em eficiência operacional, conformidade aduaneira e regimes especiais (Drawback e Ex-Tarifário), além de otimização de custos, contribuindo para a tomada de decisão e evolução da cadeia de suprimentos." & vbCrLf & vbCrLf & _
"% Experiência Section" & vbCrLf & _
"\section*{EXPERIÊNCIA}" & vbCrLf & _
"\begin{itemize}[leftmargin=0pt]" & vbCrLf & _
"% EXPERIENCIAS_AQUI" & vbCrLf & _
"\end{itemize}" & vbCrLf & vbCrLf & _
"% Formação Acadêmica Section" & vbCrLf & _
"\section*{FORMAÇÃO ACADÊMICA}" & vbCrLf & _
"\begin{itemize}[leftmargin=0pt]" & vbCrLf & _
"% FORMACAO_AQUI" & vbCrLf & _
"\end{itemize}" & vbCrLf & vbCrLf & _
"% Idiomas Section" & vbCrLf & _
"\section*{IDIOMAS}" & vbCrLf & _
"\textbf{Inglês} \quad \rule{2cm}{0.4pt} \quad Avançado" & vbCrLf & vbCrLf & _
"% Habilidades Section" & vbCrLf & _
"\section*{HABILIDADES}" & vbCrLf & _
"SISCOMEX \quad $|$ \quad Excel \quad $|$ \quad VBA \quad $|$ \quad SQL \quad $|$ \quad Power BI \quad $|$ \quad SAP \quad $|$ \quad ImportSys \quad $|$ \quad LOGIX \quad $|$ \quad Protheus \quad $|$ \quad iGlobal" & vbCrLf & vbCrLf & _
"\end{document}"
End Function

' ----------------------------------------------------------------------------
' Função: Template LaTeX em Inglês (embutido)
' ----------------------------------------------------------------------------
Private Function ObterTemplateEN() As String
    ObterTemplateEN = "\documentclass[11pt]{article}" & vbCrLf & vbCrLf & _
"% Including necessary packages" & vbCrLf & _
"\usepackage[utf8]{inputenc}" & vbCrLf & _
"\usepackage[T1]{fontenc}" & vbCrLf & _
"\usepackage{microtype}" & vbCrLf & _
"\usepackage{garamondlibre}" & vbCrLf & _
"\usepackage[a4paper, left=0.6in, right=0.6in, top=0.5in, bottom=0.5in]{geometry}" & vbCrLf & _
"\usepackage{enumitem}" & vbCrLf & _
"\usepackage{xcolor}" & vbCrLf & _
"\usepackage[hidelinks]{hyperref}" & vbCrLf & _
"\usepackage{fontawesome5}" & vbCrLf & _
"\usepackage{pifont}" & vbCrLf & vbCrLf & _
"% Defining custom colors" & vbCrLf & _
"\definecolor{mainblue}{RGB}{0,70,127}" & vbCrLf & vbCrLf & _
"% Customizing section headers" & vbCrLf & _
"\usepackage{titlesec}" & vbCrLf & _
"\titleformat{\section}{\large\bfseries\scshape\color{mainblue}}{\thesection}{1em}{}[\color{mainblue}\titlerule]" & vbCrLf & _
"\titlespacing*{\section}{0pt}{1em}{0.5em}" & vbCrLf & vbCrLf & _
"% Customizing itemize for tighter spacing and ensuring bullet symbol" & vbCrLf & _
"\setlist[itemize]{leftmargin=2.0em, itemsep=0.2em, topsep=0.2em, label=\textcolor{mainblue}{\raisebox{0.3ex}{\large\textbullet}}}" & vbCrLf & vbCrLf & _
"% Removing page numbers" & vbCrLf & _
"\pagestyle{empty}" & vbCrLf & vbCrLf & _
"\begin{document}" & vbCrLf & vbCrLf & _
"% Header with name and desired position aligned to the left with horizontal lines" & vbCrLf & _
"\begin{flushleft}" & vbCrLf & _
"    {\Huge \textbf{\color{mainblue}TIAGO JÚLIO DE MORAES}} \\[0.5em]" & vbCrLf & _
"    {\LARGE \textbf{\color{mainblue}Trade Operations Supervisor – Import and Export}} \\[0.5em]" & vbCrLf & _
"    \small" & vbCrLf & _
"        \begin{tabular}{@{}l@{\hspace{1em}}l@{\hspace{2em}}l@{\hspace{1em}}l@{\hspace{2em}}l@{\hspace{1em}}l@{\hspace{2em}}l@{\hspace{1em}}l@{}}" & vbCrLf & _
"            \color{mainblue}\faWhatsapp & \color{mainblue}(43) 99804-9453 &" & vbCrLf & _
"            \color{mainblue}\faEnvelope & \color{mainblue}\href{mailto:tiagoj.moraes@gmail.com}{tiagoj.moraes@gmail.com} &" & vbCrLf & _
"            \color{mainblue}\faLinkedin & \color{mainblue}\href{https://www.linkedin.com/in/tiagojmoraes/ }{LinkedIn} &" & vbCrLf & _
"            \color{mainblue}\faMapMarker* & \color{mainblue}Ribeirão Preto, SP" & vbCrLf & _
"    \end{tabular}" & vbCrLf & _
"    \vspace{0.5em}" & vbCrLf & _
"    {\color{mainblue}\hrule height 0.4pt depth 0pt}" & vbCrLf & _
"\end{flushleft}" & vbCrLf & vbCrLf & _
"% Professional Summary Section" & vbCrLf & _
"\section*{PROFESSIONAL SUMMARY}" & vbCrLf & _
"Professional with over 15 years of experience in import, export, and international logistics, managing operations" & vbCrLf & _
"and integrating Foreign Trade, Logistics, and Supply Chain operations. Expertise in special regimes (Drawback, Ex-Tariff)," & vbCrLf & _
"negotiation with domestic and international suppliers, performance indicators (KPIs) management, and process" & vbCrLf & _
"automation. Focused on reducing costs, standardizing processes, ensuring customs compliance, and optimizing fiscal benefits." & vbCrLf & vbCrLf & _
"% Experiência Section" & vbCrLf & _
"\section*{PROFESSIONAL EXPERIENCE}" & vbCrLf & _
"\begin{itemize}[leftmargin=0pt]" & vbCrLf & _
"% EXPERIENCIAS_AQUI" & vbCrLf & _
"\end{itemize}" & vbCrLf & vbCrLf & _
"% Formação Acadêmica Section" & vbCrLf & _
"\section*{EDUCATION}" & vbCrLf & _
"\begin{itemize}[leftmargin=0pt]" & vbCrLf & _
"% FORMACAO_AQUI" & vbCrLf & _
"\end{itemize}" & vbCrLf & vbCrLf & _
"% Idiomas Section" & vbCrLf & _
"\section*{LANGUAGES}" & vbCrLf & _
"\textbf{English} \quad \rule{2cm}{0.4pt} \quad Advanced \\" & vbCrLf & _
"\textbf{Spanish} \quad \rule{2cm}{0.4pt} \quad Intermediate" & vbCrLf & vbCrLf & _
"% Habilidades Section" & vbCrLf & _
"\section*{SKILLS}" & vbCrLf & _
"SISCOMEX \quad $|$ \quad Excel \quad $|$ \quad VBA \quad $|$ \quad SQL \quad $|$ \quad Power BI \quad $|$ \quad SAP \quad $|$ \quad ImportSys \quad $|$ \quad LOGIX \quad $|$ \quad Protheus \quad $|$ \quad iGlobal" & vbCrLf & vbCrLf & _
"\end{document}"
End Function
