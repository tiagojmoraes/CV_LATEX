Attribute VB_Name = "Gerar_LaTeX_Melhorias"

' ========== GERAR LATEX A PARTIR DOS BULLETS ==========
Sub GerarLaTeX()
    Dim wsBullets As Worksheet, wsExp As Worksheet, wsForm As Worksheet
    Dim wsHab As Worksheet, wsDadosPessoais As Worksheet
    Dim latexOutput As String
    Dim idiomaEscolhido As String
    Dim resposta As VbMsgBoxResult
    Dim cargoPretendido As String
    Dim resumoProfissional As String
    Dim nomeCompleto As String
    Dim telefone As String, email As String, linkedin As String, localizacao As String
    Dim habilidadesLista As String
    Dim i As Integer, ultimaLinha As Integer
    Dim empresaAnterior As String
    
    ' Selecionar idioma
    resposta = MsgBox("Escolha o idioma do LaTeX:" & vbCrLf & vbCrLf & _
                      "Sim = Portugues" & vbCrLf & _
                      "Nao = Ingles", _
                      vbQuestion + vbYesNo, "Idioma do LaTeX")
    
    If resposta = vbYes Then
        idiomaEscolhido = "pt-BR"
    Else
        idiomaEscolhido = "en-US"
    End If
    
    ' Verificar se existe a aba de bullets
    On Error Resume Next
    Set wsBullets = ThisWorkbook.Sheets("Bullets_Gerados")
    If wsBullets Is Nothing Then
        MsgBox "A aba 'Bullets_Gerados' nao foi encontrada!" & vbCrLf & _
               "Por favor, cole os bullets gerados pela IA nesta aba.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Set wsExp = ThisWorkbook.Sheets("Experiencias")
    Set wsForm = ThisWorkbook.Sheets("Formacao")
    Set wsHab = ThisWorkbook.Sheets("Habilidades")
    
    ' Cabeçalho do LaTeX
    latexOutput = "\documentclass[11pt]{article}" & vbCrLf & vbCrLf
    latexOutput = latexOutput & "% Including necessary packages" & vbCrLf
    latexOutput = latexOutput & "\usepackage[utf8]{inputenc}" & vbCrLf
    latexOutput = latexOutput & "\usepackage[T1]{fontenc}" & vbCrLf
    latexOutput = latexOutput & "\usepackage{microtype}" & vbCrLf
    latexOutput = latexOutput & "\usepackage{garamondlibre}" & vbCrLf
    latexOutput = latexOutput & "\usepackage[a4paper, left=0.6in, right=0.6in, top=0.5in, bottom=0.5in]{geometry}" & vbCrLf
    latexOutput = latexOutput & "\usepackage{enumitem}" & vbCrLf
    latexOutput = latexOutput & "\usepackage{xcolor}" & vbCrLf
    latexOutput = latexOutput & "\usepackage[hidelinks]{hyperref} % Remove borders from links" & vbCrLf
    latexOutput = latexOutput & "\usepackage{fontawesome5}" & vbCrLf
    latexOutput = latexOutput & "\usepackage{pifont}" & vbCrLf & vbCrLf
    
    ' Cores e formatacao
    latexOutput = latexOutput & "% Defining custom colors" & vbCrLf
    latexOutput = latexOutput & "\definecolor{mainblue}{RGB}{0,70,127}" & vbCrLf & vbCrLf
    
    latexOutput = latexOutput & "% Customizing section headers" & vbCrLf
    latexOutput = latexOutput & "\usepackage{titlesec}" & vbCrLf
    latexOutput = latexOutput & "\titleformat{\section}{\large\bfseries\scshape\color{mainblue}}{\thesection}{1em}{}[\color{mainblue}\titlerule]" & vbCrLf
    latexOutput = latexOutput & "\titlespacing*{\section}{0pt}{1em}{0.5em}" & vbCrLf & vbCrLf
    
    latexOutput = latexOutput & "% Customizing itemize for tighter spacing and ensuring bullet symbol" & vbCrLf
    latexOutput = latexOutput & "\setlist[itemize]{leftmargin=2.0em, itemsep=0.2em, topsep=0.2em, label=\textcolor{mainblue}{\raisebox{0.3ex}{\large\textbullet}}}" & vbCrLf & vbCrLf
    
    latexOutput = latexOutput & "% Removing page numbers" & vbCrLf
    latexOutput = latexOutput & "\pagestyle{empty}" & vbCrLf & vbCrLf
    
    latexOutput = latexOutput & "\begin{document}" & vbCrLf & vbCrLf
    
    ' Header - buscar dados
    On Error Resume Next
    Set wsDadosPessoais = ThisWorkbook.Sheets("Dados_Pessoais")
    If Not wsDadosPessoais Is Nothing Then
        nomeCompleto = wsDadosPessoais.Range("B2").Value
        telefone = wsDadosPessoais.Range("B3").Value
        email = wsDadosPessoais.Range("B4").Value
        linkedin = wsDadosPessoais.Range("B5").Value
        localizacao = wsDadosPessoais.Range("B6").Value
    End If
    
    If nomeCompleto = "" Then nomeCompleto = "NOME COMPLETO"
    If telefone = "" Then telefone = "(00) 00000-0000"
    If email = "" Then email = "email@exemplo.com"
    If linkedin = "" Then linkedin = "linkedin.com/in/usuario"
    If localizacao = "" Then localizacao = "Cidade, Estado"
    
    cargoPretendido = ThisWorkbook.Sheets("VagaPretendida").Range("A2").Value
    If cargoPretendido = "" Then cargoPretendido = "Cargo Pretendido"
    On Error GoTo 0
    
    latexOutput = latexOutput & "% Header with name and desired position" & vbCrLf
    latexOutput = latexOutput & "\begin{flushleft}" & vbCrLf
    latexOutput = latexOutput & "    {\Huge \textbf{\color{mainblue}" & UCase(nomeCompleto) & "}} \\[0.5em]" & vbCrLf
    latexOutput = latexOutput & "    {\LARGE \textbf{\color{mainblue}" & cargoPretendido & "}} \\[0.5em]" & vbCrLf
    latexOutput = latexOutput & "    \small" & vbCrLf
    latexOutput = latexOutput & "    \begin{tabular}{@{}l@{\hspace{1em}}l@{\hspace{2em}}l@{\hspace{1em}}l@{\hspace{2em}}l@{\hspace{1em}}l@{\hspace{2em}}l@{\hspace{1em}}l@{}}" & vbCrLf
    latexOutput = latexOutput & "        \color{mainblue}\faWhatsapp & \color{mainblue}" & telefone & " &" & vbCrLf
    latexOutput = latexOutput & "        \color{mainblue}\faEnvelope & \color{mainblue}\href{mailto:" & email & "}{" & email & "} &" & vbCrLf
    latexOutput = latexOutput & "        \color{mainblue}\faLinkedin & \color{mainblue}\href{https://www.linkedin.com/in/" & Replace(linkedin, "https://www.linkedin.com/in/", "") & "}{LinkedIn} &" & vbCrLf
    latexOutput = latexOutput & "        \color{mainblue}\faMapMarker* & \color{mainblue}" & localizacao & vbCrLf
    latexOutput = latexOutput & "    \end{tabular}" & vbCrLf
    latexOutput = latexOutput & "    \vspace{0.5em}" & vbCrLf
    latexOutput = latexOutput & "    {\color{mainblue}\hrule height 0.4pt depth 0pt}" & vbCrLf
    latexOutput = latexOutput & "\end{flushleft}" & vbCrLf & vbCrLf
    
    ' Professional Summary
    If idiomaEscolhido = "pt-BR" Then
        latexOutput = latexOutput & "% Professional Summary Section" & vbCrLf
        latexOutput = latexOutput & "\section*{PERFIL PROFISSIONAL}" & vbCrLf
        resumoProfissional = wsBullets.Range("A1").Value
        If resumoProfissional = "" Then resumoProfissional = "Resumo profissional aqui."
        latexOutput = latexOutput & resumoProfissional & vbCrLf & vbCrLf
    Else
        latexOutput = latexOutput & "% Professional Summary Section" & vbCrLf
        latexOutput = latexOutput & "\section*{PROFESSIONAL SUMMARY}" & vbCrLf
        resumoProfissional = wsBullets.Range("A1").Value
        If resumoProfissional = "" Then resumoProfissional = "Professional summary here."
        latexOutput = latexOutput & resumoProfissional & vbCrLf & vbCrLf
    End If
    
    ' Experiencia Profissional
    If idiomaEscolhido = "pt-BR" Then
        latexOutput = latexOutput & "% Experiencia Section" & vbCrLf
        latexOutput = latexOutput & "\section*{EXPERIENCIA}" & vbCrLf
    Else
        latexOutput = latexOutput & "% Experience Section" & vbCrLf
        latexOutput = latexOutput & "\section*{PROFESSIONAL EXPERIENCE}" & vbCrLf
    End If
    
    latexOutput = latexOutput & "\begin{itemize}[leftmargin=0pt]" & vbCrLf
    
    ' Processar cada experiencia da aba Bullets_Gerados
    ultimaLinha = wsBullets.Cells(Rows.Count, 1).End(xlUp).Row
    empresaAnterior = ""
    
    For i = 2 To ultimaLinha
        Dim empresa As String, cargo As String, periodo As String, bullets As String
        empresa = wsBullets.Cells(i, 1).Value
        cargo = wsBullets.Cells(i, 2).Value
        periodo = wsBullets.Cells(i, 3).Value
        bullets = wsBullets.Cells(i, 4).Value
        
        If empresa <> "" Then
            If empresaAnterior <> "" Then
                latexOutput = latexOutput & "    \end{itemize}" & vbCrLf
                latexOutput = latexOutput & vbCrLf
            End If
            
            latexOutput = latexOutput & "\item[] % Empty item to control spacing" & vbCrLf
            latexOutput = latexOutput & "    \begin{minipage}[t]{\textwidth}" & vbCrLf
            latexOutput = latexOutput & "        \textbf{\Large " & cargo & "} \, | \, \textbf{" & empresa & "} \hfill \textbf{\small " & periodo & "}" & vbCrLf
            latexOutput = latexOutput & "    \end{minipage}" & vbCrLf
            latexOutput = latexOutput & "    \begin{itemize}" & vbCrLf
            
            empresaAnterior = empresa
        End If
        
        If bullets <> "" Then
            latexOutput = latexOutput & "        \item " & bullets & vbCrLf
        End If
    Next i
    
    If empresaAnterior <> "" Then
        latexOutput = latexOutput & "    \end{itemize}" & vbCrLf
    End If
    
    latexOutput = latexOutput & "\end{itemize}" & vbCrLf & vbCrLf
    
    ' Formacao Academica
    If idiomaEscolhido = "pt-BR" Then
        latexOutput = latexOutput & "% Formacao Academica Section" & vbCrLf
        latexOutput = latexOutput & "\section*{FORMACAO ACADEMICA}" & vbCrLf
    Else
        latexOutput = latexOutput & "% Education Section" & vbCrLf
        latexOutput = latexOutput & "\section*{EDUCATION}" & vbCrLf
    End If
    
    latexOutput = latexOutput & "\begin{itemize}[leftmargin=0pt]" & vbCrLf
    
    ultimaLinha = wsForm.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        Dim instituicao As String, curso As String, inicio As String, fim As String
        instituicao = wsForm.Cells(i, 1).Value
        curso = wsForm.Cells(i, 2).Value
        inicio = wsForm.Cells(i, 3).Value
        fim = wsForm.Cells(i, 4).Value
        
        latexOutput = latexOutput & "    \item[] % Empty item to control spacing" & vbCrLf
        latexOutput = latexOutput & "    \begin{minipage}[t]{\textwidth}" & vbCrLf
        latexOutput = latexOutput & "        \textbf{" & curso & "} \\" & vbCrLf
        latexOutput = latexOutput & "        \textbf{" & instituicao & "} {\color{mainblue}$\bullet$} " & inicio & " - " & fim & vbCrLf
        latexOutput = latexOutput & "    \end{minipage}" & vbCrLf
        
        If i < ultimaLinha Then
            latexOutput = latexOutput & "    " & vbCrLf
            latexOutput = latexOutput & "    \vspace{0.5em} % Adding separation between entries" & vbCrLf
            latexOutput = latexOutput & "    " & vbCrLf
        End If
    Next i
    
    latexOutput = latexOutput & "\end{itemize}" & vbCrLf & vbCrLf
    
    ' Idiomas
    If idiomaEscolhido = "pt-BR" Then
        latexOutput = latexOutput & "% Idiomas Section" & vbCrLf
        latexOutput = latexOutput & "\section*{IDIOMAS}" & vbCrLf
        latexOutput = latexOutput & "\textbf{Ingles} \quad \rule{2cm}{0.4pt} \quad Avancado" & vbCrLf
    Else
        latexOutput = latexOutput & "% Languages Section" & vbCrLf
        latexOutput = latexOutput & "\section*{LANGUAGES}" & vbCrLf
        latexOutput = latexOutput & "\textbf{English} \quad \rule{2cm}{0.4pt} \quad Advanced" & vbCrLf
    End If
    latexOutput = latexOutput & vbCrLf
    
    ' Habilidades
    habilidadesLista = wsHab.Range("A2").Value
    
    If idiomaEscolhido = "pt-BR" Then
        latexOutput = latexOutput & "% Habilidades Section" & vbCrLf
        latexOutput = latexOutput & "\section*{HABILIDADES}" & vbCrLf
    Else
        latexOutput = latexOutput & "% Skills Section" & vbCrLf
        latexOutput = latexOutput & "\section*{SKILLS}" & vbCrLf
    End If
    
    Dim habilidadesArray() As String
    Dim habItem As String
    habilidadesArray = Split(habilidadesLista, ", ")
    
    For i = 0 To UBound(habilidadesArray)
        habItem = Trim(habilidadesArray(i))
        If i < UBound(habilidadesArray) Then
            latexOutput = latexOutput & habItem & " \quad $|$ \quad "
        Else
            latexOutput = latexOutput & habItem
        End If
    Next i
    
    latexOutput = latexOutput & vbCrLf & vbCrLf
    latexOutput = latexOutput & "\end{document}"
    
    ' Salvar arquivo LaTeX
    Dim stream As Object
    Dim filePath As String
    Set stream = CreateObject("ADODB.Stream")
    
    If idiomaEscolhido = "pt-BR" Then
        filePath = ThisWorkbook.Path & "\curriculo_pt.tex"
    Else
        filePath = ThisWorkbook.Path & "\resume_en.tex"
    End If
    
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText latexOutput
        .SaveToFile filePath, 2
        .Close
    End With
    
    MsgBox "LaTeX gerado com sucesso!" & vbCrLf & _
           "Idioma: " & IIf(idiomaEscolhido = "pt-BR", "Portugues", "Ingles") & vbCrLf & _
           "Arquivo: " & filePath, vbInformation, "Sucesso!"
End Sub

' ========== MELHORAR RESULTADOS (LOOP DE MELHORIA CONTINUA) ==========
Sub MelhorarResultados()
    Dim wsExp As Worksheet, wsMelhoria As Worksheet
    Dim json As Object, contexto As Object, profissional As Object
    Dim historicoArray As Variant
    Dim jsonString As String
    Dim promptMelhoria As String
    Dim i As Integer, j As Integer, k As Integer
    Dim experienciaObj As Object, cargoObj As Object, atividadeObj As Object
    Dim bulletsAtuais As String
    
    ' Verificar se existe a aba de melhoria
    On Error Resume Next
    Set wsMelhoria = ThisWorkbook.Sheets("Melhores_Resultados")
    If wsMelhoria Is Nothing Then
        Set wsMelhoria = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMelhoria.Name = "Melhores_Resultados"
        
        ' Configurar cabecalhos
        wsMelhoria.Range("A1").Value = "Empresa"
        wsMelhoria.Range("B1").Value = "Cargo"
        wsMelhoria.Range("C1").Value = "Periodo"
        wsMelhoria.Range("D1").Value = "Bullets Atuais"
        wsMelhoria.Range("E1").Value = "Sugestao de Melhoria"
        wsMelhoria.Range("F1").Value = "Bullets Otimizados"
        wsMelhoria.Range("G1").Value = "Nota Qualidade (1-10)"
        wsMelhoria.Range("H1").Value = "Acoes Implementadas"
        
        wsMelhoria.Range("A1:H1").Font.Bold = True
        wsMelhoria.Range("A1:H1").Interior.Color = RGB(0, 70, 127)
        wsMelhoria.Range("A1:H1").Font.Color = RGB(255, 255, 255)
    End If
    On Error GoTo 0
    
    Set wsExp = ThisWorkbook.Sheets("Experiencias")
    
    ' Limpar dados anteriores (mantendo cabecalho)
    wsMelhoria.Range("A2:H1000").ClearContents
    
    ' Gerar JSON com historico profissional filtrado (apenas SIM na coluna I)
    Set json = CreateObject("Scripting.Dictionary")
    Set contexto = CreateObject("Scripting.Dictionary")
    Set profissional = CreateObject("Scripting.Dictionary")
    
    historicoArray = GerarHistoricoProfissional(wsExp)
    profissional("candidato_experiencia") = historicoArray
    
    Set contexto("profissional") = profissional
    Set json("contexto") = contexto
    
    ' Converter para JSON
    jsonString = JsonConverter.ConvertToJson(json, Whitespace:=2)
    jsonString = RemoveUnicodeEscapes(jsonString)
    
    ' Criar prompt de melhoria para cada experiencia
    Dim row As Integer
    row = 2
    
    ' Iterar sobre experiencias
    For i = 0 To UBound(historicoArray)
        Set experienciaObj = historicoArray(i)
        Dim nomeEmpresa As String
        nomeEmpresa = experienciaObj("nome_empresa")
        
        ' Iterar sobre cargos
        For j = 0 To UBound(experienciaObj("cargos"))
            Set cargoObj = experienciaObj("cargos")(j)
            Dim nomeCargo As String, periodoCargo As String
            nomeCargo = cargoObj("cargo_nome")
            
            ' Extrair periodo do objeto
            If cargoObj.Exists("periodo_cargo") Then
                If cargoObj("periodo_cargo")("em_andamento") Then
                    periodoCargo = cargoObj("periodo_cargo")("inicio") & " - Atual"
                Else
                    periodoCargo = cargoObj("periodo_cargo")("inicio") & " - " & cargoObj("periodo_cargo")("fim")
                End If
            Else
                periodoCargo = "Nao informado"
            End If
            
            ' Iterar sobre areas de atuacao
            For k = 0 To UBound(cargoObj("areas_atuacao"))
                Set atividadeObj = cargoObj("areas_atuacao")(k)
                Dim area As String
                area = atividadeObj("area")
                
                ' Iterar sobre responsabilidades
                Dim l As Integer
                For l = 0 To UBound(atividadeObj("responsabilidades"))
                    Dim respObj As Object
                    Set respObj = atividadeObj("responsabilidades")(l)
                    
                    ' Montar bullet atual
                    Dim descricao As String, acoes As String, impacto As String
                    descricao = respObj("descricao_resumida")
                    
                    If respObj.Exists("acoes_principais") Then
                        acoes = Join(respObj("acoes_principais"), "; ")
                    End If
                    
                    If respObj.Exists("impacto") Then
                        impacto = respObj("impacto")
                    End If
                    
                    bulletsAtuais = descricao & " | " & acoes & " | " & impacto
                    
                    ' Preencher planilha
                    wsMelhoria.Cells(row, 1).Value = nomeEmpresa
                    wsMelhoria.Cells(row, 2).Value = nomeCargo
                    wsMelhoria.Cells(row, 3).Value = periodoCargo
                    wsMelhoria.Cells(row, 4).Value = bulletsAtuais
                    
                    ' Criar prompt de melhoria
                    promptMelhoria = "Analise o seguinte bullet point de curriculo e sugira melhorias:" & vbCrLf & _
                                    "Bullet atual: " & bulletsAtuais & vbCrLf & vbCrLf & _
                                    "Criterios de melhoria:" & vbCrLf & _
                                    "1. Adicione metricas e resultados quantificaveis" & vbCrLf & _
                                    "2. Use verbos de acao fortes no inicio" & vbCrLf & _
                                    "3. Destaque impacto nos negocios" & vbCrLf & _
                                    "4. Otimize para ATS (palavras-chave relevantes)" & vbCrLf & _
                                    "5. Mantenha concisao (maximo 2 linhas)" & vbCrLf & vbCrLf & _
                                    "Forneca:" & vbCrLf & _
                                    "- Nota de qualidade atual (1-10)" & vbCrLf & _
                                    "- Sugestao especifica de melhoria" & vbCrLf & _
                                    "- Versao otimizada do bullet"
                    
                    wsMelhoria.Cells(row, 5).Value = promptMelhoria
                    wsMelhoria.Cells(row, 6).Formula = "=D" & row
                    wsMelhoria.Cells(row, 7).Value = "Aguardando analise"
                    wsMelhoria.Cells(row, 8).Value = "Pendente"
                    
                    row = row + 1
                Next l
            Next k
        Next j
    Next i
    
    ' Ajustar largura das colunas
    wsMelhoria.Columns("A:C").AutoFit
    wsMelhoria.Columns("D:D").ColumnWidth = 40
    wsMelhoria.Columns("E:E").ColumnWidth = 50
    wsMelhoria.Columns("F:F").ColumnWidth = 40
    wsMelhoria.Columns("G:G").ColumnWidth = 20
    wsMelhoria.Columns("H:H").ColumnWidth = 25
    
    ' Congelar paineis
    wsMelhoria.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    MsgBox "Planilha 'Melhores_Resultados' criada com sucesso!" & vbCrLf & _
           "Total de bullets analisados: " & (row - 2) & vbCrLf & vbCrLf & _
           "Proximos passos:" & vbCrLf & _
           "1. Clique em 'MelhorarBulletsComIA' para processar automaticamente" & vbCrLf & _
           "2. Ou revise manualmente as sugestoes na coluna E", vbInformation, "Sucesso!"
End Sub

' ========== FUNCOES AUXILIARES ==========

' Gerar historico profissional filtrado (apenas experiencias marcadas como SIM)
Function GerarHistoricoProfissional(wsExp As Worksheet) As Variant
    Dim historico() As Object
    Dim i As Integer, expIndex As Integer
    Dim ultimaLinha As Integer
    Dim incluirComoBoolean As Boolean
    
    ultimaLinha = wsExp.Cells(Rows.Count, 1).End(xlUp).Row
    expIndex = 0
    ReDim historico(0 To 10) ' Array dinamico inicial
    
    For i = 2 To ultimaLinha
        ' Verificar coluna I (Incluir?)
        If UCase(Trim(wsExp.Cells(i, 9).Value)) = "SIM" Then
            Dim empresaObj As Object
            Set empresaObj = CreateObject("Scripting.Dictionary")
            
            empresaObj("nome_empresa") = wsExp.Cells(i, 1).Value
            
            ' Criar array de cargos
            Dim cargos() As Object
            ReDim cargos(0 To 0)
            Set cargos(0) = CreateObject("Scripting.Dictionary")
            
            cargos(0)("cargo_nome") = wsExp.Cells(i, 2).Value
            
            ' Periodo
            Dim periodoObj As Object
            Set periodoObj = CreateObject("Scripting.Dictionary")
            periodoObj("inicio") = wsExp.Cells(i, 3).Value
            periodoObj("fim") = wsExp.Cells(i, 4).Value
            periodoObj("em_andamento") = (UCase(Trim(wsExp.Cells(i, 5).Value)) = "SIM")
            cargos(0)("periodo_cargo") = periodoObj
            
            ' Areas de atuacao
            Dim areas() As Object
            ReDim areas(0 To 0)
            Set areas(0) = CreateObject("Scripting.Dictionary")
            areas(0)("area") = wsExp.Cells(i, 6).Value
            
            ' Responsabilidades
            Dim resp() As Object
            ReDim resp(0 To 0)
            Set resp(0) = CreateObject("Scripting.Dictionary")
            resp(0)("descricao_resumida") = wsExp.Cells(i, 7).Value
            resp(0)("acoes_principais") = Split(wsExp.Cells(i, 8).Value, ";")
            resp(0)("impacto") = wsExp.Cells(i, 10).Value
            
            areas(0)("responsabilidades") = resp
            cargos(0)("areas_atuacao") = areas
            empresaObj("cargos") = cargos
            
            Set historico(expIndex) = empresaObj
            expIndex = expIndex + 1
            
            If expIndex > UBound(historico) Then
                ReDim Preserve historico(0 To expIndex + 10)
            End If
        End If
    Next i
    
    ' Redimensionar array final
    If expIndex > 0 Then
        ReDim Preserve historico(0 To expIndex - 1)
        GerarHistoricoProfissional = historico
    Else
        GerarHistoricoProfissional = Array()
    End If
End Function

' Remover escapes Unicode do JSON
Function RemoveUnicodeEscapes(jsonString As String) As String
    Dim result As String
    result = jsonString
    
    ' Substituir escapes Unicode comuns
    result = Replace(result, "\u00e9", "é")
    result = Replace(result, "\u00ea", "ê")
    result = Replace(result, "\u00ed", "í")
    result = Replace(result, "\u00f3", "ó")
    result = Replace(result, "\u00f4", "ô")
    result = Replace(result, "\u00fa", "ú")
    result = Replace(result, "\u00e7", "ç")
    result = Replace(result, "\u00c7", "Ç")
    result = Replace(result, "\u00e3", "ã")
    result = Replace(result, "\u00f5", "õ")
    result = Replace(result, "\u00e1", "á")
    result = Replace(result, "\u00c1", "Á")
    result = Replace(result, "\u00c9", "É")
    result = Replace(result, "\u00ca", "Ê")
    result = Replace(result, "\u00cd", "Í")
    result = Replace(result, "\u00d3", "Ó")
    result = Replace(result, "\u00d4", "Ô")
    result = Replace(result, "\u00da", "Ú")
    result = Replace(result, "\u00d5", "Õ")
    
    RemoveUnicodeEscapes = result
End Function

' ========== MELHORAR BULLETS COM IA (API NVIDIA) ==========
Sub MelhorarBulletsComIA()
    Dim wsMelhoria As Worksheet
    Dim row As Integer, ultimaLinha As Integer
    Dim bulletAtual As String
    Dim respostaIA As String
    Dim notaQualidade As String
    Dim bulletOtimizado As String
    
    ' Verificar se existe a aba de melhoria
    On Error Resume Next
    Set wsMelhoria = ThisWorkbook.Sheets("Melhores_Resultados")
    If wsMelhoria Is Nothing Then
        MsgBox "A aba 'Melhores_Resultados' nao foi encontrada!" & vbCrLf & _
               "Execute primeiro a funcao 'MelhorarResultados'.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ultimaLinha = wsMelhoria.Cells(Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha < 2 Then
        MsgBox "Nenhum bullet encontrado para melhorar!", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim processedCount As Integer
    processedCount = 0
    
    ' Iterar sobre cada bullet
    For row = 2 To ultimaLinha
        bulletAtual = wsMelhoria.Cells(row, 4).Value
        
        If bulletAtual <> "" And wsMelhoria.Cells(row, 6).Value = "" Then
            ' Chamar API para melhorar o bullet
            respostaIA = ChamarAPINvidia(bulletAtual)
            
            If respostaIA <> "" Then
                ' Parse da resposta (formato esperado: JSON ou texto estruturado)
                Call ExtrairRespostaIA(respostaIA, notaQualidade, bulletOtimizado)
                
                wsMelhoria.Cells(row, 6).Value = bulletOtimizado
                wsMelhoria.Cells(row, 7).Value = notaQualidade
                wsMelhoria.Cells(row, 8).Value = "Implementado via IA"
                
                processedCount = processedCount + 1
                
                ' Pequena pausa para nao sobrecarregar a API
                DoEvents
            End If
        End If
    Next row
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Processo concluido!" & vbCrLf & _
           "Bullets processados: " & processedCount & vbCrLf & _
           "Total de bullets: " & (ultimaLinha - 1), vbInformation, "Melhoria com IA"
End Sub

' Chamar API da NVIDIA
Function ChamarAPINvidia(bulletAtual As String) As String
    Dim httpReq As Object
    Dim apiUrl As String
    Dim apiKey As String
    Dim requestBody As String
    Dim responseText As String
    Dim promptCompleto As String
    
    ' Configuracoes da API
    apiUrl = "https://integrate.api.nvidia.com/v1/chat/completions"
    apiKey = "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"
    
    ' Criar prompt otimizado
    promptCompleto = "Analise o seguinte bullet point de curriculo e sugira melhorias:" & vbCrLf & vbCrLf & _
                     "Bullet atual: " & bulletAtual & vbCrLf & vbCrLf & _
                     "Criterios de melhoria:" & vbCrLf & _
                     "1. Adicione metricas e resultados quantificaveis (use numeros, porcentagens, valores)" & vbCrLf & _
                     "2. Use verbos de acao fortes no inicio (liderou, implementou, otimizou, etc.)" & vbCrLf & _
                     "3. Destaque impacto nos negocios (receita, economia, eficiencia, satisfacao)" & vbCrLf & _
                     "4. Otimize para ATS (palavras-chave relevantes da area)" & vbCrLf & _
                     "5. Mantenha concisao (maximo 2 linhas)" & vbCrLf & vbCrLf & _
                     "Forneca APENAS no formato JSON abaixo, sem texto adicional:" & vbCrLf & _
                     "{" & vbCrLf & _
                     "  ""nota_qualidade"": ""7""," & vbCrLf & _
                     "  ""bullet_otimizado"": ""Versao melhorada do bullet aqui""" & vbCrLf & _
                     "}"
    
    ' Montar request body
    requestBody = "{""model"":""minimaxai/minimax-m3""," & _
                  """messages"":[{""role"":""user"",""content"":""" & EscapeJson(promptCompleto) & """}]," & _
                  """max_tokens"":8192," & _
                  """temperature"":1.0," & _
                  """top_p"":0.95," & _
                  """stream"":false}"
    
    ' Criar objeto HTTP
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    
    On Error Resume Next
    httpReq.Open "POST", apiUrl, False
    httpReq.setRequestHeader "Content-Type", "application/json"
    httpReq.setRequestHeader "Authorization", "Bearer " & apiKey
    httpReq.setRequestHeader "Accept", "application/json"
    httpReq.send requestBody
    
    If httpReq.Status = 200 Then
        responseText = httpReq.responseText
        
        ' Extrair conteudo da resposta
        ChamarAPINvidia = ExtrairConteudoResposta(responseText)
    Else
        ChamarAPINvidia = ""
        Debug.Print "Erro API: " & httpReq.Status & " - " & httpReq.statusText
    End If
    On Error GoTo 0
End Function

' Escapar caracteres especiais para JSON
Function EscapeJson(texto As String) As String
    Dim result As String
    result = texto
    
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\\n")
    result = Replace(result, vbCr, "\\n")
    result = Replace(result, vbLf, "\\n")
    result = Replace(result, vbTab, "\\t")
    
    EscapeJson = result
End Function

' Extrair conteudo da resposta JSON
Function ExtrairConteudoResposta(jsonResponse As String) As String
    Dim json As Object
    Dim content As String
    
    On Error Resume Next
    Set json = JsonConverter.ParseJson(jsonResponse)
    
    If Not json Is Nothing Then
        If json.Exists("choices") Then
            If IsArray(json("choices")) Or TypeName(json("choices")) = "Collection" Then
                content = json("choices")(1)("message")("content")
            End If
        End If
    End If
    
    ExtrairConteudoResposta = content
    On Error GoTo 0
End Function

' Extrair dados da resposta da IA
Sub ExtrairRespostaIA(respostaIA As String, ByRef notaQualidade As String, ByRef bulletOtimizado As String)
    Dim json As Object
    Dim startPos As Integer, endPos As Integer
    Dim jsonStr As String
    
    notaQualidade = ""
    bulletOtimizado = ""
    
    ' Tentar encontrar JSON na resposta
    startPos = InStr(respostaIA, "{")
    endPos = InStrRev(respostaIA, "}")
    
    If startPos > 0 And endPos > startPos Then
        jsonStr = Mid(respostaIA, startPos, endPos - startPos + 1)
        
        On Error Resume Next
        Set json = JsonConverter.ParseJson(jsonStr)
        
        If Not json Is Nothing Then
            If json.Exists("nota_qualidade") Then
                notaQualidade = CStr(json("nota_qualidade"))
            End If
            If json.Exists("bullet_otimizado") Then
                bulletOtimizado = CStr(json("bullet_otimizado"))
            End If
        End If
        On Error GoTo 0
    End If
    
    ' Se nao conseguiu parsear, usar resposta bruta
    If bulletOtimizado = "" Then
        bulletOtimizado = respostaIA
        notaQualidade = "N/A"
    End If
End Sub
