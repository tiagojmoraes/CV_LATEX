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
           "1. Revise as sugestoes de melhoria na coluna E" & vbCrLf & _
           "2. Aplique as melhorias e preencha a coluna F" & vbCrLf & _
           "3. Avalie a qualidade (1-10) na coluna G" & vbCrLf & _
           "4. Marque as acoes implementadas na coluna H", vbInformation, "Sucesso!"
End Sub
