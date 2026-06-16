Attribute VB_Name = "GerarLaTex"
Option Explicit

' ============================================================================
' Módulo: GerarLaTex
' Finalidade: Gerar arquivo LaTeX de currículo a partir dos bullets gerados
' Autor: Assistente IA
' Data: 2026
' ============================================================================

Private Const MARKER_EXPERIENCIAS As String = "<!-- EXPERIENCIAS_AQUI -->"
Private Const MARKER_FORMACAO As String = "<!-- FORMACAO_AQUI -->"
Private Const MARKER_IDIOMAS As String = "<!-- IDIOMAS_AQUI -->"
Private Const MARKER_HABILIDADES As String = "<!-- HABILIDADES_AQUI -->"

Sub GerarCurriculoLaTeX()
    On Error GoTo ErrorHandler
    
    Dim wsBullets As Worksheet
    Dim jsonRoot As Object
    Dim jsonConfig As Object
    Dim jsonContexto As Object
    Dim jsonProfissional As Object
    Dim jsonExperiencias As Object
    Dim jsonFormacao As Object
    Dim jsonCompetencias As Object
    Dim idiomaSaida As String
    Dim templatePath As String
    Dim templateContent As String
    Dim outputContent As String
    Dim outputPath As String
    Dim experienciasLatex As String
    Dim formacaoLatex As String
    Dim habilidadesLatex As String
    Dim empIndex As Long, cargoIndex As Long
    Dim i As Long
    Dim cargos As Object
    Dim hardSkills As Variant
    
    Set wsBullets = ThisWorkbook.Sheets("Bullets_Gerados")
    
    Set jsonRoot = CarregarJSONHistorico()
    If jsonRoot Is Nothing Then
        MsgBox "Erro ao carregar JSON.", vbCritical
        Exit Sub
    End If
    
    Set jsonConfig = jsonRoot("config")
    idiomaSaida = jsonConfig("idioma_saida")
    
    If LCase(idiomaSaida) = "en-us" Or LCase(idiomaSaida) = "en" Then
        templatePath = ThisWorkbook.Path & "\templates\template_en.tex"
    Else
        templatePath = ThisWorkbook.Path & "\templates\template_pt.tex"
    End If
    
    If Dir(templatePath) = "" Then
        MsgBox "Template nao encontrado: " & templatePath, vbCritical
        Exit Sub
    End If
    
    templateContent = LerArquivoTexto(templatePath)
    
    Set jsonContexto = jsonRoot("contexto")
    Set jsonProfissional = jsonContexto("profissional")
    Set jsonExperiencias = jsonProfissional("candidato_experiencia")
    
    experienciasLatex = ""
    
    For empIndex = 0 To jsonExperiencias.Count - 1
        Dim empresaNome As String
        empresaNome = jsonExperiencias(empIndex)("nome_empresa")
        Set cargos = jsonExperiencias(empIndex)("cargos")
        
        For cargoIndex = 0 To cargos.Count - 1
            Dim cargoNome As String
            Dim periodoInicio As String
            Dim periodoFim As String
            Dim localizacao As String
            
            cargoNome = cargos(cargoIndex)("cargo_nome")
            periodoInicio = cargos(cargoIndex)("periodo_cargo")("inicio")
            periodoFim = cargos(cargoIndex)("periodo_cargo")("fim")
            If cargos(cargoIndex)("periodo_cargo")("em_andamento") = True Then
                periodoFim = "Atual"
            End If
            
            localizacao = ObterLocalizacaoEmpresa(empresaNome)
            
            experienciasLatex = experienciasLatex & GerarBlocoExperienciaLatex( _
                cargoNome, empresaNome, localizacao, periodoInicio, periodoFim, wsBullets)
        Next cargoIndex
    Next empIndex
    
    Set jsonFormacao = jsonProfissional("candidato_formacao")
    formacaoLatex = ""
    
    For i = 0 To jsonFormacao.Count - 1
        Dim nomeCurso As String
        Dim nomeInstituicao As String
        Dim inicioFormacao As String
        Dim fimFormacao As String
        
        nomeCurso = jsonFormacao(i)("nome_curso")
        nomeInstituicao = jsonFormacao(i)("nome_instituicao")
        inicioFormacao = jsonFormacao(i)("mes_ano_inicio")
        fimFormacao = jsonFormacao(i)("mes_ano_fim")
        
        formacaoLatex = formacaoLatex & GerarBlocoFormacaoLatex( _
            nomeCurso, nomeInstituicao, inicioFormacao, fimFormacao)
    Next i
    
    Set jsonCompetencias = jsonProfissional("candidato_competencias")
    hardSkills = jsonCompetencias("hard_skills")
    habilidadesLatex = Join(hardSkills, " $|$ ")
    
    outputContent = templateContent
    outputContent = Replace(outputContent, MARKER_EXPERIENCIAS, experienciasLatex)
    outputContent = Replace(outputContent, MARKER_FORMACAO, formacaoLatex)
    outputContent = Replace(outputContent, MARKER_IDIOMAS, "")
    outputContent = Replace(outputContent, MARKER_HABILIDADES, habilidadesLatex)
    
    outputPath = ThisWorkbook.Path & "\curriculo_" & LCase(idiomaSaida) & ".tex"
    
    SalvarArquivoTexto outputPath, outputContent
    
    MsgBox "Curriculo LaTeX gerado com sucesso!" & vbCrLf & "Arquivo: " & outputPath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao gerar curriculo LaTeX: " & Err.Description, vbCritical
End Sub

Function CarregarJSONHistorico() As Object
    On Error Resume Next
    
    Dim wsJson As Worksheet
    Dim jsonText As String
    Dim filePath As String
    
    For Each wsJson In ThisWorkbook.Worksheets
        If InStr(1, wsJson.Name, "historico", vbTextCompare) > 0 Then
            jsonText = wsJson.Range("A1").Value
            Exit For
        End If
    Next wsJson
    
    If jsonText = "" Then
        filePath = ThisWorkbook.Path & "\historico_profissional.json"
        If Dir(filePath) <> "" Then
            jsonText = LerArquivoTexto(filePath)
        End If
    End If
    
    If jsonText = "" Then
        Set CarregarJSONHistorico = Nothing
        Exit Function
    End If
    
    Set CarregarJSONHistorico = JsonConverter.ParseJson(jsonText)
    
    On Error GoTo 0
End Function

Function GerarBlocoExperienciaLatex( _
    ByVal cargoNome As String, _
    ByVal empresaNome As String, _
    ByVal localizacao As String, _
    ByVal periodoInicio As String, _
    ByVal periodoFim As String, _
    ByRef wsBullets As Worksheet) As String
    
    Dim result As String
    Dim bullets As String
    Dim lastRow As Long
    Dim i As Long
    Dim bulletText As String
    Dim foundBullet As Boolean
    Dim periodoFormatado As String
    
    bullets = ""
    foundBullet = False
    
    lastRow = wsBullets.Cells(wsBullets.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
        If LCase(Trim(wsBullets.Cells(i, "B").Value)) = LCase(Trim(cargoNome)) Then
            bulletText = wsBullets.Cells(i, "C").Value
            If bulletText <> "" Then
                bulletText = EscapeLaTeX(bulletText)
                bullets = bullets & "        \item " & bulletText & vbCrLf
                foundBullet = True
            End If
        End If
    Next i
    
    If Not foundBullet Then
        bullets = "        \item Responsabilidades relacionadas a " & EscapeLaTeX(cargoNome) & "." & vbCrLf
    End If
    
    periodoFormatado = periodoInicio & " - " & periodoFim
    
    result = "    \item[] % Empty item to control spacing" & vbCrLf
    result = result & "    \begin{minipage}[t]{\textwidth}" & vbCrLf
    result = result & "        \textbf{\Large " & EscapeLaTeX(cargoNome) & "} \, | \, \textbf{" & EscapeLaTeX(empresaNome) & " - " & EscapeLaTeX(localizacao) & "} \hfill \textbf{\small " & periodoFormatado & "}" & vbCrLf
    result = result & "    \end{minipage}" & vbCrLf
    result = result & "    \begin{itemize}" & vbCrLf
    result = result & bullets
    result = result & "    \end{itemize}" & vbCrLf
    result = result & vbCrLf
    
    GerarBlocoExperienciaLatex = result
End Function

Function GerarBlocoFormacaoLatex( _
    ByVal nomeCurso As String, _
    ByVal nomeInstituicao As String, _
    ByVal inicioFormacao As String, _
    ByVal fimFormacao As String) As String
    
    Dim result As String
    Dim periodoFormatado As String
    
    periodoFormatado = inicioFormacao & " - " & fimFormacao
    
    result = "    \item[] % Empty item to control spacing" & vbCrLf
    result = result & "    \begin{minipage}[t]{\textwidth}" & vbCrLf
    result = result & "        \textbf{" & EscapeLaTeX(nomeCurso) & "} \\" & vbCrLf
    result = result & "        \textbf{" & EscapeLaTeX(nomeInstituicao) & "} {\color{mainblue}$\bullet$} " & periodoFormatado & vbCrLf
    result = result & "    \end{minipage}" & vbCrLf
    result = result & vbCrLf
    result = result & "    \vspace{0.5em} % Adding separation between entries" & vbCrLf
    result = result & vbCrLf
    
    GerarBlocoFormacaoLatex = result
End Function

Function EscapeLaTeX(ByVal texto As String) As String
    Dim result As String
    result = texto
    
    result = Replace(result, "\", "\textbackslash{}")
    result = Replace(result, "&", "\&")
    result = Replace(result, "%", "\%")
    result = Replace(result, "$", "\$")
    result = Replace(result, "#", "\#")
    result = Replace(result, "_", "\_")
    result = Replace(result, "{", "\{")
    result = Replace(result, "}", "\}")
    result = Replace(result, "~", "\textasciitilde{}")
    result = Replace(result, "^", "\textasciicircum{}")
    
    EscapeLaTeX = result
End Function

Function ObterLocalizacaoEmpresa(ByVal empresaNome As String) As String
    Dim localizacao As String
    
    Select Case LCase(empresaNome)
        Case "agrobiotech agro.", "agrobiotech agronegocio"
            localizacao = "Ribeirao Preto/SP"
        Case "grupo btz"
            localizacao = "Londrina/PR"
        Case "elevadores atlas schindler"
            localizacao = "Londrina/PR"
        Case "greenwich internacional"
            localizacao = "Sao Paulo/SP"
        Case Else
            localizacao = "Brasil"
    End Select
    
    ObterLocalizacaoEmpresa = localizacao
End Function

Function LerArquivoTexto(ByVal filePath As String) As String
    On Error Resume Next
    
    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -1)
    
    If Err.Number <> 0 Then
        LerArquivoTexto = ""
        Exit Function
    End If
    
    LerArquivoTexto = ts.ReadAll
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
    
    On Error GoTo 0
End Function

Sub SalvarArquivoTexto(ByVal filePath As String, ByVal content As String)
    On Error Resume Next
    
    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    
    If Err.Number <> 0 Then
        MsgBox "Erro ao salvar arquivo: " & Err.Description, vbCritical
        Exit Sub
    End If
    
    ts.Write content
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
    
    On Error GoTo 0
End Sub
