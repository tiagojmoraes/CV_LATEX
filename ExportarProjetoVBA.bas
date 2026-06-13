Attribute VB_Name = "ExportarProjetoVBA"
Sub ExportarProjetoVBA()
    Dim vbaProjeto As Object
    Dim componente As Object
    Dim caminhoPasta As String
    Dim extensao As String
    
    ' Define a pasta onde os arquivos serão salvos (mesma pasta da planilha)
    caminhoPasta = ActiveWorkbook.Path & "\VBA_Codigo\"
    
    ' Cria a pasta se ela não existir
    If Dir(caminhoPasta, vbDirectory) = "" Then
        MkDir caminhoPasta
    End If
    
    Set vbaProjeto = ActiveWorkbook.VBProject
    
    ' Loop por todos os componentes do projeto VBA
    For Each componente In vbaProjeto.VBComponents
        ' Ignora planilhas e a EstaPasta_de_Trabalho vazias
        If componente.Type <> 100 Then
            Select Case componente.Type
                Case 1: extensao = ".bas" ' Módulo padrão
                Case 2: extensao = ".cls" ' Módulo de Classe
                Case 3: extensao = ".frm" ' Formulário (UserForm)
            End Select
            
            ' Exporta o arquivo para a pasta criada
            componente.Export caminhoPasta & componente.Name & extensao
        End If
    Next componente
    
    MsgBox "Todos os códigos foram exportados para: " & caminhoPasta, vbInformation
End Sub

