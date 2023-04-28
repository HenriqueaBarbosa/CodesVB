Private Sub atualizarBtn_Click()

    Dim id As Integer
    id = idTexto.Value
    
    Atualizar (id)
    
    novoBtn_Click

End Sub

Private Sub cadastrarBtn_Click()
    
    If nomeTexto.Value = "" Then
        MsgBox "Preencher campo nome é obrigatório"
        Exit Sub
    ElseIf sexoBox.Value = "" Then
        MsgBox "Preencher campo sexo é obrigatório"
        Exit Sub
    ElseIf cepTexto.Value = "" Then
        MsgBox "Preencher campo cep é obrigatório"
        Exit Sub
    ElseIf enderecoTexto.Value = "" Then
        MsgBox "Preencher campo endereco é obrigatório"
        Exit Sub
    ElseIf bairroTexto.Value = "" Then
        MsgBox "Preencher campo bairro é obrigatório"
        Exit Sub
    ElseIf cidadeTexto.Value = "" Then
        MsgBox "Preencher campo cidade é obrigatório"
        Exit Sub
    ElseIf ufTexto.Value = "" Then
        MsgBox "Preencher campo nome é obrigatório"
        Exit Sub
    End If
    
    Call Cadastrar
    
    novoBtn_Click
    
End Sub

Private Sub deletarBtn_Click()

    Dim id As Integer
    Dim resposta As VbMsgBoxResult
    
    resposta = MsgBox("Deseja deletar dados?", vbOKCancel + vbQuestion)
    
    If resposta = vbCancel Then Exit Sub
    
    id = idTexto.Value
    Deletar (id)
    
    novoBtn_Click
    
End Sub

Private Sub limparBtn_Click()

    Call LimparControles

End Sub

Private Sub novoBtn_Click()
    
    Dim id As Integer
    
    id = pBase.Range("i1").Value
    
    idTexto.Value = id
    Call LimparControles
    
    nomeTexto.SetFocus
    
    cadastrarBtn.Enabled = True
    
    atualizarBtn.Enabled = False
    deletarBtn.Enabled = False
    
End Sub

Private Sub pesquisarBtnUm_Click()

    UserForm2.Show

End Sub

Private Sub UserForm_Initialize()

    sexoBox.AddItem "Masculino"
    sexoBox.AddItem "Feminino"
    
    novoBtn_Click

End Sub

Private Sub pesquisarBtnDois_Click()

    'Em ferramentas -> referencias... -> habilitar a referencia(Microsoft XML, V6) e
    'Microsoft HTML OBJECT LIBRARY

    Dim api As New MSXML2.ServerXMLHTTP60
    Dim HTML As New HTMLDocument
    Dim cep As String
    Dim url As String
    
    On Error GoTo fim:
    
    cep = cepTexto.Value
    
    url = "https://viacep.com.br/ws/" & cep & "/xml/"
    
    api.Open "GET", url
    api.send
    
    HTML.body.innerHTML = api.responseText
    enderecoTexto.Value = HTML.getElementsByTagName("logradouro")(0).innerText
    bairroTexto.Value = HTML.getElementsByTagName("bairro")(0).innerText
    cidadeTexto.Value = HTML.getElementsByTagName("localidade")(0).innerText
    ufTexto.Value = HTML.getElementsByTagName("uf")(0).innerText
    
    Exit Sub
    
fim:
    enderecoTexto.Value = ""
    bairroTexto.Value = ""
    cidadeTexto.Value = ""
    ufTexto.Value = ""
    
    MsgBox "CEP inválido, por favor verificar se está correto ou digitar endereço manualmente", vbQuestion
    
End Sub