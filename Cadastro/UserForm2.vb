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