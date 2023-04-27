Sub Exibir()

    UserForm1.Show

End Sub

Sub LimparControles()

    UserForm1.nomeTexto.Value = ""
    UserForm1.sexoBox.Value = ""
    UserForm1.cepTexto.Value = ""
    UserForm1.enderecoTexto.Value = ""
    UserForm1.bairroTexto.Value = ""
    UserForm1.cidadeTexto.Value = ""
    UserForm1.ufTexto.Value = ""

End Sub

Sub Cadastrar()

    Dim nlin As Integer
    Dim id As Integer
    
    nlin = pBase.Range("A1").CurrentRegion.Rows.Count + 1
    id = pBase.Range("i1").Value
    
    pBase.Cells(nlin, 1).Value = id
    pBase.Cells(nlin, 2).Value = UserForm1.nomeTexto.Value
    pBase.Cells(nlin, 3).Value = UserForm1.sexoBox.Value
    pBase.Cells(nlin, 4).Value = UserForm1.cepTexto.Value
    pBase.Cells(nlin, 5).Value = UserForm1.enderecoTexto.Value
    pBase.Cells(nlin, 6).Value = UserForm1.bairroTexto.Value
    pBase.Cells(nlin, 7).Value = UserForm1.cidadeTexto.Value
    pBase.Cells(nlin, 8).Value = UserForm1.ufTexto.Value
    
    pBase.Range("i1").Value = id + 1
    
    MsgBox "Dados cadastrados com sucesso"

End Sub

Sub Pesquisar(id As Integer)
    
    Dim linha As Integer
       
    linha = pBase.Columns(1).Find(id, , , xlWhole).Row
    
    UserForm1.idTexto.Value = pBase.Cells(linha, 1).Value
    UserForm1.nomeTexto.Value = pBase.Cells(linha, 2).Value
    UserForm1.sexoBox.Value = pBase.Cells(linha, 3).Value
    UserForm1.cepTexto.Value = pBase.Cells(linha, 4).Value
    UserForm1.enderecoTexto.Value = pBase.Cells(linha, 5).Value
    UserForm1.bairroTexto.Value = pBase.Cells(linha, 6).Value
    UserForm1.cidadeTexto.Value = pBase.Cells(linha, 7).Value
    UserForm1.ufTexto.Value = pBase.Cells(linha, 8).Value
        
End Sub

Sub Atualizar(id As Integer)
    
    Dim linha As Integer
       
    linha = pBase.Columns(1).Find(id, , , xlWhole).Row
    
    pBase.Cells(linha, 1).Value = id
    pBase.Cells(linha, 2).Value = UserForm1.nomeTexto.Value
    pBase.Cells(linha, 3).Value = UserForm1.sexoBox.Value
    pBase.Cells(linha, 4).Value = UserForm1.cepTexto.Value
    pBase.Cells(linha, 5).Value = UserForm1.enderecoTexto.Value
    pBase.Cells(linha, 6).Value = UserForm1.bairroTexto.Value
    pBase.Cells(linha, 7).Value = UserForm1.cidadeTexto.Value
    pBase.Cells(linha, 8).Value = UserForm1.ufTexto.Value
    
    
    MsgBox "Dados atualizados com sucesso"

    
End Sub

Sub Deletar(id As Integer)

    Dim linha As Integer
    
    linha = pBase.Columns(1).Find(id, , , xlWhole).Row
    pBase.Range("A" & linha & ":H" & linha).Delete xlUp
    
    MsgBox "Dados deletados"

End Sub
