Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

   Dim id As Integer
   
   id = ListBox1.List(ListBox1.ListIndex, 0)
   
   Pesquisar (id)
   
   Unload Me
   
   UserForm1.cadastrarBtn.Enabled = False
   
   UserForm1.atualizarBtn.Enabled = True
   UserForm1.deletarBtn.Enabled = True

End Sub

Private Sub nomeTxt_Change()

    Dim nlin As Integer
    Dim i As Integer
    Dim linha As Integer
    Dim nomeDigitado As String
    Dim nomeProcurado As String
    
    nlin = pBase.Range("A1").CurrentRegion.Rows.Count
    nomeDigitado = UCase(nomeTxt.Value)
    
    ListBox1.Clear
    
    For i = 2 To nlin
    
        nomeProcurado = UCase(pBase.Cells(i, 2).Value)
        
        If InStr(1, nomeProcurado, nomeDigitado) > 0 Then
        
        ListBox1.AddItem
        linha = ListBox1.ListCount
        
        ListBox1.List(linha - 1, 0) = pBase.Cells(i, 1).Value
        ListBox1.List(linha - 1, 1) = pBase.Cells(i, 2).Value
        
        End If
        
    Next
    
End Sub