Sub CriarApresentacao()
    Dim Apresentacao As Object
    Dim Slide As Object
    
    ' Criar uma nova apresentação
    Set Apresentacao = CreateObject("PowerPoint.Application").Presentations.Add
    
    ' Slide 1: Introdução
    Set Slide = Apresentacao.Slides.Add(1, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Agenda Telefônica"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Bem-vindo à Agenda Telefônica!"
    
    ' Slide 2: Opção 1 - Cadastrar Contato
    Set Slide = Apresentacao.Slides.Add(2, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Opção 1 - Cadastrar Contato"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Essa opção permite cadastrar um novo contato na agenda."
    
    ' Slide 3: Opção 2 - Buscar Contato
    Set Slide = Apresentacao.Slides.Add(3, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Opção 2 - Buscar Contato"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Essa opção permite buscar um contato na agenda pelo nome."
    
    ' Slide 4: Opção 3 - Editar Contato
    Set Slide = Apresentacao.Slides.Add(4, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Opção 3 - Editar Contato"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Essa opção permite editar as informações de um contato existente na agenda."
    
    ' Slide 5: Opção 4 - Excluir Contato
    Set Slide = Apresentacao.Slides.Add(5, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Opção 4 - Excluir Contato"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Essa opção permite excluir um contato da agenda."
    
    ' Slide 6: Opção 5 - Exportar Agenda
    Set Slide = Apresentacao.Slides.Add(6, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Opção 5 - Exportar Agenda"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Essa opção permite exportar a agenda para um arquivo de texto."
    
    ' Slide 7: Opção 6 - Importar Agenda
    Set Slide = Apresentacao.Slides.Add(7, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Opção 6 - Importar Agenda"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Essa opção permite importar contatos de um arquivo de texto para a agenda."
    
    ' Slide 8: Opção 0 - Sair
    Set Slide = Apresentacao.Slides.Add(8, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Opção 0 - Sair"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Essa opção permite sair do programa."
    
    ' Slide 9: Conclusão
    Set Slide = Apresentacao.Slides.Add(9, 1)
    Slide.Shapes.Title.TextFrame.TextRange.Text = "Conclusão"
    Slide.Shapes.AddTextbox(1, 100, 150, 500, 200).TextFrame.TextRange.Text = "Obrigado por usar a Agenda Telefônica!"
    
    ' Salvar a apresentação como um arquivo PPTX
    Apresentacao.SaveAs "C:\Caminho\para\a\Apresentacao.pptx"
    
    ' Fechar a aplicação do PowerPoint
    Apresentacao.Close
    Set Apresentacao = Nothing
    Set Slide = Nothing
    
    MsgBox "Apresentação criada com sucesso!"
End Sub
