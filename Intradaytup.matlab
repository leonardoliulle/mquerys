let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela8"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"REGIONAL", type text}, {"UF", type text}, {"TERMINAL", type text}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "Marcado", each "Marcado")
in
    #"Personalização Adicionada"