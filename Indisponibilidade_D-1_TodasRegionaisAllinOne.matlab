// Controle de Indisponibilidade do portal da SEREDE
let


    Fonte1 = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\controle_solicitacoes_.xlsx"), null, true),
    controle_solicitacoes__Sheet = Fonte1{[Item="controle_solicitacoes_",Kind="Sheet"]}[Data],
    #"Cabeçalhos promovidos" = Table.PromoteHeaders(controle_solicitacoes__Sheet),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Cabeçalhos promovidos",{{"SOLICITAÇÃO", "SOLICITACAO"}, {"TR", "MATRICULA TECNICO"}, {"NOME TÉCNICO", "NOME TECNICO"}, {"MOTIVO (JUSTIFICATIVA)", "MOTIVO"}, {"OBSERVAÇÕES", "OBSERVACOES"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"TEC. SUBSTITUTO", "NOME TEC. SUBSTITUTO", "OBS DA TRAMITAÇÃO", "NUM ARS", "GRAM PARA", "GRA PARA", "SETOR PARA", "RAIO PARA", "SKILL PARA", "PONTO DE PARTIDA", "AREA DE TRABALHO", "SOLICITACAO VINCULADA", "SETOR VINCULADO"}),



 // CONTORLE DE INDISPONIBILIDADE DO PORTAL DO R1


    Fonte2 = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\fila_solicitacoes_R1.xlsx"), null, true),
    fila_solicitacoes__Sheet2 = Fonte2{[Item="fila_solicitacoes_2019_07_28_18",Kind="Sheet"]}[Data],
    #"Tipo Alterado3" = Table.TransformColumnTypes(fila_solicitacoes__Sheet2,{{"Column1", type any}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type any}, {"Column9", type any}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type any}, {"Column19", type any}, {"Column20", type text}}),
    #"Cabeçalhos promovidos3" = Table.PromoteHeaders(#"Tipo Alterado3"),
    #"Colunas Removidas3" = Table.RemoveColumns(#"Cabeçalhos promovidos3",{"MAT. CONFIG", "ULT. MOVIMENTACAO", "Column19", "obs Configurador", "TEC. SUBSTITUTO", "BAs AGENDADOS"}),



 // CONTORLE DE INDISPONIBILIDADE DO PORTAL DO R2



    Fonte3 = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\fila_solicitacoes_R2.xlsx"), null, true),
    fila_solicitacoes__Sheet3 = Fonte3{[Item="fila_solicitacoes_2019_07_28_18",Kind="Sheet"]}[Data],
    #"Tipo Alterado2" = Table.TransformColumnTypes(fila_solicitacoes__Sheet3,{{"Column1", type any}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type any}, {"Column9", type any}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type any}, {"Column19", type any}, {"Column20", type text}}),
    #"Cabeçalhos promovidos2" = Table.PromoteHeaders(#"Tipo Alterado2"),
    #"Colunas Removidas2" = Table.RemoveColumns(#"Cabeçalhos promovidos2",{"MAT. CONFIG", "ULT. MOVIMENTACAO", "Column19", "obs Configurador", "TEC. SUBSTITUTO", "BAs AGENDADOS"}),



 // UNION DE VARIAS CONSULTAS

    Fonte = Table.Combine({#"Colunas Removidas2",#"Colunas Removidas3",#"Colunas Removidas"}),
    #"Duplicatas Removidas" = Table.Distinct(Fonte, {"SOLICITACAO", "MATRICULA TECNICO", "TIPO", "DATA INICIO", "DATA FIM", "MOTIVO"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Duplicatas Removidas", each ([STATUS] <> "DEVOLVIDO")),
    #"Data Inserida" = Table.AddColumn(#"Linhas Filtradas", "Date", each DateTime.Date([ABERTURA]), type date),
    #"Idade Inserida" = Table.AddColumn(#"Data Inserida", "AgeFromABERTURA", each Date.From(DateTime.LocalNow()) - [Date], type duration),
    #"Linhas Filtradas1" = Table.SelectRows(#"Idade Inserida", each ([AgeFromABERTURA] = #duration(1, 0, 0, 0))),
    #"Data Inserida1" = Table.AddColumn(#"Linhas Filtradas1", "DINI", each DateTime.Date([DATA INICIO]), type date),
    #"Data Inserida2" = Table.AddColumn(#"Data Inserida1", "DFIM", each DateTime.Date([DATA FIM]), type date),
    #"Personalização Adicionada" = Table.AddColumn(#"Data Inserida2", "Custom", each [DFIM]-[DINI]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Personalização Adicionada",{{"Custom", Int64.Type}}),
    #"Personalização Adicionada1" = Table.AddColumn(#"Tipo Alterado", "Date.1", each List.Dates([DINI], [Custom]+1, #duration(1,0,0,0))),
    #"Date.1 Expandido" = Table.ExpandListColumn(#"Personalização Adicionada1", "Date.1"),
    #"Idade Inserida1" = Table.AddColumn(#"Date.1 Expandido", "AgeFromDate", each Date.From(DateTime.LocalNow()) - [Date.1], type duration),
    #"Linhas Filtradas2" = Table.SelectRows(#"Idade Inserida1", each ([AgeFromDate] = #duration(1, 0, 0, 0)))
in
    #"Linhas Filtradas2"// Controle de Indisponibilidade do portal da SEREDE
let


    Fonte1 = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\controle_solicitacoes_.xlsx"), null, true),
    controle_solicitacoes__Sheet = Fonte1{[Item="controle_solicitacoes_",Kind="Sheet"]}[Data],
    #"Cabeçalhos promovidos" = Table.PromoteHeaders(controle_solicitacoes__Sheet),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Cabeçalhos promovidos",{{"SOLICITAÇÃO", "SOLICITACAO"}, {"TR", "MATRICULA TECNICO"}, {"NOME TÉCNICO", "NOME TECNICO"}, {"MOTIVO (JUSTIFICATIVA)", "MOTIVO"}, {"OBSERVAÇÕES", "OBSERVACOES"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"TEC. SUBSTITUTO", "NOME TEC. SUBSTITUTO", "OBS DA TRAMITAÇÃO", "NUM ARS", "GRAM PARA", "GRA PARA", "SETOR PARA", "RAIO PARA", "SKILL PARA", "PONTO DE PARTIDA", "AREA DE TRABALHO", "SOLICITACAO VINCULADA", "SETOR VINCULADO"}),



 // CONTORLE DE INDISPONIBILIDADE DO PORTAL DO R1


    Fonte2 = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\fila_solicitacoes_R1.xlsx"), null, true),
    fila_solicitacoes__Sheet2 = Fonte2{[Item="fila_solicitacoes_2019_07_28_18",Kind="Sheet"]}[Data],
    #"Tipo Alterado3" = Table.TransformColumnTypes(fila_solicitacoes__Sheet2,{{"Column1", type any}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type any}, {"Column9", type any}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type any}, {"Column19", type any}, {"Column20", type text}}),
    #"Cabeçalhos promovidos3" = Table.PromoteHeaders(#"Tipo Alterado3"),
    #"Colunas Removidas3" = Table.RemoveColumns(#"Cabeçalhos promovidos3",{"MAT. CONFIG", "ULT. MOVIMENTACAO", "Column19", "obs Configurador", "TEC. SUBSTITUTO", "BAs AGENDADOS"}),



 // CONTORLE DE INDISPONIBILIDADE DO PORTAL DO R2



    Fonte3 = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\fila_solicitacoes_R2.xlsx"), null, true),
    fila_solicitacoes__Sheet3 = Fonte3{[Item="fila_solicitacoes_2019_07_28_18",Kind="Sheet"]}[Data],
    #"Tipo Alterado2" = Table.TransformColumnTypes(fila_solicitacoes__Sheet3,{{"Column1", type any}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type any}, {"Column9", type any}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type any}, {"Column19", type any}, {"Column20", type text}}),
    #"Cabeçalhos promovidos2" = Table.PromoteHeaders(#"Tipo Alterado2"),
    #"Colunas Removidas2" = Table.RemoveColumns(#"Cabeçalhos promovidos2",{"MAT. CONFIG", "ULT. MOVIMENTACAO", "Column19", "obs Configurador", "TEC. SUBSTITUTO", "BAs AGENDADOS"}),



 // UNION DE VARIAS CONSULTAS

    Fonte = Table.Combine({#"Colunas Removidas2",#"Colunas Removidas3",#"Colunas Removidas"}),
    #"Duplicatas Removidas" = Table.Distinct(Fonte, {"SOLICITACAO", "MATRICULA TECNICO", "TIPO", "DATA INICIO", "DATA FIM", "MOTIVO"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Duplicatas Removidas", each ([STATUS] <> "DEVOLVIDO")),
    #"Data Inserida" = Table.AddColumn(#"Linhas Filtradas", "Date", each DateTime.Date([ABERTURA]), type date),
    #"Idade Inserida" = Table.AddColumn(#"Data Inserida", "AgeFromABERTURA", each Date.From(DateTime.LocalNow()) - [Date], type duration),
    #"Linhas Filtradas1" = Table.SelectRows(#"Idade Inserida", each ([AgeFromABERTURA] = #duration(1, 0, 0, 0))),
    #"Data Inserida1" = Table.AddColumn(#"Linhas Filtradas1", "DINI", each DateTime.Date([DATA INICIO]), type date),
    #"Data Inserida2" = Table.AddColumn(#"Data Inserida1", "DFIM", each DateTime.Date([DATA FIM]), type date),
    #"Personalização Adicionada" = Table.AddColumn(#"Data Inserida2", "Custom", each [DFIM]-[DINI]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Personalização Adicionada",{{"Custom", Int64.Type}}),
    #"Personalização Adicionada1" = Table.AddColumn(#"Tipo Alterado", "Date.1", each List.Dates([DINI], [Custom]+1, #duration(1,0,0,0))),
    #"Date.1 Expandido" = Table.ExpandListColumn(#"Personalização Adicionada1", "Date.1"),
    #"Idade Inserida1" = Table.AddColumn(#"Date.1 Expandido", "AgeFromDate", each Date.From(DateTime.LocalNow()) - [Date.1], type duration),
    #"Linhas Filtradas2" = Table.SelectRows(#"Idade Inserida1", each ([AgeFromDate] = #duration(1, 0, 0, 0)))
in
    #"Linhas Filtradas2"