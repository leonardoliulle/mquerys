
// Controle de Indisponibilidade do portal da SEREDE
let
    Fonte = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\controle_solicitacoes_.xlsx"), null, true),
    controle_solicitacoes__Sheet = Fonte{[Item="controle_solicitacoes_",Kind="Sheet"]}[Data],
    #"Cabeçalhos promovidos" = Table.PromoteHeaders(controle_solicitacoes__Sheet),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos promovidos",{{"SOLICITAÇÃO", Int64.Type}, {"TR", type text}, {"NOME TÉCNICO", type text}, {"UF", type text}, {"GRA", type text}, {"SETOR", type text}, {"TIPO", type text}, {"DATA INICIO", type datetime}, {"DATA FIM", type datetime}, {"MOTIVO (JUSTIFICATIVA)", type text}, {"OBSERVAÇÕES", type text}, {"SOLICITANTE", type text}, {"TEC. SUBSTITUTO", type text}, {"NOME TEC. SUBSTITUTO", type text}, {"ABERTURA", type datetime}, {"STATUS", type text}, {"OBS DA TRAMITAÇÃO", type text}, {"NUM ARS", type any}, {"GRAM PARA", type any}, {"GRA PARA", type any}, {"SETOR PARA", type any}, {"RAIO PARA", type any}, {"SKILL PARA", type any}, {"PONTO DE PARTIDA", type any}, {"AREA DE TRABALHO", type any}, {"SOLICITACAO VINCULADA", type any}, {"SETOR VINCULADO", type any}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado",{{"SOLICITAÇÃO", "SOLICITACAO"}, {"TR", "MATRICULA TECNICO"}, {"NOME TÉCNICO", "NOME TECNICO"}, {"MOTIVO (JUSTIFICATIVA)", "MOTIVO"}, {"OBSERVAÇÕES", "OBSERVACOES"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"TEC. SUBSTITUTO", "NOME TEC. SUBSTITUTO", "OBS DA TRAMITAÇÃO", "NUM ARS", "GRAM PARA", "GRA PARA", "SETOR PARA", "RAIO PARA", "SKILL PARA", "PONTO DE PARTIDA", "AREA DE TRABALHO", "SOLICITACAO VINCULADA", "SETOR VINCULADO"})
in
    #"Colunas Removidas"


 // CONTORLE DE INDISPONIBILIDADE DO PORTAL DO R1

 let
    Fonte = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\fila_solicitacoes_R1.xlsx"), null, true),
    fila_solicitacoes__Sheet = Fonte{[Item="fila_solicitacoes_",Kind="Sheet"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(fila_solicitacoes__Sheet,{{"Column1", type any}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type any}, {"Column9", type any}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type any}, {"Column19", type any}, {"Column20", type text}}),
    #"Cabeçalhos promovidos" = Table.PromoteHeaders(#"Tipo Alterado"),
    #"Colunas Removidas" = Table.RemoveColumns(#"Cabeçalhos promovidos",{"MAT. CONFIG", "ULT. MOVIMENTACAO", "Column19", "obs Configurador", "TEC. SUBSTITUTO", "BAs AGENDADOS"})
in
    #"Colunas Removidas"


 // CONTORLE DE INDISPONIBILIDADE DO PORTAL DO R2


let
    Fonte = Excel.Workbook(File.Contents("W:\CICLO\2_REPORTS\GERAL\Bases Indisponibilidade\fila_solicitacoes_R2.xlsx"), null, true),
    fila_solicitacoes__Sheet = Fonte{[Item="fila_solicitacoes_",Kind="Sheet"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(fila_solicitacoes__Sheet,{{"Column1", type any}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type any}, {"Column9", type any}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type any}, {"Column14", type any}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type any}, {"Column19", type any}, {"Column20", type text}}),
    #"Cabeçalhos promovidos" = Table.PromoteHeaders(#"Tipo Alterado"),
    #"Colunas Removidas" = Table.RemoveColumns(#"Cabeçalhos promovidos",{"MAT. CONFIG", "ULT. MOVIMENTACAO", "Column19", "obs Configurador", "TEC. SUBSTITUTO", "BAs AGENDADOS"})
in
    #"Colunas Removidas"


 // UNION DE VARIAS CONSULTAS

 let
    Fonte = Table.Combine({controle_solicitacoes_,fila_solicitacoes_R1,fila_solicitacoes_R2}),
    #"Duplicatas Removidas" = Table.Distinct(Fonte, {"SOLICITACAO", "MATRICULA TECNICO", "TIPO", "DATA INICIO", "DATA FIM", "MOTIVO"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Duplicatas Removidas", each ([STATUS] <> "DEVOLVIDO")),
    #"Data Inserida" = Table.AddColumn(#"Linhas Filtradas", "Date", each DateTime.Date([ABERTURA]), type date),
    #"Idade Inserida" = Table.AddColumn(#"Data Inserida", "AgeFromDate", each Date.From(DateTime.LocalNow()) - [Date], type duration),
    #"Linhas Filtradas1" = Table.SelectRows(#"Idade Inserida", each ([AgeFromDate] = #duration(1, 0, 0, 0)))
in
    #"Linhas Filtradas1"