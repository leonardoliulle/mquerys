let
    Fonte = Csv.Document(File.Contents("W:\CRM\CONFIGURACAO\1_BASES\TarefasAbertas_.csv"),[Delimiter="|",Encoding=65001]),
    #"Cabeçalhos promovidos" = Table.PromoteHeaders(Fonte),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos promovidos",{{"""NRBA""", Int64.Type}, {"ESTADO", type text}, {"INICIOAGENDAMENTO", type datetime}, {"FIMAGENDAMENTO", type datetime}, {"DATALIMITEAGENDAMENTOMANUAL", type datetime}, {"ODC", Int64.Type}, {"APRAZADO", type text}, {"AREA", type text}, {"AREATECNICA", type text}, {"AREATRABALHO", type text}, {"SERVICO", type text}, {"ATIVIDADE", type text}, {"BLOQUEADO", type text}, {"CABO", type text}, {"CAIXA", type text}, {"CODENCERRAMENTO", type text}, {"CPF", Int64.Type}, {"CONTATO1", type text}, {"CONTATO2", type text}, {"CONTATO3", type text}, {"CONTATOEFETIVO", type text}, {"DATAABERTURAOS", type datetime}, {"DATAABERTURABA", type datetime}, {"DATAPROMESSA", type datetime}, {"DATAINDICADOR", type datetime}, {"INICIOEXECUCAO", type datetime}, {"FIMEXECUCAO", type datetime}, {"DURACAO", Int64.Type}, {"EMPRESA", type text}, {"ERROINTEGRACAO", type text}, {"ESTACAO", type text}, {"FILAPOSTO", type text}, {"GERENCIA", type text}, {"IDDADOSREDE", type text}, {"INSTALACAOUNICA", type text}, {"VENDACONJUNTA", type text}, {"INSTUNICAREALIZADA", type text}, {"LATITUDE", Int64.Type}, {"LOCALIDADE", type text}, {"LOCALIDADEREMOTA", Int64.Type}, {"LONGITUDE", Int64.Type}, {"MACROATIVIDADE", type text}, {"MOTIVOREPARO", type text}, {"NOMECLIENTE", type text}, {"CIRCUITO", type text}, {"DOCUMENTOASSOCIADO", type text}, {"NUMEROFISICO", Int64.Type}, {"OBSERVACOES", type text}, {"OPERADORAGENDAMENTO", type text}, {"PORTABILIDADE", type text}, {"POSSUIVELOX", type text}, {"PRIORIDADE", Int64.Type}, {"Prioridade op.", type text}, {"PSUP", type text}, {"QTDDPONTOS", Int64.Type}, {"RECLAMACAO", type text}, {"REGIAO", type text}, {"REINCIDENCIA", Int64.Type}, {"REPAROGARANTIA", type text}, {"SISTEMAABERTURA", type text}, {"SKILLS", type text}, {"TECNICO", type text}, {"TELEFONEBUSCA", Int64.Type}, {"TERMINAL", type text}, {"BINAGEMEXECUTADA", type text}, {"SETOR", type text}, {"TIPOCLIENTE", type text}, {"UF", type text}, {"USARIOCIENTEUP", type text}, {"PRONTOPARAEXECUCAO", type text}, {"FLAGAGENDAMENTO", type text}, {"GRAM", type text}, {"GRA", type text}, {"MATRICULATECNICO", type text}, {"TIPOCOORDENADAS", type text}, {"CODIGOROTEAMENTO", type text}, {"DISTANCIACALCULADA", type text}, {"DISTANCIAPERMITIDA", Int64.Type}, {"TIPODETERMINAL", type text}, {"MERCADO", type text}, {"VELOCIDADECONTRATADA", type text}, {"TESTEFINAL", type text}, {"ORIGEM", type text}, {"RELATEDORDER", type text}, {"CONTRATO", Int64.Type}, {"BUNDLEID", type text}, {"MANUAL", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each [SKILLS] <> "CRTRCX" and [SKILLS] <> "PREVARD" and [SKILLS] <> "PREVCB" and [SKILLS] <> "PREVCX" and [SKILLS] <> "PREVLAT" and [UF] = "CE" or [UF] = "PB" or [UF] = "PE" or [UF] = "RN" or [UF] = "AL" or [UF] = "BA" or [UF] = "SE"),
    #"Outras Colunas Removidas" = Table.SelectColumns(#"Linhas Filtradas",{"""NRBA""", "ESTADO", "INICIOAGENDAMENTO", "FIMAGENDAMENTO", "AREATECNICA", "AREATRABALHO", "SERVICO", "ATIVIDADE", "BLOQUEADO", "CAIXA", "CONTATO1", "CONTATO2", "CONTATO3", "CONTATOEFETIVO", "DATAABERTURAOS", "DATAABERTURABA", "DATAPROMESSA", "DATAINDICADOR", "INICIOEXECUCAO", "FIMEXECUCAO", "DURACAO", "ESTACAO", "FILAPOSTO", "LATITUDE", "LONGITUDE", "NOMECLIENTE", "CIRCUITO", "DOCUMENTOASSOCIADO", "PRIORIDADE", "Prioridade op.", "QTDDPONTOS", "RECLAMACAO", "REINCIDENCIA", "REPAROGARANTIA", "SKILLS", "TECNICO", "TERMINAL", "SETOR", "TIPOCLIENTE", "UF", "PRONTOPARAEXECUCAO", "FLAGAGENDAMENTO", "MATRICULATECNICO", "TIPOCOORDENADAS", "CODIGOROTEAMENTO", "MERCADO", "VELOCIDADECONTRATADA", "TESTEFINAL", "ORIGEM", "RELATEDORDER", "CONTRATO"}),
    #"Personalização Adicionada" = Table.AddColumn(#"Outras Colunas Removidas", "HojeRef", each DateTime.LocalNow()),
    #"Personalização Adicionada1" = Table.AddColumn(#"Personalização Adicionada", "Custom", each [INICIOEXECUCAO]-[HojeRef]),

	
	
#"Tipo Alterado1" = Table.TransformColumnTypes(#"Personalização Adicionada1",{{"Custom", type number}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado1",{{"Custom", "DifInic"}}),
    #"Personalização Adicionada2" = Table.AddColumn(#"Colunas Renomeadas", "AtribComPeso", each [DURACAO]/35.5/60),
    #"Personalização Adicionada3" = Table.AddColumn(#"Personalização Adicionada2", "DifAberturaOS", each [HojeRef]-[DATAABERTURAOS]),
    #"Tipo Alterado2" = Table.TransformColumnTypes(#"Personalização Adicionada3",{{"DifAberturaOS", type number}}),
    #"Coluna Condicional Adicionada" = Table.AddColumn(#"Tipo Alterado2", "Faixa", each if [DifAberturaOS] < 1 then "1. Menor que 24h" else if [DifAberturaOS] < 2 then "2. Entre 24h e 48h" else if [DifAberturaOS] < 3 then "3. Entre 48h e 72h" else if [DifAberturaOS] < 4 then "4. Entre 72h e 96h" else if [DifAberturaOS] < 5.8 then "5. Entre 96h e 140h" else "6. Maior 140h" ),
    #"Coluna Duplicada" = Table.DuplicateColumn(#"Coluna Condicional Adicionada", "INICIOEXECUCAO", "INICIOEXECUCAO - Copiar"),
    #"Coluna Duplicada1" = Table.DuplicateColumn(#"Coluna Duplicada", "HojeRef", "HojeRef - Copiar"),
    #"Dia Inserido" = Table.AddColumn(#"Coluna Duplicada1", "Day", each Date.Day([#"INICIOEXECUCAO - Copiar"]), type number),
    #"Colunas Removidas" = Table.RemoveColumns(#"Dia Inserido",{"Day"}),
    #"Data Inserida" = Table.AddColumn(#"Colunas Removidas", "Date", each DateTime.Date([#"INICIOEXECUCAO - Copiar"]), type date),
    #"Data Inserida1" = Table.AddColumn(#"Data Inserida", "Date.1", each DateTime.Date([#"HojeRef - Copiar"]), type date),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Data Inserida1",{{"Date.1", "HojeRefe"}, {"Date", "IniAtrib"}}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Colunas Renomeadas1",{"INICIOEXECUCAO - Copiar", "HojeRef - Copiar"}),
    #"Personalização Adicionada4" = Table.AddColumn(#"Colunas Removidas1", "Custom", each [IniAtrib]-[HojeRefe]),
    #"Tipo Alterado3" = Table.TransformColumnTypes(#"Personalização Adicionada4",{{"Custom", type number}}),
    #"Colunas Renomeadas2" = Table.RenameColumns(#"Tipo Alterado3",{{"Custom", "DDifAtrib"}}),
    #"Coluna Condicional Adicionada1" = Table.AddColumn(#"Colunas Renomeadas2", "DifAtrib", each if [DDifAtrib] = null then "Não Atribuído" else if [DDifAtrib] < 0 then "Vencido" else if [DDifAtrib] < 1 then "Atribuído Hoje" else if [DDifAtrib] < 2 then "Amanhã" else if [DDifAtrib] >= 2 then "Futuro" else null ),
    #"Outras Colunas Removidas1" = Table.SelectColumns(#"Coluna Condicional Adicionada1",{"""NRBA""", "ESTADO", "INICIOAGENDAMENTO", "FIMAGENDAMENTO", "AREATRABALHO", "SERVICO", "ATIVIDADE", "BLOQUEADO", "CAIXA", "CONTATO1", "CONTATO2", "CONTATO3", "CONTATOEFETIVO", "DATAABERTURAOS", "DATAABERTURABA", "DATAPROMESSA", "DATAINDICADOR", "INICIOEXECUCAO", "FIMEXECUCAO", "DURACAO", "ESTACAO", "FILAPOSTO", "LATITUDE", "LONGITUDE", "NOMECLIENTE", "CIRCUITO", "DOCUMENTOASSOCIADO", "PRIORIDADE", "Prioridade op.", "QTDDPONTOS", "RECLAMACAO", "REINCIDENCIA", "REPAROGARANTIA", "SKILLS", "TECNICO", "TERMINAL", "SETOR", "UF", "PRONTOPARAEXECUCAO", "FLAGAGENDAMENTO", "MATRICULATECNICO", "TIPOCOORDENADAS", "CODIGOROTEAMENTO", "MERCADO", "VELOCIDADECONTRATADA", "CONTRATO", "HojeRef", "AtribComPeso", "Faixa", "IniAtrib", "HojeRefe", "DDifAtrib", "DifAtrib"}),
    #"Valor Substituído" = Table.ReplaceValue(#"Outras Colunas Removidas1","Sim","AG",Replacer.ReplaceText,{"FLAGAGENDAMENTO"}),
    #"Valor Substituído1" = Table.ReplaceValue(#"Valor Substituído","Não","NAG",Replacer.ReplaceText,{"FLAGAGENDAMENTO"}),
    #"Coluna Duplicada2" = Table.DuplicateColumn(#"Valor Substituído1", "SETOR", "SETOR - Copiar"),
    #"Colunas Renomeadas3" = Table.RenameColumns(#"Coluna Duplicada2",{{"SETOR - Copiar", "GRA"}}),
    #"Dividir Coluna pelo Delimitador" = Table.SplitColumn(#"Colunas Renomeadas3","GRA",Splitter.SplitTextByDelimiter("."),{"GRA.1", "GRA.2", "GRA.3", "GRA.4"}),
#"Tipo Alterado4" = Table.TransformColumnTypes(#"Dividir Coluna pelo Delimitador",{{"GRA.1", type text}, {"GRA.2", type text}, {"GRA.3", type text}, {"GRA.4", Int64.Type}}),
    #"Colunas Removidas2" = Table.RemoveColumns(#"Tipo Alterado4",{"GRA.1", "GRA.2", "GRA.4"}),
    #"Colunas Renomeadas4" = Table.RenameColumns(#"Colunas Removidas2",{{"GRA.3", "GRA"}}),
    #"Personalização Adicionada5" = Table.AddColumn(#"Colunas Renomeadas4", "B2", each if [MERCADO] = "EMPRESARIAL" then "B2B" else if [MERCADO] = "CORPORATIVO" then "B2B" else "B2C"),
    #"Personalização Adicionada6" = Table.AddColumn(#"Personalização Adicionada5", "DifPROMESSA", each [DATAPROMESSA]-DateTime.LocalNow()),
    #"Coluna Duplicada3" = Table.DuplicateColumn(#"Personalização Adicionada6", "TERMINAL", "CN"),
    #"Dividir Coluna pela Posição" = Table.SplitColumn(#"Coluna Duplicada3","CN",Splitter.SplitTextByPositions({0, 2}, false),{"CN.1", "CN.2"}),
    #"Tipo Alterado6" = Table.TransformColumnTypes(#"Dividir Coluna pela Posição",{{"CN.1", Int64.Type}, {"CN.2", Int64.Type}}),
    #"Colunas Removidas3" = Table.RemoveColumns(#"Tipo Alterado6",{"CN.2"}),
    #"Colunas Renomeadas5" = Table.RenameColumns(#"Colunas Removidas3",{{"CN.1", "CN"}}),
    #"Hora Inserida" = Table.AddColumn(#"Colunas Renomeadas5", "Hour", each Time.Hour([INICIOEXECUCAO]), type number),
    #"Personalização Adicionada8" = Table.AddColumn(#"Hora Inserida", "Turno", each if [Hour] > 12 then "Tarde" else "Manhã"),
    #"Personalização Adicionada9" = Table.AddColumn(#"Personalização Adicionada8", "Prioridadeop2", each if Text.Contains([#"Prioridade op."], "Priorit") or Text.Contains([#"Prioridade op."], "VIP") or Text.Contains([#"Prioridade op."], "Ouvidoria") then "Cliente Prioritário" else if Text.Contains([#"Prioridade op."], "Anatel") or Text.Contains([#"Prioridade op."], "JEC") or Text.Contains([#"Prioridade op."], "Procon") then "Orgãos Oficiais" else "Sem prioridade"),
    #"Duplicatas Removidas" = Table.Distinct(#"Personalização Adicionada9", {"""NRBA"""}),
    #"Personalização Adicionada7" = Table.AddColumn(#"Duplicatas Removidas", "DifPROMESSA.1", each [DATAABERTURAOS]+0.33333333-DateTime.LocalNow()),
    #"Tipo Alterado5" = Table.TransformColumnTypes(#"Personalização Adicionada7",{{"DifPROMESSA", type number}}),
    #"Personalização Adicionada10" = Table.AddColumn(#"Tipo Alterado5", "TUP ART-Prazo", each if [DifPROMESSA] < 0 then "Vencido" else if [DifPROMESSA] < 2 then "Menor que 2h pra vencer" else "Dentro do Prazo")
in
    #"Personalização Adicionada10"