let
    Fonte = Excel.Workbook(File.Contents("W:\CRM\CONFIGURACAO\1_BASES\TarefasFechadasHoje_.xlsx"), null, true),
    TarefasFechadasHoje_Sheet = Fonte{[Item="TarefasFechadasHoje",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(TarefasFechadasHoje_Sheet),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"NRBA", Int64.Type}, {"ESTADO", type text}, {"INICIOAGENDAMENTO", type datetime}, {"FIMAGENDAMENTO", type datetime}, {"DATALIMITEAGENDAMENTOMANUAL", type any}, {"ODC", Int64.Type}, {"APRAZADO", Int64.Type}, {"AREA", type text}, {"AREATECNICA", type text}, {"AREATRABALHO", type text}, {"ATIVIDADE", type text}, {"BLOQUEADO", Int64.Type}, {"CABO", type text}, {"CAIXA", type text}, {"CODENCERRAMENTO", type text}, {"CPF", Int64.Type}, {"CONTATO1", type text}, {"CONTATO2", type text}, {"CONTATO3", type text}, {"CONTATOEFETIVO", type text}, {"DATAABERTURAOS", type datetime}, {"DATAABERTURABA", type datetime}, {"DATAPROMESSA", type datetime}, {"DATAINDICADOR", type datetime}, {"INICIOEXECUCAO", type datetime}, {"FIMEXECUCAO", type datetime}, {"DURACAO", Int64.Type}, {"EMPRESA", type text}, {"ERROINTEGRACAO", type any}, {"ESTACAO", type text}, {"FILAPOSTO", type text}, {"GERENCIA", type text}, {"IDDADOSREDE", type text}, {"INSTALACAOUNICA", Int64.Type}, {"VENDACONJUNTA", Int64.Type}, {"INSTUNICAREALIZADA", Int64.Type}, {"LATITUDE", Int64.Type}, {"LOCALIDADE", type any}, {"LOCALIDADEREMOTA", Int64.Type}, {"LONGITUDE", Int64.Type}, {"MACROATIVIDADE", type text}, {"MOTIVOREPARO", type any}, {"NOMECLIENTE", type text}, {"CIRCUITO", type text}, {"DOCUMENTOASSOCIADO", type text}, {"NUMEROFISICO", Int64.Type}, {"OBSERVACOES", type text}, {"OPERADORAGENDAMENTO", type text}, {"PORTABILIDADE", Int64.Type}, {"POSSUIVELOX", Int64.Type}, {"PRIORIDADE", Int64.Type}, {"PRIORIDADE OP.", type text}, {"PSUP", Int64.Type}, {"QTDDPONTOS", Int64.Type}, {"RECLAMACAO", type text}, {"REGIAO", type text}, {"REINCIDENCIA", Int64.Type}, {"REPAROGARANTIA", Int64.Type}, {"SISTEMAABERTURA", type text}, {"SKILLS", type text}, {"TECNICO", type text}, {"TELEFONEBUSCA", Int64.Type}, {"TERMINAL", Int64.Type}, {"BINAGEMEXECUTADA", Int64.Type}, {"SETOR", type text}, {"TIPOCLIENTE", type text}, {"UF", type text}, {"USARIOCIENTEUP", type any}, {"CODIGOROTEAMENTO", type text}, {"PRONTOPARAEXECUCAO", type text}, {"FLAGAGENDAMENTO", type text}, {"GRAM", type text}, {"GRA", type text}, {"MATRICULATECNICO", type text}, {"VELOCIDADECONTRATADA", type text}, {"TESTEFINAL", type text}, {"TIPODEENCERRAMENTO", type text}, {"SERVICO", type text}, {"TIPOTERMINAL", type text}, {"MERCADO", type text}}),
    #"Outras Colunas Removidas" = Table.SelectColumns(#"Tipo Alterado",{"NRBA", "ESTADO", "INICIOAGENDAMENTO", "FIMAGENDAMENTO", "DATALIMITEAGENDAMENTOMANUAL", "AREA", "AREATECNICA", "AREATRABALHO", "ATIVIDADE", "BLOQUEADO", "CODENCERRAMENTO", "DATAABERTURAOS", "DATAABERTURABA", "INICIOEXECUCAO", "FIMEXECUCAO", "DURACAO", "EMPRESA", "FILAPOSTO", "LOCALIDADE", "LOCALIDADEREMOTA", "MACROATIVIDADE", "DOCUMENTOASSOCIADO", "PRIORIDADE", "PRIORIDADE OP.", "QTDDPONTOS", "RECLAMACAO", "REINCIDENCIA", "REPAROGARANTIA", "SKILLS", "TECNICO", "TERMINAL", "SETOR", "TIPOCLIENTE", "UF", "CODIGOROTEAMENTO", "PRONTOPARAEXECUCAO", "FLAGAGENDAMENTO", "GRAM", "GRA", "MATRICULATECNICO", "VELOCIDADECONTRATADA", "MERCADO"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Outras Colunas Removidas", each [UF] = "PREVLAT" and [UF] = "CE" or [UF] = "PB" or [UF] = "PE" or [UF] = "RN" or [UF] = "AL" or [UF] = "BA" or [UF] = "SE"),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "PESO", each if [ESTADO] = "Concluído com sucesso" then [DURACAO]/35.5/60 else null),
    #"Consultas Mescladas" = Table.NestedJoin(#"Personalização Adicionada",{"MATRICULATECNICO"},TC,{"MATRÍCULA"},"TC
",JoinKind.LeftOuter),
    #"TC#(lf) Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "TC#(lf)", {"INDISPONIBILIDADE", "INICIOINDISP", "FIMINDISP", "PERFILTECNICO"}, {"TC#(lf).INDISPONIBILIDADE", "TC#(lf).INICIOINDISP", "TC#(lf).FIMINDISP", "TC#(lf).PERFILTECNICO"}),
    #"Consultas Mescladas1" = Table.NestedJoin(#"TC#(lf) Expandido",{"MATRICULATECNICO"},QtqFechou,{"MATRICULATECNICO"},"QT"),
    #"QT Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas1", "QT", {"QtLinha"}, {"QTBAIXAS"}),
    #"Personalização Adicionada1" = Table.AddColumn(#"QT Expandido", "FRACAO.TC", each 1/[QTBAIXAS]),
    #"Coluna Duplicada" = Table.DuplicateColumn(#"Personalização Adicionada1", "SETOR", "SETOR - Copiar"),
    #"Dividir Coluna pelo Delimitador" = Table.SplitColumn(#"Coluna Duplicada","SETOR - Copiar",Splitter.SplitTextByDelimiter("."),{"SETOR - Copiar.1", "SETOR - Copiar.2", "SETOR - Copiar.3", "SETOR - Copiar.4"}),
    #"Tipo Alterado10" = Table.TransformColumnTypes(#"Dividir Coluna pelo Delimitador",{{"SETOR - Copiar.1", type text}, {"SETOR - Copiar.2", type text}, {"SETOR - Copiar.3", type text}, {"SETOR - Copiar.4", Int64.Type}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado10",{{"GRA", "GRAA"}, {"GRAM", "GRAMA"}, {"SETOR - Copiar.2", "GRAM"}, {"SETOR - Copiar.3", "GRA"}}),
    #"Consultas Mescladas2" = Table.NestedJoin(#"Colunas Renomeadas",{"GRA"},hieraquia,{"GRA"},"HC"),
    #"HC Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas2", "HC", {"REG.GERENTE", "GERENTE"}, {"HC.REG.GERENTE", "HC.GERENTE"}),
    #"Personalização Adicionada2" = Table.AddColumn(#"HC Expandido", "ComSucesso", each if [ESTADO] = "Concluído com sucesso" then 1 else null),
    #"Personalização Adicionada3" = Table.AddColumn(#"Personalização Adicionada2", "SemSucesso", each if [ESTADO] = "Concluído sem sucesso" then 1 else null),
    #"Personalização Adicionada4" = Table.AddColumn(#"Personalização Adicionada3", "HojeRef", each DateTime.LocalNow()),
    #"Personalização Adicionada5" = Table.AddColumn(#"Personalização Adicionada4", "DIFOS", each [HojeRef]-[DATAABERTURAOS]),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Personalização Adicionada5",{{"DIFOS", type number}}),
    #"Coluna Condicional Adicionada" = Table.AddColumn(#"Tipo Alterado1", "Faixa", each if [DIFOS] < 1 then "Entre 0 e 24h" else if [DIFOS] < 2 then "Entre 24h e 48h" else if [DIFOS] < 3 then "Entre 48h  e 72h" else if [DIFOS] < 4 then "Entre 72h e 96h" else if [DIFOS] >= 4 then "Maior que 96h" else null ),
    #"Personalização Adicionada6" = Table.AddColumn(#"Coluna Condicional Adicionada", "ComSucessoREP", each if [ESTADO] = "Concluído com sucesso" and [MACROATIVIDADE] = "REP" then 1 else 0),
    #"Personalização Adicionada7" = Table.AddColumn(#"Personalização Adicionada6", "SemSucessoREP", each if [ESTADO] = "Concluído sem sucesso" and [MACROATIVIDADE] = "REP" then 1 else 0),
    #"Personalização Adicionada8" = Table.AddColumn(#"Personalização Adicionada7", "ComSucessoREP96h", each if [ESTADO] = "Concluído com sucesso" and [MACROATIVIDADE] = "REP" and [DIFOS] > 3 then 1 else 0),
    #"Personalização Adicionada9" = Table.AddColumn(#"Personalização Adicionada8", "SemSucessoREP96h", each if [ESTADO] = "Concluído sem sucesso" and [MACROATIVIDADE] = "REP" and [DIFOS] > 3 then 1 else 0),
    #"Coluna Duplicada2" = Table.DuplicateColumn(#"Personalização Adicionada9", "CODENCERRAMENTO", "codbaixa"),
    #"Tipo Alterado20" = Table.TransformColumnTypes(#"Coluna Duplicada2",{{"codbaixa", type number}}),
#"Erros Substituídos" = Table.ReplaceErrorValues(#"Tipo Alterado20", {{"codbaixa", -1}}),
    #"Consultas Mescladas3" = Table.NestedJoin(#"Erros Substituídos",{"ESTADO"},#"tbl_codbaixas (2)",{"NOVO_ESTADO"},"BX"),
    #"BX Expandido1" = Table.ExpandTableColumn(#"Consultas Mescladas3", "BX", {"DESC_FINAL", "NOVO_ESTADO", "MOTIVO_NOVO", "FINAL_MACRO"}, {"BX.DESC_FINAL", "BX.NOVO_ESTADO", "BX.MOTIVO_NOVO", "BX.FINAL_MACRO"}),
    #"Duplicatas Removidas" = Table.Distinct(#"BX Expandido1", {"NRBA"}),
    #"Consultas Mescladas4" = Table.NestedJoin(#"Duplicatas Removidas",{"SETOR"},RNE_hierarquia,{"SETOR"},"NewColumn",JoinKind.LeftOuter),
    #"NewColumn Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas4", "NewColumn", {"COORDENADOR", "GERENTE", "GG"}, {"COORDENADOR", "GERENTE", "GG"}),
    #"Consultas mescladas" = Table.NestedJoin(#"NewColumn Expandido",{"SETOR"},RNE_hierarquia,{"SETOR"},"NewColumn"),
    #"NewColumn Expandido1" = Table.ExpandTableColumn(#"Consultas mescladas", "NewColumn", {"REGIONAL"}, {"NewColumn.REGIONAL"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"NewColumn Expandido1",{{"NewColumn.REGIONAL", "REGIONAL"}})
in
    #"Colunas Renomeadas1"