let
    Fonte = Excel.Workbook(File.Contents("W:\CRM\CONFIGURACAO\1_BASES\apoio\GA_RV.1.xlsx"), null, true),
    Base_Sheet = Fonte{[Item="Base",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Base_Sheet),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"UF", type text}, {"CPF", type any}, {"COORDENADOR", type text}, {"GRA", type text}, {"SETOR", type text}, {"SEGMENTO", type text}, {"CONFIRMADO_COM_OPERAÇÃO", type text}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Tipo Alterado",{"UF", "SETOR", "COORDENADOR", "CPF", "GRA", "SEGMENTO", "CONFIRMADO_COM_OPERAÇÃO"})
in
    #"Colunas Reordenadas"