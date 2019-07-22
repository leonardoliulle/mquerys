let
    Fonte = Excel.Workbook(File.Contents("W:\CRM\CONFIGURACAO\1_BASES\apoio\RNE_hierarquia.xlsx"), null, true),
    BASE_Sheet = Fonte{[Item="BASE",Kind="Sheet"]}[Data],
    #"Cabeçalhos promovidos" = Table.PromoteHeaders(BASE_Sheet),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Cabeçalhos promovidos",{{"SETOR", type text}, {"COORDENADOR", type text}, {"TEL. COORD", Int64.Type}, {"GERENTE", type text}, {"TEL. GER", Int64.Type}, {"GG", type text}, {"TEL. GG", Int64.Type}, {"GAA", type text}, {"UF", type text}, {"SUPERVISOR CL", type text}, {"COORDENADOR CL", type text}, {"GRA", type text}})
in
    #"Tipo Alterado1"