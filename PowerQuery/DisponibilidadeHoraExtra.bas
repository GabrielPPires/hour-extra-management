let
    Fonte = SharePoint.Tables("Disponibilidade-hora-extra", [Implementation=null, ApiVersion=15]),
    #"Id2" = Fonte{[Id2"]}[Items],
    #"Author Expandido" = Table.ExpandRecordColumn(#"Id2", "Author", {"Id", "Title", "Name", "EMail", "FirstName", "LastName", "UserName", "ID", "Created", "AuthorId", "EditorId", "Author"}, {"Author.Id", "Author.Title", "Author.Name", "Author.EMail", "Author.FirstName", "Author.LastName", "Author.UserName", "Author.ID.1", "Author.Created", "Author.AuthorId", "Author.EditorId", "Author.Author"}),
    #"Outras Colunas Removidas1" = Table.SelectColumns(#"Author Expandido",{"Data para realiza_x0", "Qual turno tem_x0020", "Quantas horas dispon", "Tem necessidade de_x", "Analistas", "SUPERVISORES", "LIBERAÇÃO DO_x", "COMENTÁRIO", "ID", "Created", "Author.Title", "Author.EMail"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Outras Colunas Removidas1",{{"Created", "DATA_RESPOSTA"}, {"Data para realiza_x0", "DATA"}, {"Qual turno tem_x0020", "TURNO"}, {"Quantas horas dispon", "QTDD_HORAS"}, {"Tem necessidade de_x", "TRANSPORTE"}, {"Author.Title", "NOME"}, {"Author.EMail", "EMAIL"}, {"ID", "ID_DISPONIBILIDADE"}, {"Analistas", "ANALISTAS"}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Colunas Renomeadas",{"ID_DISPONIBILIDADE", "DATA_RESPOSTA", "EMAIL", "NOME", "DATA", "TURNO", "QTDD_HORAS", "TRANSPORTE", "ANALISTAS", "LIBERAÇÃO DO_x", "SUPERVISORES", "COMENTÁRIO"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Reordenadas",{{"DATA_RESPOSTA", type datetime}, {"DATA", type date}}),
    #"Consultas Mescladas" = Table.NestedJoin(#"Tipo Alterado", {"DATA"}, #"DIA-SEMANA", {"DATA"}, "DIA-SEMANA", JoinKind.Inner),
    #"Consultas Mescladas1" = Table.NestedJoin(#"Consultas Mescladas", {"EMAIL"}, COLABORADORES, {"E-MAIL"}, "COLABORADORES", JoinKind.LeftOuter),
    #"COLABORADORES Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas1", "COLABORADORES", {"IPN"}, {"COLABORADORES.IPN"}),
    #"Colunas Removidas" = Table.RemoveColumns(#"COLABORADORES Expandido",{"DIA-SEMANA"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Colunas Removidas",{{"COLABORADORES.IPN", "IPN"}}),
    #"Colunas Reordenadas1" = Table.ReorderColumns(#"Colunas Renomeadas1",{"ID_DISPONIBILIDADE", "DATA_RESPOSTA", "EMAIL", "NOME", "IPN", "DATA", "TURNO", "QTDD_HORAS", "TRANSPORTE", "ANALISTAS", "LIBERAÇÃO DO_x", "SUPERVISORES", "COMENTÁRIO"})
in
    #"Colunas Reordenadas1"
