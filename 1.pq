let
    Source = Excel.Workbook(File.Contents(pWorkbookPath), null, true),
    #"223 form_Sheet" = Source{[Item=pWorksheet,Kind="Sheet"]}[Data],
    #"Added Index" = Table.AddIndexColumn(#"223 form_Sheet", "Index", 1, 1),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Added Index", {"Index"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Index", "Row"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns","Column","",Replacer.ReplaceText,{"Attribute"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"Attribute", Int64.Type}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type1",{{"Attribute", "Column"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns1",{{"Value", type text}}),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Changed Type", {{"Value", "*** CELL ERROR  ***"}})
in
    #"Replaced Errors"