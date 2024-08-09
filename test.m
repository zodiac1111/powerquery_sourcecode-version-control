//######Reconcile Sheet A## This is delimiter. Dont remove it
let
    Source = Excel.Workbook(File.Contents(pFileA), null, true),
    #"223 form_Sheet" = Source{[Item=pSheetA,Kind="Sheet"]}[Data],
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

//######Reconcile Sheet B## This is delimiter. Dont remove it
let
    Source = Excel.Workbook(File.Contents(pFileB), null, true),
    #"223 form_Sheet" = Source{[Item=pSheetB,Kind="Sheet"]}[Data],
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

//######Compare Worksheets## This is delimiter. Dont remove it
let
    Source = Table.NestedJoin(#"Reconcile Sheet A",{"Row", "Column"},#"Reconcile Sheet B",{"Row", "Column"},"Reconcile Sheet B",JoinKind.FullOuter),
    #"Expanded Reconcile Sheet B" = Table.ExpandTableColumn(Source, "Reconcile Sheet B", {"Row", "Column", "Value"}, {"Reconcile Sheet B.Row", "Reconcile Sheet B.Column", "Reconcile Sheet B.Value"}),
    #"Added Conditional Column" = Table.AddColumn(#"Expanded Reconcile Sheet B", "Match", each if [Value] = [Reconcile Sheet B.Value] then "Match" else "Non Match" ),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([Match] = "Non Match"))
in
    #"Filtered Rows"

//######Matches## This is delimiter. Dont remove it
let
    Source = Table.NestedJoin(#"Reconcile Sheet A",{"Row", "Column"},#"Reconcile Sheet B",{"Row", "Column"},"Reconcile Sheet B",JoinKind.FullOuter),
    #"Expanded Reconcile Sheet B" = Table.ExpandTableColumn(Source, "Reconcile Sheet B", {"Row", "Column", "Value"}, {"Reconcile Sheet B.Row", "Reconcile Sheet B.Column", "Reconcile Sheet B.Value"}),
    #"Added Conditional Column" = Table.AddColumn(#"Expanded Reconcile Sheet B", "Match", each if [Value] = [Reconcile Sheet B.Value] then "Match" else "Non Match" ),
    #"Grouped Rows" = Table.Group(#"Added Conditional Column", {"Match"}, {{"Row Count", each Table.RowCount(_), type number}})
in
    #"Grouped Rows"

//######pFileA## This is delimiter. Dont remove it
let
    Source = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Parameter Name", type text}, {"Parameter Value", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Parameter Name] = "Workbook A Path")),
    #"Parameter Value" = #"Filtered Rows"{0}[Parameter Value]
in
    #"Parameter Value"

//######pFileB## This is delimiter. Dont remove it
let
    Source = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Parameter Name", type text}, {"Parameter Value", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Parameter Name] = "Workbook B Path")),
    #"Parameter Value" = #"Filtered Rows"{0}[Parameter Value]
in
    #"Parameter Value"

//######pSheetA## This is delimiter. Dont remove it
let
    Source = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Parameter Name", type text}, {"Parameter Value", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Parameter Name] = "Worksheet A")),
    #"Parameter Value" = #"Filtered Rows"{0}[Parameter Value]
in
    #"Parameter Value"

//######pSheetB## This is delimiter. Dont remove it
let
    Source = Excel.CurrentWorkbook(){[Name="Parameters"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Parameter Name", type text}, {"Parameter Value", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Parameter Name] = "Worksheet B")),
    #"Parameter Value" = #"Filtered Rows"{0}[Parameter Value]
in
    #"Parameter Value"

//######ProcessWorksheetTemplate## This is delimiter. Dont remove it
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

//######pWorksheet## This is delimiter. Dont remove it
"Data" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

//######pWorkbookPath## This is delimiter. Dont remove it
"T:\_CMBI_Training\Datasets\File Compare\Financial Sample Before.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

//######fuProcessWorksheet## This is delimiter. Dont remove it
let
    Source = (pWorksheet as text, pWorkbookPath as text) => let
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
in
    Source

//######Workbook A Data## This is delimiter. Dont remove it
let
    Source = Excel.Workbook(File.Contents(pFileA), null, true),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Kind] = "Sheet")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Item", "Kind", "Data"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "WorkbookPath", each pFileA),
    #"Invoked Custom Function" = Table.AddColumn(#"Added Custom", "Worksheet Data", each fuProcessWorksheet([Name], [WorkbookPath])),
    #"Expanded Worksheet Data" = Table.ExpandTableColumn(#"Invoked Custom Function", "Worksheet Data", {"Row", "Column", "Value"}, {"Row", "Column", "Value"})
in
    #"Expanded Worksheet Data"

//######Workbook B Data## This is delimiter. Dont remove it
let
    Source = Excel.Workbook(File.Contents(pFileB), null, true),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Kind] = "Sheet")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Item", "Kind", "Data"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "WorkbookPath", each pFileB),
    #"Invoked Custom Function" = Table.AddColumn(#"Added Custom", "Worksheet Data", each fuProcessWorksheet([Name], [WorkbookPath])),
    #"Expanded Worksheet Data" = Table.ExpandTableColumn(#"Invoked Custom Function", "Worksheet Data", {"Row", "Column", "Value"}, {"Row", "Column", "Value"})
in
    #"Expanded Worksheet Data"

//######Workbook Compare## This is delimiter. Dont remove it
let
    Source = Table.NestedJoin(#"Workbook A Data", {"Name", "Row", "Column"}, #"Workbook B Data", {"Name", "Row", "Column"}, "Workbook B Data", JoinKind.FullOuter),
    #"Expanded Workbook B Data" = Table.ExpandTableColumn(Source, "Workbook B Data", {"Name", "Hidden", "WorkbookPath", "Row", "Column", "Value"}, {"Workbook B Data.Name", "Workbook B Data.Hidden", "Workbook B Data.WorkbookPath", "Workbook B Data.Row", "Workbook B Data.Column", "Workbook B Data.Value"}),
    #"Added Conditional Column" = Table.AddColumn(#"Expanded Workbook B Data", "Match", each if [Value] = [Workbook B Data.Value] then "Match" else "Non Match")
in
    #"Added Conditional Column"

//######Workbook Difference Summary## This is delimiter. Dont remove it
let
    Source = #"Workbook Compare",
    #"Grouped Rows" = Table.Group(Source, {"Name", "Match", "Workbook B Data.Name"}, {{"Cell Count", each Table.RowCount(_), type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows",{{"Name", "Worksheet"}})
in
    #"Renamed Columns"

//######Workbook Difference Detail## This is delimiter. Dont remove it
let
    Source = #"Workbook Compare",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Match] = "Non Match")),
    #"Reordered Columns" = Table.ReorderColumns(#"Filtered Rows",{"Name", "Hidden", "WorkbookPath", "Row", "Column", "Value", "Workbook B Data.Value", "Workbook B Data.Name", "Workbook B Data.Hidden", "Workbook B Data.WorkbookPath", "Workbook B Data.Row", "Workbook B Data.Column", "Match"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns",{{"Value", "Workbook A Data Value"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns",{"Name", "Hidden", "Row", "Column", "Workbook A Data Value", "Workbook B Data.Value", "Workbook B Data.Name", "Workbook B Data.Hidden", "WorkbookPath", "Workbook B Data.WorkbookPath", "Workbook B Data.Row", "Workbook B Data.Column", "Match"})
in
    #"Reordered Columns1"

//######LibPQ## This is delimiter. Dont remove it
/**
LibPQ:
    Access Power Query functions and queries stored in source code modules
    on filesystem or on the Web.

Project website:
    https://libpq.ml

This code was last modified on 2019-10-11

Timestamp in docstring is necessary for further updating, because this code
will be copied into the workbook and will be managed manually afterwards.
**/

let
    /* Read LibPQ settings */
    Sources.Local = LibPQPath[Local],
    Sources.Web   = LibPQPath[Web],

    /* Constants */
    EXTENSION = ".pq",
    PATHSEPLOCAL = Text.Start("\\",1),
    PATHSEPREMOTE = "/",
    ERR_SOURCE_UNREADABLE = "LibPQ.ReadError",
    DESCRIPTION_FOOTER = (path) =>
        "<br><br>" &
        "<i><div>" &
        "This module was loaded with LibPQ: https://github.com/sio/LibPQ" &
        "</div><div>" &
        "Module source code: " &
        path &
        "</div></i>",

    /* Load text content from local file or from web */
    Read.Text = (destination as text, optional local as logical) =>
        let
            Local = if local is null then true else local,
            Fetcher = if Local then File.Contents else Web.Contents
        in
            Text.FromBinary(
                Binary.Buffer(
                    try
                        Fetcher(destination)
                    otherwise
                        error Error.Record(
                            ERR_SOURCE_UNREADABLE,
                            "Read.Text: can not fetch from destination",
                            destination
                        )
                )
            ),

    /*
    Read the first multiline comment from the source code in Power Query
    Formula language (also known as M language). That comment is considered a
    docstring for LibPQ
    */
    Read.Docstring = (source_code as text) =>
    let
        Docstring = [
            start = "/*",
            end = "*/"
        ],
        DocstringDirty = Text.BeforeDelimiter(
            source_code,
            Docstring[end]
        ),
        BeforeDocstring = Text.BeforeDelimiter(
            DocstringDirty,
            Docstring[start]
        ),
        MustBeEmpty = Text.Trim(BeforeDocstring),
        DocstringText =
            if
                Text.Length(MustBeEmpty) = 0
            then
                Text.Trim(DocstringDirty, {"*", "/", " ", "#(cr)", "#(lf)", "#(tab)"})
            else
                ""
    in
        DocstringText,

    /*
       Load Power Query function or module from file,
       return null if destination unreadable
    */
    Module.FromPath = (path as text, optional local as logical, optional name as text) =>
        let
            SourceCode = try Read.Text(path, local)
                         otherwise "null",
            LoadedObject = Expression.Evaluate(SourceCode, #shared),
            LoadTry = try LoadedObject,
            CustomError = Record.TransformFields(
                LoadTry[Error],
                {"Detail", each CustomErrorDetail}
            ),
            CustomErrorDetail = [
                Original = LoadTry[Error][Detail],
                LibPQ = ExtraMetadata
            ],
            ExtraMetadata = [
                LibPQ.Module = name,
                LibPQ.Source = path,
                LibPQ.Docstring = Read.Docstring(SourceCode),
                Documentation.Name = LibPQ.Module,
                Documentation.Description =
                    Text.Replace(LibPQ.Docstring, "#(lf)", "<br>") &
                    DESCRIPTION_FOOTER(path)
            ],
            OldMetadata = Value.Metadata(LoadedObject),
            TypeMetadata = Value.Metadata(Value.Type(LoadedObject)),
            Module = Value.ReplaceType(
                try LoadedObject otherwise error CustomError,
                Value.ReplaceMetadata(
                    Value.Type(LoadedObject),
                    Record.Combine({ExtraMetadata, TypeMetadata})
                )
            ) meta Record.Combine({ExtraMetadata, OldMetadata})
        in
            Module,

    /* Calculate where the function code is located */
    Module.BuildPath = (funcname as text, directory as text, optional local as logical) =>
        let
            /* Defaults */
            Local = if local is null then true else local,
            PathSep = if Local then PATHSEPLOCAL else PATHSEPREMOTE,

            /* Path building */
            ProperDir = if Text.EndsWith(directory, PathSep)
                        then directory
                        else directory & PathSep,
            ProperName = Module.NameToProper(funcname),
            Return = ProperDir & ProperName & EXTENSION
        in
            Return,

    /* Module name converters */
    Module.NameToProper = (name) => Text.Replace(name, "_", "."),
    Module.NameFromProper = (name) => Text.Replace(name, ".", "_"),

    /* Find all modules in the list of directories */
    Module.Explore = (directories as list) =>
        let
            Files = List.Generate(
                () => [i = -1, results = 0],
                each [i] < List.Count(directories),
                each [
                    i = [i]+1,
                    folder = directories{i},
                    iserror = (try Table.RowCount(
                                    Folder.Contents(folder)
                               ))[HasError],  // For some weird reason try does
                                              // not catch DataSource error.
                                              // Check "try Folder.Contents("C:\none")"
                                              // it will return [HasError]=false
                    files = if iserror then
                                #table({"Name","Extension"},{})
                            else
                                Folder.Contents(folder),
                    filter = Table.SelectRows(
                                files,
                                each [Extension] = EXTENSION
                            ),
                    results = Table.RowCount(filter),
                    module = List.Transform(
                                    filter[Name],
                                    each Text.BeforeDelimiter(
                                        _,
                                        EXTENSION,
                                        {0,RelativePosition.FromEnd}
                                    )
                                )
                ],
                each [
                    folder = [folder],
                    module = [module],
                    results = [results]
                ]
            ),
            Return = try
                        Table.ExpandListColumn(
                            Table.FromRecords(
                                List.Select(Files, each [results]>0)
                            ),
                            "module"
                        )
                     otherwise
                        #table({"folder", "module", "results"},{})
        in
            Return,

    /* Import module (first match) from the list of possible locations */
    Module.ImportAny = (name as text, locations as list, optional local as logical) =>
        let
            Paths = List.Transform(
                        locations,
                        each Module.BuildPath(name, _, local)
                    ),
            Loop = List.Generate(
                () => [
                    i = -1,
                    object = null,
                    lasterror = null
                ],
                each [i] < List.Count(Paths),
                each [
                    // `load` should be evaluated only if absolutely necessary.
                    // If path is unreadable, no error is raised but null value
                    // is returned (see Module.FromPath)
                    load = try Module.FromPath(Paths{i}, local, name),
                    object = if [object] is null and not load[HasError]
                             then load[Value]
                             else [object],
                    lasterror = if [object] is null and load[HasError]
                                then load[Error]
                                else [lasterror],
                    i = [i] + 1
                ]
            ),
            Return = try
                        List.Select(Loop, each [object] <> null){0}[object]
                     otherwise
                        error if List.Last(Loop)[lasterror] <> null
                              then List.Last(Loop)[lasterror]
                              else Error.Record(
                                    ERR_SOURCE_UNREADABLE,
                                    "Module not found: " & name
                                   )
        in
            Return,

    /* Import a module from default locations (LibPQPath) */
    Module.Import = (name as text) =>
        let
            Attempts = {
                // {expression, silent errors}
                {Record.Field(#shared, Module.NameFromProper(name)), true},
                {Record.Field(#shared, Module.NameToProper(name)), true},
                {Record.Field(Helpers, name), true},
                {Module.ImportAny(name, Sources.Local), false},
                {Module.ImportAny(name, Sources.Web, false), false}
            },
            Results = List.Last(List.Generate(
                () => [
                    i = -1,
                    module = null,
                    error = null
                ],
                each [i] < List.Count(Attempts),
                each [
                    i = [i] + 1,
                    load = try Attempts{i}{0},
                    error = if
                                [module] is null and
                                not Attempts{i}{1} and
                                load[HasError] and
                                load[Error][Reason] <> ERR_SOURCE_UNREADABLE
                            then
                                load[Error]
                            else
                                [error],
                    module = if
                                [module] is null and
                                not load[HasError]
                             then
                                load[Value]
                             else
                                [module]
                ]
            )),
            Module = if Results[module] <> null
                     then Results[module]
                     else if Results[module] is null and Results[error] <> null
                     then error Results[error]
                     else error Error.Record(
                        ERR_SOURCE_UNREADABLE,
                        "Module not found: " & name
                     )
        in
            Module,

    /* Last touch: export helper functions defined above */
    Helpers = [
        Read.Text = Read.Text,
        Read.Docstring = Read.Docstring,
        Module.FromPath = Module.FromPath,
        Module.BuildPath = Module.BuildPath,
        Module.NameToProper = Module.NameToProper,
        Module.NameFromProper = Module.NameFromProper,
        Module.Explore = Module.Explore,
        Module.ImportAny = Module.ImportAny
    ],
    Library.Names = List.Distinct(Module.Explore(Sources.Local)[module]),
    Library = List.Last(
        List.Generate(
            () => [i=-1,record=[]],
            each [i] < List.Count(Library.Names),
            each [
                i = [i] + 1,
                record = Record.AddField(
                    [record],
                    Library.Names{i},
                    let
                        Try = try Module.Import(Library.Names{i}),
                        Return = if Try[HasError] then Try[Error] else Try[Value]
                    in
                        Return
                )
            ],
            each [record]
        )
    ),
    Module.Library = Record.Combine({Helpers, Library}),

    /* Main function */
    Main = (optional modulename as nullable text) =>
        if modulename is null
        or modulename = "Module.Library"
        then
            Module.Library
        else if modulename = "Module.Import"
        then
            Module.Import
        else
            Module.Import(modulename)
in
    Main

//######LibPQPath## This is delimiter. Dont remove it
/**
LibPQPath:
    Places where LibPQ will search for modules.

    Local sources have priority over Web sources.
    Local and Web locations must be listed in the priority order: from
    high to lower priority.  111222
**/

[
    Local = {
        "C:\ExtractLocation\LibPQ\Modules",
        "C:\ExtractLocation\LibPQ\Tests",
        "D:\ExtractLocation\LibPQ\Modules",
        "D:\ExtractLocation\LibPQ\Tests"
    },
    Web = {
        "https://github.com/sio/LibPQ/raw/master/Modules/",
        "https://github.com/tycho01/pquery/raw/master/"
    }
]

