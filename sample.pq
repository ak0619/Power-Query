// 旧売上傾向
let
    ソース = Folder.Contents("D:\★Exports\data"),
    行を選択 = Table.SelectRows(ソース, each ([Extension] = ".xlsx")),
    ブックを追加 = Table.AddColumn(行を選択, "ブック", each Excel.Workbook([Content])),
    ブックを展開 = Table.ExpandTableColumn(ブックを追加, "ブック", {"Data", "Item", "Kind"}),
    行を選択2 = Table.SelectRows(
        ブックを展開,
        each ([Item] = "大分類別集計") and ([Kind] = "Sheet")
    ),
    テーブルを追加 = Table.AddColumn(行を選択2, "テーブル", each Table.PromoteHeaders([Data])),
    テーブルを展開 = Table.ExpandTableColumn(
        テーブルを追加,
        "テーブル",
        {"月度", "店舗コード", "大分類コード", "新品売上金額", "新品粗利額", "中古売上金額", "中古粗利額"}
    ),
    列を削除 = Table.RemoveColumns(
        テーブルを展開,
        {
            "Content",
            "Name",
            "Extension",
            "Date accessed",
            "Date modified",
            "Date created",
            "Attributes",
            "Folder Path",
            "Data",
            "Item",
            "Kind"
        }
    ),
    列を分割 = Table.SplitColumn(
        列を削除,
        "月度",
        Splitter.SplitTextByRepeatedLengths(4),
        {"年度", "月度"}
    ),
    データ型を変更 = Table.TransformColumnTypes(
        列を分割,
        {
            {"年度", Int64.Type},
            {"月度", type text},
            {"店舗コード", Int64.Type},
            {"大分類コード", Int64.Type},
            {"新品売上金額", Int64.Type},
            {"新品粗利額", Int64.Type},
            {"中古売上金額", Int64.Type},
            {"中古粗利額", Int64.Type}
        }
    ),
    行を選択3 = Table.SelectRows(
        データ型を変更,
        each List.Contains(
            {3, 4, 5, 6, 8, 9, 10, 11, 12, 20, 80, 81, 83, 84, 85},
            [店舗コード]
        )
    ),
    他の列をピボット解除 = Table.UnpivotOtherColumns(
        行を選択3,
        {"年度", "月度", "店舗コード", "大分類コード"},
        "属性",
        "値"
    ),
    新品中古を追加 = Table.AddColumn(
        他の列をピボット解除,
        "新品中古",
        each if Text.StartsWith([属性], "新品") then 1 else 2,
        Int64.Type
    ),
    列の値を変換 = Table.TransformColumns(
        新品中古を追加,
        {
            {
                "属性",
                each Text.Replace(Text.Replace(_, "新品", ""), "中古", ""),
                type text
            }
        }
    ),
    列をピボット = Table.Pivot(列の値を変換, List.Distinct(列の値を変換[属性]), "属性", "値"),
    列名を変更 = Table.RenameColumns(列をピボット, {{"粗利額", "粗利金額"}})
in
    列名を変更

// 新売上傾向
let
    ソース = Folder.Contents("C:\imdb_files\データ"),
    ブックを追加 = Table.AddColumn(ソース, "ブック", each Excel.Workbook([Content])),
    ブックを展開 = Table.ExpandTableColumn(ブックを追加, "ブック", {"Data", "Item", "Kind"}),
    行を選択 = Table.SelectRows(
        ブックを展開,
        each ([Item] = "大分類別") and ([Kind] = "Sheet")
    ),
    テーブルを追加 = Table.AddColumn(行を選択, "テーブル", each Table.PromoteHeaders([Data])),
    テーブルを展開 = Table.ExpandTableColumn(
        テーブルを追加,
        "テーブル",
        {"月度", "店舗コード", "大分類", "新品中古", "売上金額", "粗利金額"}
    ),
    列を削除 = Table.RemoveColumns(
        テーブルを展開,
        {
            "Content",
            "Name",
            "Extension",
            "Date accessed",
            "Date modified",
            "Date created",
            "Attributes",
            "Folder Path",
            "Data",
            "Item",
            "Kind"
        }
    ),
    列を分割 = Table.SplitColumn(
        列を削除,
        "月度",
        Splitter.SplitTextByRepeatedLengths(4),
        {"年度", "月度"}
    ),
    データ型を変更 = Table.TransformColumnTypes(
        列を分割,
        {
            {"年度", Int64.Type},
            {"月度", type text},
            {"店舗コード", Int64.Type},
            {"大分類", Int64.Type},
            {"新品中古", Int64.Type},
            {"売上金額", Int64.Type},
            {"粗利金額", Int64.Type}
        }
    ),
    列名を変更 = Table.RenameColumns(データ型を変更, {{"大分類", "大分類コード"}}),
    行を選択3 = Table.SelectRows(
        列名を変更,
        each [大分類コード]
            <> 0
            and List.Contains(
                {3, 4, 5, 6, 8, 9, 10, 11, 12, 20, 80, 81, 83, 84, 85},
                [店舗コード]
            )
    )
in
    行を選択3

// 店舗マスタ
let
    ソース = #table(
        type table [店舗コード = Int64.Type, 店舗名 = text],
        {
            {3, "御茶ノ水本店"},
            {4, "FINEST GUITARS"},
            {5, "渋谷店"},
            {6, "新宿店"},
            {8, "池袋店"},
            {9, "横浜店"},
            {10, "名古屋栄店"},
            {11, "梅田店"},
            {12, "心斎橋店"},
            {20, "福岡店"},
            {80, "EC受注センター"},
            {81, "EC受注センター"},
            {83, "EC受注センター"},
            {84, "リユース事業部"},
            {85, "リユース事業部"}
        }
    )
in
    ソース

// 大分類マスタ
let
    ソース = #table(
        type table [大分類コード = Int64.Type, 大分類名 = text],
        {
            {11, "弦楽器エレキ"},
            {12, "弦楽器アコースティック"},
            {13, "アンプ"},
            {14, "エフェクター"},
            {15, "弦楽器アクセサリー"},
            {21, "電子ピアノ"},
            {22, "シンセ/キーボード"},
            {23, "PA/録音機器"},
            {24, "DJ/VJ/映像機器"},
            {25, "DTM/音楽制作"},
            {26, "デジタルアクセサリー"},
            {31, "管楽器"},
            {34, "管楽器アクセサリー"},
            {41, "電子ドラム"},
            {42, "生ドラム/他"},
            {43, "ドラムアクセサリー"},
            {51, "その他楽器"},
            {52, "楽譜/雑誌/教則"},
            {53, "スクール"},
            {54, "その他売上"}
        }
    )
in
    ソース

// 新旧売上傾向
let
    ソース = Table.Combine({旧売上傾向, 新売上傾向})
in
    ソース

// 第一四半期
let
    ソース = 新旧売上傾向,
    行を選択 = Table.SelectRows(
        ソース,
        each List.Contains({"06", "07", "08"}, [月度])
    ),
    行をグループ化 = Table.Group(
        行を選択,
        {"年度", "店舗コード", "大分類コード", "新品中古"},
        {
            {"月度", each "第一四半期", type text},
            {"売上金額", each List.Sum([売上金額]), Int64.Type},
            {"粗利金額", each List.Sum([粗利金額]), Int64.Type}
        }
    ),
    列を並替 = Table.ReorderColumns(
        行をグループ化,
        {"年度", "月度", "店舗コード", "大分類コード", "新品中古", "売上金額", "粗利金額"}
    )
in
    列を並替

// 半期
let
    ソース = 新旧売上傾向,
    行を選択 = Table.SelectRows(
        ソース,
        each List.Contains({"06", "07", "08", "09", "10", "11"}, [月度])
    ),
    行をグループ化 = Table.Group(
        行を選択,
        {"年度", "店舗コード", "大分類コード", "新品中古"},
        {
            {"月度", each "半期", type text},
            {"売上金額", each List.Sum([売上金額]), Int64.Type},
            {"粗利金額", each List.Sum([粗利金額]), Int64.Type}
        }
    ),
    列を並替 = Table.ReorderColumns(
        行をグループ化,
        {"年度", "月度", "店舗コード", "大分類コード", "新品中古", "売上金額", "粗利金額"}
    )
in
    列を並替

// 第三四半期（累計）
let
    ソース = 新旧売上傾向,
    行を選択 = Table.SelectRows(
        ソース,
        each List.Contains(
            {"06", "07", "08", "09", "10", "11", "12", "01", "02"},
            [月度]
        )
    ),
    行をグループ化 = Table.Group(
        行を選択,
        {"年度", "店舗コード", "大分類コード", "新品中古"},
        {
            {"月度", each "第三四半期（累計）", type text},
            {"売上金額", each List.Sum([売上金額]), Int64.Type},
            {"粗利金額", each List.Sum([粗利金額]), Int64.Type}
        }
    ),
    列を並替 = Table.ReorderColumns(
        行をグループ化,
        {"年度", "月度", "店舗コード", "大分類コード", "新品中古", "売上金額", "粗利金額"}
    )
in
    列を並替

// 通期
let
    ソース = 新旧売上傾向,
    行をグループ化 = Table.Group(
        ソース,
        {"年度", "店舗コード", "大分類コード", "新品中古"},
        {
            {"月度", each "通期", type text},
            {"売上金額", each List.Sum([売上金額]), Int64.Type},
            {"粗利金額", each List.Sum([粗利金額]), Int64.Type}
        }
    ),
    列を並替 = Table.ReorderColumns(
        行をグループ化,
        {"年度", "月度", "店舗コード", "大分類コード", "新品中古", "売上金額", "粗利金額"}
    )
in
    列を並替

// 売上傾向
let
    ソース = Table.Combine({新旧売上傾向, 第一四半期, 半期, #"第三四半期（累計）", 通期}),
    レコードを追加 = Table.AddColumn(
        ソース,
        "レコード",
        each [
            店舗名 = 店舗マスタ{[店舗コード = [店舗コード]]}[店舗名],
            大分類名 = 大分類マスタ{[大分類コード = [大分類コード]]}[大分類名]
        ]
    ),
    レコードを展開 = Table.ExpandRecordColumn(
        レコードを追加,
        "レコード",
        {"店舗名", "大分類名"},
        {"店舗名", "大分類名"}
    ),
    列を並替 = Table.ReorderColumns(
        レコードを展開,
        {"年度", "月度", "店舗コード", "店舗名", "大分類コード", "大分類名", "新品中古", "売上金額", "粗利金額"}
    )
in
    列を並替
