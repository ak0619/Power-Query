// 接続文字列
"Provider = Microsoft.ACE.OLEDB.16.0; Data Source = C:\imprg_settings\database\財務利用者限定ＤＢ.accdb"
    meta [
        IsParameterQuery = true,
        Type = "Text",
        IsParameterQueryRequired = true
    ]

// POS_現金
let
    ソース = OleDb.Query(接続文字列, "SELECT * FROM dbo_APV_RF_SPOS_SALES"),
    必要列を選択 = Table.SelectColumns(ソース, {"取引ＩＤ", "取引区分", "内現金支払金額", "レシート番号"}),
    必要行を選択 = Table.SelectRows(
        必要列を選択,
        each ([取引区分] = "1") and ([内現金支払金額] <> "0")
    ),
    レコード列を追加 = Table.AddColumn(
        必要行を選択,
        "レコード",
        each [
            セクション = "スマレジ",
            支払方法 = "現金",
            支払番号 = "",
            支払額 = Text.Replace([内現金支払金額], "-", "")
        ]
    ),
    レコード列を展開 = Table.ExpandRecordColumn(
        レコード列を追加,
        "レコード",
        {"セクション", "支払方法", "支払番号", "支払額"}
    ),
    不要列の削除 = Table.RemoveColumns(レコード列を展開, {"取引区分", "内現金支払金額"}),
    列名を変更 = Table.RenameColumns(不要列の削除, {{"取引ＩＤ", "伝票番号"}, {"レシート番号", "受注番号"}}),
    列順を並替 = Table.ReorderColumns(
        列名を変更,
        {"伝票番号", "セクション", "支払方法", "支払番号", "支払額", "受注番号"}
    )
in
    列順を並替

// POS_ポイント
let
    ソース = OleDb.Query(接続文字列, "SELECT * FROM dbo_APV_RF_SPOS_SALES"),
    必要列を選択 = Table.SelectColumns(ソース, {"取引ＩＤ", "取引区分", "ポイント値引き", "レシート番号"}),
    必要行を選択 = Table.SelectRows(
        必要列を選択,
        each ([取引区分] = "1") and ([ポイント値引き] <> "0")
    ),
    レコード列を追加 = Table.AddColumn(
        必要行を選択,
        "レコード",
        each [
            セクション = "スマレジ",
            支払方法 = "IMCポイント",
            支払番号 = "",
            支払額 = Text.Replace([ポイント値引き], "-", "")
        ]
    ),
    レコード列を展開 = Table.ExpandRecordColumn(
        レコード列を追加,
        "レコード",
        {"セクション", "支払方法", "支払番号", "支払額"}
    ),
    不要列の削除 = Table.RemoveColumns(レコード列を展開, {"取引区分", "ポイント値引き"}),
    列名を変更 = Table.RenameColumns(不要列の削除, {{"取引ＩＤ", "伝票番号"}, {"レシート番号", "受注番号"}}),
    列順を並替 = Table.ReorderColumns(
        列名を変更,
        {"伝票番号", "セクション", "支払方法", "支払番号", "支払額", "受注番号"}
    )
in
    列順を並替

// POS_クレジット
let
    ソース = OleDb.Query(接続文字列, "SELECT * FROM dbo_APV_RF_SPOS_SALES"),
    必要列を選択 = Table.SelectColumns(
        ソース,
        {"取引ＩＤ", "取引区分", "内クレジット支払金額", "伝票番号", "取扱カード会社", "レシート番号"}
    ),
    必要行を選択 = Table.SelectRows(
        必要列を選択,
        each ([取引区分] = "1") and ([内クレジット支払金額] <> "0")
    ),
    レコード列を追加 = Table.AddColumn(
        必要行を選択,
        "レコード",
        each [
            セクション = "スマレジ",
            支払方法 =
                if [取扱カード会社] = " " then
                    "クレジット手入力"
                else if [取扱カード会社] = "Transportation system IC" then
                    "交通系IC"
                else
                    [取扱カード会社],
            支払番号 = if [伝票番号] = " " then "" else Text.PadStart([伝票番号], 5, "0"),
            支払額 = Text.Replace([内クレジット支払金額], "-", "")
        ]
    ),
    レコード列を展開 = Table.ExpandRecordColumn(
        レコード列を追加,
        "レコード",
        {"セクション", "支払方法", "支払番号", "支払額"}
    ),
    不要列の削除 = Table.RemoveColumns(
        レコード列を展開,
        {"取引区分", "内クレジット支払金額", "伝票番号", "取扱カード会社"}
    ),
    列名を変更 = Table.RenameColumns(不要列の削除, {{"取引ＩＤ", "伝票番号"}, {"レシート番号", "受注番号"}}),
    列順を並替 = Table.ReorderColumns(
        列名を変更,
        {"伝票番号", "セクション", "支払方法", "支払番号", "支払額", "受注番号"}
    )
in
    列順を並替

// POS_他支払
let
    ソース = OleDb.Query(
        接続文字列,
        Text.Combine(
            {
                "   SELECT *",
                "     FROM dbo_APV_RF_SPOS_SALES_PAYMENT AS a",
                "LEFT JOIN dbo_APV_RF_SPOS_SALES AS b",
                "       ON a.送受信管理番号 = b.送受信管理番号",
                "      AND a.送受信管理行番号 = b.送受信管理行番号"
            },
            "#(lf)"
        )
    ),
    必要列を選択 = Table.SelectColumns(
        ソース,
        {"a.取引ＩＤ", "支払方法名", "預かり金その他", "メモ", "レシート番号"}
    ),
    必要行を選択 = Table.SelectRows(必要列を選択, each ([預かり金その他] <> "0")),
    レコード列を追加 = Table.AddColumn(
        必要行を選択,
        "レコード",
        each [
            セクション = "スマレジ",
            支払方法 =
                if Text.StartsWith([支払方法名], "IMCポイント") then
                    "IMCポイント"
                else
                    [支払方法名],
            支払番号 = Text.Trim(Text.Clean([メモ])),
            支払額 = Text.Replace([預かり金その他], "-", "")
        ]
    ),
    レコード列を展開 = Table.ExpandRecordColumn(
        レコード列を追加,
        "レコード",
        {"セクション", "支払方法", "支払番号", "支払額"}
    ),
    不要列の削除 = Table.RemoveColumns(レコード列を展開, {"支払方法名", "預かり金その他", "メモ"}),
    列名を変更 = Table.RenameColumns(
        不要列の削除,
        {{"a.取引ＩＤ", "伝票番号"}, {"レシート番号", "受注番号"}}
    ),
    列順を並替 = Table.ReorderColumns(
        列名を変更,
        {"伝票番号", "セクション", "支払方法", "支払番号", "支払額", "受注番号"}
    )
in
    列順を並替

// POS_前受金
let
    ソース = OleDb.Query(
        接続文字列,
        Text.Combine(
            {
                "   SELECT *",
                "     FROM dbo_APV_RF_SPOS_SALES AS a",
                "LEFT JOIN dbo_APV_RF_SPOS_SALES_PAYMENT AS b",
                "       ON a.送受信管理番号 = b.送受信管理番号",
                "      AND a.送受信管理行番号 = b.送受信管理行番号",
                "    WHERE b.送受信管理番号 IS NULL"
            },
            "#(lf)"
        )
    ),
    必要列を選択 = Table.SelectColumns(
        ソース,
        {"a.取引ＩＤ", "取引区分", "合計", "内現金支払金額", "内クレジット支払金額", "預かり金", "レシート番号"}
    ),
    必要行を選択 = Table.SelectRows(
        必要列を選択,
        each ([取引区分] = "1")
            and ([合計] <> "0")
            and ([内現金支払金額] = "0")
            and ([内クレジット支払金額] = "0")
    ),
    レコード列を追加 = Table.AddColumn(
        必要行を選択,
        "レコード",
        each [
            セクション = "スマレジ",
            支払方法 = "前受金",
            支払番号 = "",
            支払額 = Text.Replace([預かり金], "-", "")
        ]
    ),
    レコード列を展開 = Table.ExpandRecordColumn(
        レコード列を追加,
        "レコード",
        {"セクション", "支払方法", "支払番号", "支払額"}
    ),
    不要列の削除 = Table.RemoveColumns(
        レコード列を展開,
        {"取引区分", "合計", "内現金支払金額", "内クレジット支払金額", "預かり金"}
    ),
    列名を変更 = Table.RenameColumns(
        不要列の削除,
        {{"a.取引ＩＤ", "伝票番号"}, {"レシート番号", "受注番号"}}
    ),
    列順を並替 = Table.ReorderColumns(
        列名を変更,
        {"伝票番号", "セクション", "支払方法", "支払番号", "支払額", "受注番号"}
    )
in
    列順を並替
