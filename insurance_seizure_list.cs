// ==============================================================================
// insurance_seizure_list.cs
// 生命保険差押予定一覧 作成ツール（WPF GUI版）
//
// 【使用方法】
// 1. build.bat を実行して insurance_seizure_list.exe を生成
// 2a. exe をダブルクリック → 初期画面で D&D またはファイル選択
// 2b. exe にファイルを D&D → 直接処理開始
// 2c. file_search.exe から呼び出し → 引数のファイルを処理後にクローズ
//
// 【処理概要】
// pipitLINQ 経由で取得した生命保険照会結果 .xlsm から必要情報を抽出し、
// 担当者の判断（執行日・シート選択）を加えて
// 生命保険差押予定一覧.csv に追記する
//
// 【ビルド方法】
// build.bat を実行（.NET Framework 4.0 の csc.exe を使用）
//
// 【必要ファイル（同じフォルダに配置）】
// ＜必須＞
// - insurance_seizure_list.cs  （ソースコード）
// - insurance_seizure_list_config.json （設定ファイル）
// - insurance_document_number_counter.json （文書番号カウンター）
// - era_mapping.json （元号マッピング）
// - insurance_seizure_list.ico （アプリケーションアイコン）
// - build.bat （ビルドスクリプト）
// ＜任意＞
// - institution_name_mapping.json（保険会社名補正）
// ==============================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Xml;

// ==============================================================
// データモデル
// ==============================================================

// プロファイル設定（config JSON の profiles 配列の1要素）
public class ProfileConfig
{
    public string Name { get; set; }                          // プロファイル名

    // シートフィルタ
    public string FilterCell { get; set; }                    // シートフィルタセル
    public string[] FilterValues { get; set; }                // シートフィルタ値（部分一致、OR条件）

    // 基本情報セル
    public string AddressNumberCell { get; set; }             // 宛名番号セル
    public string StaffCell { get; set; }                     // 職員名セル
    public string NameCell { get; set; }                      // 漢字氏名（照会側）セル
    public string KanaNameCell { get; set; }                  // カナ氏名（照会側）セル
    public string AddressCell { get; set; }                   // 漢字住所（照会側）セル
    public string BirthdayCell { get; set; }                  // 生年月日（照会側）セル
    public string InstitutionCodeCell { get; set; }           // 金融機関コードセル
    public string InstitutionNameCell { get; set; }           // 金融機関名称セル
    public string ContractExistsCell { get; set; }            // 保険契約有無セル

    // 金融機関側情報セル
    public string RespKanaNameCell { get; set; }              // カナ氏名（金融機関側）セル
    public string RespNameCell { get; set; }                  // 漢字氏名（金融機関側）セル
    public string RespBirthdayCell { get; set; }              // 生年月日（金融機関側）セル
    public string RespAddressCell { get; set; }               // 漢字住所（金融機関側）セル

    // 契約情報セル
    public string SeizureRightsCell { get; set; }             // 差押権利者の有無セル
    public string PolicyNumberCell { get; set; }              // 証券番号セル
    public string ContractStatusCell { get; set; }            // 契約の状態セル
    public string PremiumCell { get; set; }                   // 保険料セル
    public string PaymentFrequencyCell { get; set; }          // 払込区分セル
    public string ContractTypeCell { get; set; }              // 契約種類セル
    public string ContractDateCell { get; set; }              // 契約年月日セル
    public string MaturityDateCell { get; set; }              // 満期年月日セル
    public string InsuredNameCell { get; set; }               // 被保険者氏名セル

    // 金額情報セル
    public string SurrenderValueCell { get; set; }            // 解約返戻金セル
    public string DividendExistsCell { get; set; }            // 配当の有無セル
    public string DividendAmountCell { get; set; }            // 配当金額セル
    public string LoanExistsCell { get; set; }                // 貸付金の有無セル
    public string LoanAmountCell { get; set; }                // 貸付金の金額セル
    public string UnpaidPremiumExistsCell { get; set; }       // 未払い保険料の有無セル
    public string UnpaidPremiumAmountCell { get; set; }       // 未払い保険料の金額セル
    public string UnpaidInterestExistsCell { get; set; }      // 未払い利息の有無セル
    public string UnpaidInterestAmountCell { get; set; }      // 未払い利息の金額セル
    public string PrepaidPremiumExistsCell { get; set; }      // 前払い保険料の有無セル
    public string PrepaidPremiumAmountCell { get; set; }      // 前払い保険料の金額セル

    // 給付金額（内部保持用、UI非表示）
    public string BenefitLabelCol { get; set; }               // 給付金額ラベル列
    public string BenefitAmountCol { get; set; }              // 給付金額列
    public int BenefitStartRow { get; set; }                  // 給付金額開始行
    public int BenefitEndRow { get; set; }                    // 給付金額終了行

    // 特約（内部保持用、UI非表示）
    public string RiderLabelCol { get; set; }                 // 特約名列
    public string RiderAmountCol { get; set; }                // 特約金額列
    public int RiderStartRow { get; set; }                    // 特約開始行
    public int RiderEndRow { get; set; }                      // 特約終了行

    // 出力先
    public string OutputFolder { get; set; }                  // CSV出力先フォルダ
    public string PrintFolder { get; set; }                   // 印刷用ファイル保存先フォルダ
    public string DefaultFolder { get; set; }                 // ファイル選択ダイアログ初期フォルダ
}

// アプリ全体の設定
public class AppConfig
{
    public List<ProfileConfig> Profiles { get; set; }

    public AppConfig()
    {
        Profiles = new List<ProfileConfig>();
    }
}

// 元号マッピング1件
public class EraEntry
{
    public string Name { get; set; }                          // 元号名（例: "令和"）
    public int StartYear { get; set; }                        // 元号元年の西暦（例: 2019）
}

// 1シート分の照会結果データ（固定セルから読み取った全フィールド）
public class InsuranceData
{
    // 基本情報
    public string AddressNum { get; set; }                    // 宛名番号
    public string Name { get; set; }                          // 漢字氏名（照会側）
    public string KanaName { get; set; }                      // カナ氏名（照会側）
    public string Staff { get; set; }                         // 職員名
    public string Address { get; set; }                       // 漢字住所（照会側）
    public string Birthday { get; set; }                      // 生年月日（照会側）
    public string InstitutionCode { get; set; }               // 金融機関コード
    public string InstitutionName { get; set; }               // 金融機関名称（マッピング補正後）
    public string ContractExists { get; set; }                // 保険契約有無

    // 金融機関側情報
    public string RespKanaName { get; set; }                  // カナ氏名（金融機関側）
    public string RespName { get; set; }                      // 漢字氏名（金融機関側）
    public string RespBirthday { get; set; }                  // 生年月日（金融機関側）
    public string RespAddress { get; set; }                   // 漢字住所（金融機関側）

    // 契約情報
    public string SeizureRights { get; set; }                 // 差押権利者の有無
    public string PolicyNumber { get; set; }                  // 証券番号
    public string ContractStatus { get; set; }                // 契約の状態
    public double PremiumValue { get; set; }                  // 保険料（数値）
    public string PaymentFrequency { get; set; }              // 払込区分（月払/年払/一時払）
    public string PremiumDisplay { get; set; }                // 保険料（表示用: "11,690円/月"）
    public string ContractType { get; set; }                  // 契約種類
    public string ContractDate { get; set; }                  // 契約年月日（表示用文字列）
    public DateTime? ContractDateParsed { get; set; }         // 契約年月日（DateTime、差押文言用）
    public string MaturityDate { get; set; }                  // 満期年月日（表示用文字列）
    public string InsuredName { get; set; }                   // 被保険者氏名

    // 金額情報（数値: 差引見込額算出用、表示: UI表示用）
    public double SurrenderValue { get; set; }                // 解約返戻金
    public string DividendExists { get; set; }                // 配当の有無
    public double DividendAmount { get; set; }                // 配当金額
    public string LoanExists { get; set; }                    // 貸付金の有無
    public double LoanAmount { get; set; }                    // 貸付金の金額
    public string UnpaidPremiumExists { get; set; }           // 未払い保険料の有無
    public double UnpaidPremiumAmount { get; set; }           // 未払い保険料の金額
    public string UnpaidInterestExists { get; set; }          // 未払い利息の有無
    public double UnpaidInterestAmount { get; set; }          // 未払い利息の金額
    public string PrepaidPremiumExists { get; set; }          // 前払い保険料の有無
    public double PrepaidPremiumAmount { get; set; }          // 前払い保険料の金額

    // 差引見込額（解約返戻金 + 配当金 - 貸付金 - 未払い保険料 - 未払い利息 + 前払い保険料）
    public double NetValue { get; set; }

    // 給付金額（内部保持、UI非表示）
    public List<string[]> Benefits { get; set; }              // {ラベル, 金額} の配列

    // 特約（内部保持、UI非表示）
    public List<string[]> Riders { get; set; }                // {ラベル, 金額} の配列

    public InsuranceData()
    {
        Benefits = new List<string[]>();
        Riders = new List<string[]>();
    }
}

// ファイルの処理状態
public enum FileProcessState
{
    Pending,       // 未処理
    Added,         // 一覧に追加済み
    Skipped,       // スキップ
    Error          // エラー
}

// 処理対象ファイル1件の情報
public class FileEntry
{
    public string FilePath { get; set; }
    public FileProcessState State { get; set; }
}

// ==============================================================
// JSON パーサー（手書き・外部ライブラリ不要）
// LGWAN 環境では NuGet パッケージが使えないため手動パース
// ==============================================================

public static class JsonHelper
{
    // JSON文字列から指定キーの文字列値を取得
    public static string GetString(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return null;
        var colonIdx = json.IndexOf(':', keyIdx + key.Length + 2);
        if (colonIdx < 0) return null;

        var rest = json.Substring(colonIdx + 1).TrimStart();
        if (rest.Length == 0 || rest[0] != '"') return null;

        var sb = new StringBuilder();
        bool escaped = false;
        for (int i = 1; i < rest.Length; i++)
        {
            if (escaped) { sb.Append(rest[i]); escaped = false; continue; }
            if (rest[i] == '\\') { escaped = true; continue; }
            if (rest[i] == '"') break;
            sb.Append(rest[i]);
        }
        return sb.ToString();
    }

    // JSON文字列から指定キーの整数値を取得
    public static int GetInt(string json, string key, int defaultValue = 0)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return defaultValue;
        var colonIdx = json.IndexOf(':', keyIdx + key.Length + 2);
        if (colonIdx < 0) return defaultValue;

        var rest = json.Substring(colonIdx + 1).TrimStart();
        var numStr = new StringBuilder();
        foreach (var c in rest)
        {
            if (char.IsDigit(c) || c == '-') numStr.Append(c);
            else if (numStr.Length > 0) break;
        }
        int result;
        return int.TryParse(numStr.ToString(), out result) ? result : defaultValue;
    }

    // JSON文字列から指定キーの文字列配列を取得
    public static string[] GetStringArray(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return new string[0];
        var arrStart = json.IndexOf('[', keyIdx);
        if (arrStart < 0) return new string[0];
        var arrEnd = FindMatchingBracket(json, arrStart, '[', ']');
        if (arrEnd < 0) return new string[0];

        var inner = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
        return ExtractQuotedStrings(inner).ToArray();
    }

    // JSON オブジェクト配列を取得（各要素を文字列として返す）
    public static List<string> GetObjectArray(string json, string key)
    {
        var result = new List<string>();
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return result;
        var arrStart = json.IndexOf('[', keyIdx);
        if (arrStart < 0) return result;
        var arrEnd = FindMatchingBracket(json, arrStart, '[', ']');
        if (arrEnd < 0) return result;

        var inner = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
        int pos = 0;
        while (pos < inner.Length)
        {
            var objStart = inner.IndexOf('{', pos);
            if (objStart < 0) break;
            var objEnd = FindMatchingBracket(inner, objStart, '{', '}');
            if (objEnd < 0) break;
            result.Add(inner.Substring(objStart, objEnd - objStart + 1));
            pos = objEnd + 1;
        }
        return result;
    }

    // JSON オブジェクトをキーバリューの辞書として取得（値は文字列のみ対応）
    public static Dictionary<string, string> GetStringDictionary(string json)
    {
        var dict = new Dictionary<string, string>();
        int pos = 0;
        while (pos < json.Length)
        {
            var keyStart = json.IndexOf('"', pos);
            if (keyStart < 0) break;
            var keyEnd = json.IndexOf('"', keyStart + 1);
            if (keyEnd < 0) break;
            var key = json.Substring(keyStart + 1, keyEnd - keyStart - 1);

            if (key.StartsWith("_")) { pos = keyEnd + 1; continue; }

            var colonIdx = json.IndexOf(':', keyEnd + 1);
            if (colonIdx < 0) break;

            var valStart = json.IndexOf('"', colonIdx + 1);
            if (valStart < 0) { pos = colonIdx + 1; continue; }
            var valEnd = json.IndexOf('"', valStart + 1);
            if (valEnd < 0) break;
            var val = json.Substring(valStart + 1, valEnd - valStart - 1);

            dict[key] = val.Replace("\\\\", "\\");
            pos = valEnd + 1;
        }
        return dict;
    }

    // 対応する閉じ括弧の位置を返す
    public static int FindMatchingBracket(string json, int openIdx, char open, char close)
    {
        int depth = 0;
        bool inString = false;
        bool escaped = false;
        for (int i = openIdx; i < json.Length; i++)
        {
            if (escaped) { escaped = false; continue; }
            if (json[i] == '\\' && inString) { escaped = true; continue; }
            if (json[i] == '"') { inString = !inString; continue; }
            if (inString) continue;
            if (json[i] == open) depth++;
            else if (json[i] == close) { depth--; if (depth == 0) return i; }
        }
        return -1;
    }

    // JSON内の引用符で囲まれた文字列を全て抽出
    public static List<string> ExtractQuotedStrings(string content)
    {
        var result = new List<string>();
        int pos = 0;
        while (pos < content.Length)
        {
            var qStart = content.IndexOf('"', pos);
            if (qStart < 0) break;
            var qEnd = content.IndexOf('"', qStart + 1);
            if (qEnd < 0) break;
            result.Add(content.Substring(qStart + 1, qEnd - qStart - 1));
            pos = qEnd + 1;
        }
        return result;
    }

    // JSON出力用のエスケープ
    public static string Escape(string s)
    {
        if (s == null) return "";
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
    }
}

// ==============================================================
// 設定ファイル読込
// ==============================================================

public static class ConfigLoader
{
    // insurance_seizure_list_config.json を読み込む
    public static AppConfig LoadConfig(string configPath)
    {
        var config = new AppConfig();
        if (!File.Exists(configPath)) return config;

        var json = File.ReadAllText(configPath, Encoding.UTF8);

        foreach (var profileJson in JsonHelper.GetObjectArray(json, "profiles"))
        {
            var p = new ProfileConfig
            {
                Name                     = JsonHelper.GetString(profileJson, "name") ?? "",

                // シートフィルタ
                FilterCell               = JsonHelper.GetString(profileJson, "filterCell"),

                // 基本情報セル
                AddressNumberCell        = JsonHelper.GetString(profileJson, "addressNumberCell"),
                StaffCell                = JsonHelper.GetString(profileJson, "staffCell"),
                NameCell                 = JsonHelper.GetString(profileJson, "nameCell"),
                KanaNameCell             = JsonHelper.GetString(profileJson, "kanaNameCell"),
                AddressCell              = JsonHelper.GetString(profileJson, "addressCell"),
                BirthdayCell             = JsonHelper.GetString(profileJson, "birthdayCell"),
                InstitutionCodeCell      = JsonHelper.GetString(profileJson, "institutionCodeCell"),
                InstitutionNameCell      = JsonHelper.GetString(profileJson, "institutionNameCell"),
                ContractExistsCell       = JsonHelper.GetString(profileJson, "contractExistsCell"),

                // 金融機関側情報セル
                RespKanaNameCell         = JsonHelper.GetString(profileJson, "respKanaNameCell"),
                RespNameCell             = JsonHelper.GetString(profileJson, "respNameCell"),
                RespBirthdayCell         = JsonHelper.GetString(profileJson, "respBirthdayCell"),
                RespAddressCell          = JsonHelper.GetString(profileJson, "respAddressCell"),

                // 契約情報セル
                SeizureRightsCell        = JsonHelper.GetString(profileJson, "seizureRightsCell"),
                PolicyNumberCell         = JsonHelper.GetString(profileJson, "policyNumberCell"),
                ContractStatusCell       = JsonHelper.GetString(profileJson, "contractStatusCell"),
                PremiumCell              = JsonHelper.GetString(profileJson, "premiumCell"),
                PaymentFrequencyCell     = JsonHelper.GetString(profileJson, "paymentFrequencyCell"),
                ContractTypeCell         = JsonHelper.GetString(profileJson, "contractTypeCell"),
                ContractDateCell         = JsonHelper.GetString(profileJson, "contractDateCell"),
                MaturityDateCell         = JsonHelper.GetString(profileJson, "maturityDateCell"),
                InsuredNameCell          = JsonHelper.GetString(profileJson, "insuredNameCell"),

                // 金額情報セル
                SurrenderValueCell       = JsonHelper.GetString(profileJson, "surrenderValueCell"),
                DividendExistsCell       = JsonHelper.GetString(profileJson, "dividendExistsCell"),
                DividendAmountCell       = JsonHelper.GetString(profileJson, "dividendAmountCell"),
                LoanExistsCell           = JsonHelper.GetString(profileJson, "loanExistsCell"),
                LoanAmountCell           = JsonHelper.GetString(profileJson, "loanAmountCell"),
                UnpaidPremiumExistsCell   = JsonHelper.GetString(profileJson, "unpaidPremiumExistsCell"),
                UnpaidPremiumAmountCell   = JsonHelper.GetString(profileJson, "unpaidPremiumAmountCell"),
                UnpaidInterestExistsCell  = JsonHelper.GetString(profileJson, "unpaidInterestExistsCell"),
                UnpaidInterestAmountCell  = JsonHelper.GetString(profileJson, "unpaidInterestAmountCell"),
                PrepaidPremiumExistsCell  = JsonHelper.GetString(profileJson, "prepaidPremiumExistsCell"),
                PrepaidPremiumAmountCell  = JsonHelper.GetString(profileJson, "prepaidPremiumAmountCell"),

                // 給付金額範囲（内部保持用）
                BenefitLabelCol          = JsonHelper.GetString(profileJson, "benefitLabelCol"),
                BenefitAmountCol         = JsonHelper.GetString(profileJson, "benefitAmountCol"),
                BenefitStartRow          = JsonHelper.GetInt(profileJson, "benefitStartRow", 55),
                BenefitEndRow            = JsonHelper.GetInt(profileJson, "benefitEndRow", 58),

                // 特約範囲（内部保持用）
                RiderLabelCol            = JsonHelper.GetString(profileJson, "riderLabelCol"),
                RiderAmountCol           = JsonHelper.GetString(profileJson, "riderAmountCol"),
                RiderStartRow            = JsonHelper.GetInt(profileJson, "riderStartRow", 55),
                RiderEndRow              = JsonHelper.GetInt(profileJson, "riderEndRow", 64),

                // 出力先
                OutputFolder             = JsonHelper.GetString(profileJson, "outputFolder"),
                PrintFolder              = JsonHelper.GetString(profileJson, "printFolder"),
                DefaultFolder            = JsonHelper.GetString(profileJson, "defaultFolder")
            };
            if (p.OutputFolder != null) p.OutputFolder = p.OutputFolder.Replace("\\\\", "\\");
            if (p.PrintFolder != null)  p.PrintFolder  = p.PrintFolder.Replace("\\\\", "\\");
            if (p.DefaultFolder != null) p.DefaultFolder = p.DefaultFolder.Replace("\\\\", "\\");

            // filterValue: 文字列または配列をサポート（配列の場合はOR条件）
            string singleFilterValue = JsonHelper.GetString(profileJson, "filterValue");
            if (singleFilterValue != null)
                p.FilterValues = new[] { singleFilterValue };
            else
                p.FilterValues = JsonHelper.GetStringArray(profileJson, "filterValue");

            config.Profiles.Add(p);
        }

        return config;
    }

    // era_mapping.json を読み込む
    public static Dictionary<int, EraEntry> LoadEraMapping(string path)
    {
        var map = new Dictionary<int, EraEntry>();
        if (!File.Exists(path)) return map;

        var json = File.ReadAllText(path, Encoding.UTF8);
        for (int code = 1; code <= 9; code++)
        {
            var key = code.ToString();
            var keyIdx = json.IndexOf("\"" + key + "\"");
            if (keyIdx < 0) continue;
            var objStart = json.IndexOf('{', keyIdx);
            if (objStart < 0) continue;
            var objEnd = JsonHelper.FindMatchingBracket(json, objStart, '{', '}');
            if (objEnd < 0) continue;
            var objJson = json.Substring(objStart, objEnd - objStart + 1);

            var name = JsonHelper.GetString(objJson, "name");
            var startYear = JsonHelper.GetInt(objJson, "startYear");
            if (name != null && startYear > 0)
            {
                map[code] = new EraEntry { Name = name, StartYear = startYear };
            }
        }
        return map;
    }

    // 汎用マッピングファイルの読込（キー: 文字列, 値: 文字列）
    // institution_name_mapping.json に使用
    public static Dictionary<string, string> LoadSimpleMapping(string path)
    {
        if (!File.Exists(path)) return null;
        var json = File.ReadAllText(path, Encoding.UTF8);
        return JsonHelper.GetStringDictionary(json);
    }

    // document_number_counter.json から次の番号を読み取る
    public static int LoadNextDocNumber(string path)
    {
        if (!File.Exists(path)) return 1;
        var json = File.ReadAllText(path, Encoding.UTF8);
        return JsonHelper.GetInt(json, "nextNumber", 1);
    }

    // document_number_counter.json に次の番号を書き込む
    public static void SaveNextDocNumber(string path, int nextNumber)
    {
        var json = "{\n    \"nextNumber\":  " + nextNumber + "\n}";
        File.WriteAllText(path, json, Encoding.UTF8);
    }
}

// ==============================================================
// ヘルパー（ビジネスロジック）
// ==============================================================

public static class BusinessLogic
{
    // 列アルファベットを列インデックスに変換（例: "C" → 3, "AA" → 27）
    public static int ColToIndex(string col)
    {
        col = col.ToUpper().Trim();
        int result = 0;
        foreach (var ch in col.ToCharArray())
            result = result * 26 + ((int)ch - (int)'A' + 1);
        return result;
    }

    // 2D配列から安全に文字列を取得（Value2 一括取得用、1始まりインデックス）
    public static string BulkCell(object[,] data, int row, int col)
    {
        if (data == null) return "";
        try { return Convert.ToString(data[row, col] ?? ""); }
        catch { return ""; }
    }

    // 2D配列から生の値を取得（日付・数値のシリアル値判定用）
    public static object BulkCellRaw(object[,] data, int row, int col)
    {
        if (data == null) return null;
        try { return data[row, col]; }
        catch { return null; }
    }

    // セルアドレスから2D配列の値を文字列として取得（A1始まり前提）
    // 例: "L8" → data[8, 12]
    public static string BulkCellByAddress(object[,] data, string cellAddr)
    {
        if (data == null || string.IsNullOrEmpty(cellAddr)) return "";
        int i = 0;
        while (i < cellAddr.Length && char.IsLetter(cellAddr[i])) i++;
        if (i == 0 || i >= cellAddr.Length) return "";
        int col = ColToIndex(cellAddr.Substring(0, i));
        int row;
        if (!int.TryParse(cellAddr.Substring(i), out row)) return "";
        return BulkCell(data, row, col);
    }

    // セルアドレスから2D配列の生の値を取得（日付・数値のシリアル値判定用）
    public static object BulkCellRawByAddress(object[,] data, string cellAddr)
    {
        if (data == null || string.IsNullOrEmpty(cellAddr)) return null;
        int i = 0;
        while (i < cellAddr.Length && char.IsLetter(cellAddr[i])) i++;
        if (i == 0 || i >= cellAddr.Length) return null;
        int col = ColToIndex(cellAddr.Substring(0, i));
        int row;
        if (!int.TryParse(cellAddr.Substring(i), out row)) return null;
        return BulkCellRaw(data, row, col);
    }

    // 宛名番号を10桁0埋め
    public static string FormatAddressNumber(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        var clean = value.Trim().Replace(" ", "").Replace("\u3000", "");
        long num;
        if (long.TryParse(clean, out num) && clean.Length <= 10)
            return clean.PadLeft(10, '0');
        return clean;
    }

    // 住所の基本加工（郵便番号除去・空白除去）
    public static string FormatAddress(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        var addr = System.Text.RegularExpressions.Regex.Replace(raw, @"〒\d{3}-\d{4}[\s\u3000]*", "");
        addr = addr.Replace(" ", "").Replace("\u3000", "");
        return addr;
    }

    // 残高・金額を通貨形式に整形（例: 17241 → "17,241円"）
    public static string FormatBalance(object value)
    {
        if (value == null) return "-";
        var str = value.ToString().Trim();
        if (string.IsNullOrEmpty(str)) return "-";
        double num;
        if (double.TryParse(str, out num))
            return string.Format("{0:N0}円", num);
        return str;
    }

    // 保険料と払込区分から表示文字列を生成（例: "11,690円/月"）
    public static string FormatPremiumDisplay(double amount, string frequency)
    {
        string amountStr = string.Format("{0:N0}円", amount);
        if (string.IsNullOrEmpty(frequency)) return amountStr;
        string freq = frequency.Trim();
        if (freq.Contains("月")) return amountStr + "/月";
        if (freq.Contains("年")) return amountStr + "/年";
        if (freq.Contains("一時")) return amountStr + "（一時払）";
        return amountStr;
    }

    // CSVフィールドのエスケープ（RFC 4180準拠）
    public static string CsvEscape(string field)
    {
        if (field == null) return "";
        if (field.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        return field;
    }

    // 半角英数字・半角カタカナを全角に変換
    public static string ToFullWidth(string str)
    {
        if (string.IsNullOrEmpty(str)) return "";

        var kanaMap = new Dictionary<string, string>
        {
            {"ｶﾞ","ガ"},{"ｷﾞ","ギ"},{"ｸﾞ","グ"},{"ｹﾞ","ゲ"},{"ｺﾞ","ゴ"},
            {"ｻﾞ","ザ"},{"ｼﾞ","ジ"},{"ｽﾞ","ズ"},{"ｾﾞ","ゼ"},{"ｿﾞ","ゾ"},
            {"ﾀﾞ","ダ"},{"ﾁﾞ","ヂ"},{"ﾂﾞ","ヅ"},{"ﾃﾞ","デ"},{"ﾄﾞ","ド"},
            {"ﾊﾞ","バ"},{"ﾋﾞ","ビ"},{"ﾌﾞ","ブ"},{"ﾍﾞ","ベ"},{"ﾎﾞ","ボ"},
            {"ｳﾞ","ヴ"},{"ﾊﾟ","パ"},{"ﾋﾟ","ピ"},{"ﾌﾟ","プ"},{"ﾍﾟ","ペ"},{"ﾎﾟ","ポ"},
            {"ｦ","ヲ"},{"ｧ","ァ"},{"ｨ","ィ"},{"ｩ","ゥ"},{"ｪ","ェ"},{"ｫ","ォ"},
            {"ｬ","ャ"},{"ｭ","ュ"},{"ｮ","ョ"},{"ｯ","ッ"},{"ｰ","ー"},
            {"ｱ","ア"},{"ｲ","イ"},{"ｳ","ウ"},{"ｴ","エ"},{"ｵ","オ"},
            {"ｶ","カ"},{"ｷ","キ"},{"ｸ","ク"},{"ｹ","ケ"},{"ｺ","コ"},
            {"ｻ","サ"},{"ｼ","シ"},{"ｽ","ス"},{"ｾ","セ"},{"ｿ","ソ"},
            {"ﾀ","タ"},{"ﾁ","チ"},{"ﾂ","ツ"},{"ﾃ","テ"},{"ﾄ","ト"},
            {"ﾅ","ナ"},{"ﾆ","ニ"},{"ﾇ","ヌ"},{"ﾈ","ネ"},{"ﾉ","ノ"},
            {"ﾊ","ハ"},{"ﾋ","ヒ"},{"ﾌ","フ"},{"ﾍ","ヘ"},{"ﾎ","ホ"},
            {"ﾏ","マ"},{"ﾐ","ミ"},{"ﾑ","ム"},{"ﾒ","メ"},{"ﾓ","モ"},
            {"ﾔ","ヤ"},{"ﾕ","ユ"},{"ﾖ","ヨ"},
            {"ﾗ","ラ"},{"ﾘ","リ"},{"ﾙ","ル"},{"ﾚ","レ"},{"ﾛ","ロ"},
            {"ﾜ","ワ"},{"ﾝ","ン"},{"ﾞ","゛"},{"ﾟ","゜"}
        };

        // まず濁点・半濁点の結合（2文字→1文字）を先に処理
        foreach (var pair in kanaMap.Where(p => p.Key.Length == 2))
            str = str.Replace(pair.Key, pair.Value);

        var sb = new StringBuilder();
        foreach (var c in str.ToCharArray())
        {
            int code = (int)c;
            if (code >= 0x41 && code <= 0x5A)       // 半角英大文字 → 全角
                sb.Append((char)(code + 0xFEE0));
            else if (code >= 0x61 && code <= 0x7A)   // 半角英小文字 → 全角
                sb.Append((char)(code + 0xFEE0));
            else if (code >= 0x30 && code <= 0x39)   // 半角数字 → 全角
                sb.Append((char)(code + 0xFEE0));
            else if (kanaMap.ContainsKey(c.ToString()))
                sb.Append(kanaMap[c.ToString()]);
            else
                sb.Append(c);
        }
        return sb.ToString();
    }

    // Excel セルの値を DateTime に変換（シリアル値・文字列の両方に対応）
    public static DateTime? ParseExcelDate(object rawValue)
    {
        if (rawValue == null) return null;
        if (rawValue is double)
            return DateTime.FromOADate((double)rawValue);
        string str = rawValue.ToString().Trim();
        if (string.IsNullOrEmpty(str)) return null;
        DateTime dt;
        if (DateTime.TryParse(str, out dt)) return dt;
        // "2013年02月11日" 形式
        var match = System.Text.RegularExpressions.Regex.Match(str, @"(\d{4})年(\d{1,2})月(\d{1,2})日");
        if (match.Success)
        {
            try
            {
                return new DateTime(
                    int.Parse(match.Groups[1].Value),
                    int.Parse(match.Groups[2].Value),
                    int.Parse(match.Groups[3].Value));
            }
            catch { }
        }
        return null;
    }

    // DateTime → 和暦全角表示（ゼロ埋めなし）
    // 差押文言用: 例 DateTime(2013,2,11) → "平成２５年２月１１日"
    public static string FormatWarekiDisplay(DateTime date, Dictionary<int, EraEntry> eraMap)
    {
        foreach (var pair in eraMap.OrderByDescending(p => p.Value.StartYear))
        {
            if (date.Year >= pair.Value.StartYear)
            {
                int warekiYear = date.Year - pair.Value.StartYear + 1;
                return pair.Value.Name
                    + ToFullWidth(warekiYear.ToString()) + "年"
                    + ToFullWidth(date.Month.ToString()) + "月"
                    + ToFullWidth(date.Day.ToString()) + "日";
            }
        }
        return "";
    }

    // 7桁和暦 → DateTime 変換（era_mapping.json 使用）
    public static DateTime? WarekiToDate(string wareki, Dictionary<int, EraEntry> eraMap)
    {
        if (string.IsNullOrWhiteSpace(wareki) || wareki.Length != 7) return null;
        int eraCode, year, month, day;
        if (!int.TryParse(wareki.Substring(0, 1), out eraCode)) return null;
        if (!int.TryParse(wareki.Substring(1, 2), out year)) return null;
        if (!int.TryParse(wareki.Substring(3, 2), out month)) return null;
        if (!int.TryParse(wareki.Substring(5, 2), out day)) return null;

        EraEntry era;
        if (!eraMap.TryGetValue(eraCode, out era)) return null;
        int adYear = era.StartYear + year - 1;

        try { return new DateTime(adYear, month, day); }
        catch { return null; }
    }

    // DateTime → 7桁和暦変換（CSV出力用）
    public static string DateToWareki(DateTime date, Dictionary<int, EraEntry> eraMap)
    {
        foreach (var pair in eraMap.OrderByDescending(p => p.Value.StartYear))
        {
            if (date.Year >= pair.Value.StartYear)
            {
                int warekiYear = date.Year - pair.Value.StartYear + 1;
                return string.Format("{0}{1:D2}{2:D2}{3:D2}",
                    pair.Key, warekiYear, date.Month, date.Day);
            }
        }
        return "";
    }

    // 入力文字列を DateTime に変換（複数形式対応）
    // 7桁和暦 / 8桁西暦 / yyyy/MM/dd / yyyy/M/d
    public static DateTime? ParseFlexibleDate(string input, Dictionary<int, EraEntry> eraMap)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        input = input.Trim();

        if (input.Contains("/"))
        {
            DateTime dt;
            if (DateTime.TryParse(input, out dt)) return dt;
            return null;
        }

        if (input.Length == 7 && input.All(char.IsDigit))
            return WarekiToDate(input, eraMap);

        if (input.Length == 8 && input.All(char.IsDigit))
        {
            DateTime dt;
            if (DateTime.TryParseExact(input, "yyyyMMdd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out dt))
                return dt;
        }

        return null;
    }

    // Excel セルの値を double に変換（数値・文字列・null に対応）
    public static double ParseAmount(object rawValue)
    {
        if (rawValue == null) return 0;
        if (rawValue is double) return (double)rawValue;
        var str = rawValue.ToString().Trim().Replace(",", "").Replace("円", "");
        double result;
        return double.TryParse(str, out result) ? result : 0;
    }

    // 差引見込額を算出
    public static double CalcNetValue(InsuranceData d)
    {
        return d.SurrenderValue
            + d.DividendAmount
            - d.LoanAmount
            - d.UnpaidPremiumAmount
            - d.UnpaidInterestAmount
            + d.PrepaidPremiumAmount;
    }
}

// ==============================================================
// メインアプリケーション
// ==============================================================

public class InsuranceSeizureApp : Application
{
    // --- 設定 ---
    private AppConfig config;
    private ProfileConfig activeProfile;
    private string exeDir;
    private Dictionary<int, EraEntry> eraMapping;
    private Dictionary<string, string> institutionNameMapping;

    // --- 状態 ---
    private List<FileEntry> fileEntries = new List<FileEntry>();
    private int currentFileIndex = -1;
    private DateTime? processingDate = null;
    private bool isFromFileSearch = false;
    private bool suppressSheetChange = false;
    private DateTime? currentContractDate = null;        // 差押文言用の契約年月日

    // --- キャッシュ済みブラシ（アンバー #E69500 ベース） ---
    private static readonly SolidColorBrush BrushBorderNormal    = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D0D0D0")));
    private static readonly SolidColorBrush BrushValidationError = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D32F2F")));
    private static readonly SolidColorBrush BrushSuccessIcon     = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#107C41")));
    private static readonly SolidColorBrush BrushAccent          = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E69500")));
    private static readonly SolidColorBrush BrushIconBgSuccess   = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F5ED")));
    private static readonly SolidColorBrush BrushIconBgError     = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCEBEB")));
    private static readonly SolidColorBrush BrushIconBgSkip      = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF3E0")));
    private static readonly SolidColorBrush BrushDetailMuted     = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#999")));
    private static SolidColorBrush Frozen(SolidColorBrush b) { b.Freeze(); return b; }

    // --- 定数 ---
    private const int CSV_WRITE_MAX_RETRY = 5;
    private const int CSV_WRITE_RETRY_INTERVAL_MS = 500;
    private const int MAX_PATH = 260;
    private const string CSV_FILENAME = "生命保険差押予定一覧.csv";
    private const string CSV_HEADER = "登録日時,宛名番号,氏名,職員名,執行日,住民票住所,届出住所,"
        + "保険会社名,証券番号,契約種類,"
        + "差押文言1,差押文言2,差押文言3,差押文言4,差押文言5,差押文言6,"
        + "文書番号,照会結果ファイル名,処理済フラグ1,処理済フラグ2,処理済フラグ3";
    private const string CONFIG_FILE = "insurance_seizure_list_config.json";
    private const string DOC_NUMBER_FILE = "insurance_document_number_counter.json";
    private const string BULK_READ_END = "AN75";

    // --- UI要素 ---
    private Window window;
    private Grid initialPanel, mainPanel, overlayPanel, loadingOverlay, resultOverlay;
    private ComboBox sheetCombo;
    private TextBlock fileLink, statusLeft, statusRight, guideText, deliveryError;
    private TextBlock resultIcon, resultTitle, resultDetail, resultSub;
    private Border resultIconBg;
    private TextBox txtAddressNum, txtName, txtInstitution, txtStaff;
    private TextBox txtResidenceAddr, txtDeliveryAddr, txtExecDate;
    private CheckBox chkDeliveryOutput;
    private Button btnAdd, btnSkip, btnLoadFile, resultButton;
    private Button btnCalendar;
    private Popup calendarPopup;
    private System.Windows.Controls.Calendar dateCalendar;
    private RotateTransform spinnerRotation;
    // 契約情報テキスト
    private TextBlock txtContractExists, txtContractStatus, txtPolicyNumber;
    private TextBlock txtContractType, txtContractDate, txtMaturityDate;
    private TextBlock txtPremium, txtInsuredName, txtSeizureRights;
    private FrameworkElement warningBanner;
    private TextBlock warningIcon, warningText;
    // 金額情報テキスト
    private TextBlock txtSurrenderValue, txtNetValue;
    private TextBlock lblDividend, txtDividend, lblLoan, txtLoan;
    private TextBlock lblUnpaidPremium, txtUnpaidPremium;
    private TextBlock lblUnpaidInterest, txtUnpaidInterest;
    private TextBlock lblPrepaidPremium, txtPrepaidPremium;

    // --- Excel COM ---
    private dynamic excel;
    private string currentFilePath;
    private string selectedSheetName;
    private string lastDocNumber;

    // ==============================================================
    // エントリポイント
    // ==============================================================

    [STAThread]
    public static void Main(string[] args)
    {
        var app = new InsuranceSeizureApp();
        app.StartupArgs = args;
        app.Run();
    }

    private string[] StartupArgs { get; set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        exeDir = AppDomain.CurrentDomain.BaseDirectory;

        // --- 設定ファイル読込 ---
        var configPath = System.IO.Path.Combine(exeDir, CONFIG_FILE);
        if (!File.Exists(configPath))
        { MessageBox.Show("設定ファイルが見つかりません。\n\n" + configPath, "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }
        config = ConfigLoader.LoadConfig(configPath);
        if (config.Profiles.Count == 0)
        { MessageBox.Show("プロファイルが設定されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 元号マッピング読込 ---
        eraMapping = ConfigLoader.LoadEraMapping(System.IO.Path.Combine(exeDir, "era_mapping.json"));
        if (eraMapping.Count == 0)
        { MessageBox.Show("era_mapping.json が見つからないか空です。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 文書番号カウンター確認 ---
        if (!File.Exists(System.IO.Path.Combine(exeDir, DOC_NUMBER_FILE)))
        { MessageBox.Show(DOC_NUMBER_FILE + " が見つかりません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 任意マッピングファイル読込 ---
        institutionNameMapping = ConfigLoader.LoadSimpleMapping(System.IO.Path.Combine(exeDir, "institution_name_mapping.json"));

        // --- プロファイル選択 ---
        activeProfile = config.Profiles[0];

        // --- プロファイルバリデーション ---
        var ve = new List<string>();
        if (string.IsNullOrWhiteSpace(activeProfile.AddressNumberCell)) ve.Add("addressNumberCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.NameCell)) ve.Add("nameCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.StaffCell)) ve.Add("staffCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.AddressCell)) ve.Add("addressCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.InstitutionNameCell)) ve.Add("institutionNameCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.PolicyNumberCell)) ve.Add("policyNumberCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.SurrenderValueCell)) ve.Add("surrenderValueCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.ContractExistsCell)) ve.Add("contractExistsCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.OutputFolder)) ve.Add("outputFolder が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.PrintFolder)) ve.Add("printFolder が未設定です");
        if (ve.Count > 0)
        { MessageBox.Show("プロファイル設定エラー:\n\n" + string.Join("\n", ve), "設定エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 出力先フォルダの自動作成 ---
        try
        {
            if (!Directory.Exists(activeProfile.OutputFolder)) Directory.CreateDirectory(activeProfile.OutputFolder);
            if (!Directory.Exists(activeProfile.PrintFolder))  Directory.CreateDirectory(activeProfile.PrintFolder);
        }
        catch (Exception dirEx)
        { MessageBox.Show("出力先フォルダの作成に失敗しました。\n\n" + dirEx.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); return; }

        // --- 起動モード判定 ---
        if (StartupArgs != null && StartupArgs.Length > 0)
        {
            isFromFileSearch = true;
            foreach (var arg in StartupArgs)
                if (File.Exists(arg)) fileEntries.Add(new FileEntry { FilePath = arg, State = FileProcessState.Pending });
        }

        // --- ウィンドウ構築 ---
        window = BuildWindow();
        FindControls();
        SetupEvents();
        InitializeUI();
        if (isFromFileSearch && fileEntries.Count > 0)
            window.ContentRendered += delegate { LoadFileAtIndex(0); };
        window.Closed += delegate { CleanupExcel(); };
        window.Show();
    }

    private void FindControls()
    {
        initialPanel = (Grid)window.FindName("InitialPanel");
        mainPanel = (Grid)window.FindName("MainPanel");
        overlayPanel = (Grid)window.FindName("OverlayPanel");
        loadingOverlay = (Grid)window.FindName("LoadingOverlay");
        resultOverlay = (Grid)window.FindName("ResultOverlay");
        sheetCombo = (ComboBox)window.FindName("SheetCombo");
        fileLink = (TextBlock)window.FindName("FileLink");
        txtAddressNum = (TextBox)window.FindName("TxtAddressNum");
        txtName = (TextBox)window.FindName("TxtName");
        txtInstitution = (TextBox)window.FindName("TxtInstitution");
        txtStaff = (TextBox)window.FindName("TxtStaff");
        txtResidenceAddr = (TextBox)window.FindName("TxtResidenceAddr");
        txtDeliveryAddr = (TextBox)window.FindName("TxtDeliveryAddr");
        txtExecDate = (TextBox)window.FindName("TxtExecDate");
        chkDeliveryOutput = (CheckBox)window.FindName("ChkDeliveryOutput");
        btnAdd = (Button)window.FindName("BtnAdd");
        btnSkip = (Button)window.FindName("BtnSkip");
        btnLoadFile = (Button)window.FindName("BtnLoadFile");
        statusLeft = (TextBlock)window.FindName("StatusLeft");
        statusRight = (TextBlock)window.FindName("StatusRight");
        guideText = (TextBlock)window.FindName("GuideText");
        deliveryError = (TextBlock)window.FindName("DeliveryError");
        resultIcon = (TextBlock)window.FindName("ResultIcon");
        resultIconBg = (Border)window.FindName("ResultIconBg");
        resultTitle = (TextBlock)window.FindName("ResultTitle");
        resultDetail = (TextBlock)window.FindName("ResultDetail");
        resultButton = (Button)window.FindName("ResultButton");
        resultSub = (TextBlock)window.FindName("ResultSub");
        btnCalendar = (Button)window.FindName("BtnCalendar");
        calendarPopup = (Popup)window.FindName("CalendarPopup");
        dateCalendar = (System.Windows.Controls.Calendar)window.FindName("DateCalendar");
        // 契約情報
        txtContractExists = (TextBlock)window.FindName("TxtContractExists");
        txtContractStatus = (TextBlock)window.FindName("TxtContractStatus");
        txtPolicyNumber = (TextBlock)window.FindName("TxtPolicyNumber");
        txtContractType = (TextBlock)window.FindName("TxtContractType");
        txtContractDate = (TextBlock)window.FindName("TxtContractDate");
        txtMaturityDate = (TextBlock)window.FindName("TxtMaturityDate");
        txtPremium = (TextBlock)window.FindName("TxtPremium");
        txtInsuredName = (TextBlock)window.FindName("TxtInsuredName");
        txtSeizureRights = (TextBlock)window.FindName("TxtSeizureRights");
        warningBanner = (FrameworkElement)window.FindName("WarningBanner");
        warningIcon = (TextBlock)window.FindName("WarningIcon");
        warningText = (TextBlock)window.FindName("WarningText");
        // 金額情報
        txtSurrenderValue = (TextBlock)window.FindName("TxtSurrenderValue");
        txtNetValue = (TextBlock)window.FindName("TxtNetValue");
        lblDividend = (TextBlock)window.FindName("LblDividend");
        txtDividend = (TextBlock)window.FindName("TxtDividend");
        lblLoan = (TextBlock)window.FindName("LblLoan");
        txtLoan = (TextBlock)window.FindName("TxtLoan");
        lblUnpaidPremium = (TextBlock)window.FindName("LblUnpaidPremium");
        txtUnpaidPremium = (TextBlock)window.FindName("TxtUnpaidPremium");
        lblUnpaidInterest = (TextBlock)window.FindName("LblUnpaidInterest");
        txtUnpaidInterest = (TextBlock)window.FindName("TxtUnpaidInterest");
        lblPrepaidPremium = (TextBlock)window.FindName("LblPrepaidPremium");
        txtPrepaidPremium = (TextBlock)window.FindName("TxtPrepaidPremium");
        // スピナー
        var spinnerElement = (FrameworkElement)window.FindName("SpinnerPath");
        if (spinnerElement != null)
            spinnerRotation = spinnerElement.RenderTransform as RotateTransform;
    }

    private void InitializeUI()
    {
        if (activeProfile.OutputFolder != null) statusRight.Text = activeProfile.OutputFolder;
        ShowState("initial");
    }

    private void ShowState(string state)
    {
        initialPanel.Visibility = state == "initial" ? Visibility.Visible : Visibility.Collapsed;
        mainPanel.Visibility = state == "main" ? Visibility.Visible : Visibility.Collapsed;
        overlayPanel.Visibility = (state == "loading" || state == "result") ? Visibility.Visible : Visibility.Collapsed;
        loadingOverlay.Visibility = state == "loading" ? Visibility.Visible : Visibility.Collapsed;
        resultOverlay.Visibility = state == "result" ? Visibility.Visible : Visibility.Collapsed;
        if (state == "initial") statusLeft.Text = "";
    }

    // ==============================================================
    // イベントハンドラ
    // ==============================================================

    private void SetupEvents()
    {
        // D&D
        window.AllowDrop = true;
        window.Drop += delegate(object s, DragEventArgs de)
        {
            if (!de.Data.GetDataPresent(DataFormats.FileDrop)) return;
            var files = ((string[])de.Data.GetData(DataFormats.FileDrop))
                .Where(f => f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)).ToArray();
            if (files.Length > 0) StartFileProcessing(files);
        };
        window.DragOver += delegate(object s, DragEventArgs de)
        { de.Effects = de.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None; de.Handled = true; };

        // 初期画面ファイル選択
        var btnSelect = (Button)window.FindName("BtnSelectFile");
        if (btnSelect != null) btnSelect.Click += delegate { DoOpenFileDialog(true); };

        // メインフォームボタン
        btnLoadFile.Click += delegate { DoOpenFileDialog(false); };
        fileLink.MouseDown += delegate
        { if (!string.IsNullOrEmpty(currentFilePath) && File.Exists(currentFilePath)) try { Process.Start(new ProcessStartInfo(currentFilePath) { UseShellExecute = true }); } catch {} };
        var btnReload = (Button)window.FindName("BtnReload");
        if (btnReload != null) btnReload.Click += delegate { if (!string.IsNullOrEmpty(currentFilePath)) LoadSingleFile(currentFilePath); };

        // シート切替
        sheetCombo.SelectionChanged += delegate { if (!suppressSheetChange && sheetCombo.SelectedItem != null) { selectedSheetName = sheetCombo.SelectedItem.ToString(); ReloadSheetData(); } };

        // 必須フィールドでボタン制御
        txtName.TextChanged += delegate { UpdateAddButton(); };
        txtStaff.TextChanged += delegate { UpdateAddButton(); };
        txtResidenceAddr.TextChanged += delegate { UpdateAddButton(); };
        txtExecDate.LostFocus += delegate
        {
            var input = txtExecDate.Text.Trim();
            if (string.IsNullOrEmpty(input)) { processingDate = null; UpdateAddButton(); return; }
            var dt = BusinessLogic.ParseFlexibleDate(input, eraMapping);
            if (dt.HasValue)
            {
                processingDate = dt.Value;
                txtExecDate.Text = dt.Value.ToString("yyyy/MM/dd");
                txtExecDate.BorderBrush = BrushBorderNormal;
            }
            else
            {
                processingDate = null;
                txtExecDate.BorderBrush = BrushValidationError;
            }
            UpdateAddButton();
        };

        // 届出住所バリデーション
        txtDeliveryAddr.TextChanged += delegate { ValidateDelivery(); };
        chkDeliveryOutput.Checked += delegate { ValidateDelivery(); };
        chkDeliveryOutput.Unchecked += delegate { ValidateDelivery(); };

        // カレンダーPopup
        btnCalendar.Click += delegate
        {
            calendarPopup.PlacementTarget = btnCalendar;
            if (!calendarPopup.IsOpen)
            {
                dateCalendar.DisplayMode = CalendarMode.Month;
                dateCalendar.DisplayDate = processingDate ?? DateTime.Today;
                dateCalendar.SelectedDates.Clear();
            }
            calendarPopup.IsOpen = !calendarPopup.IsOpen;
        };
        dateCalendar.SelectedDatesChanged += delegate
        {
            if (dateCalendar.SelectedDate.HasValue)
            {
                processingDate = dateCalendar.SelectedDate.Value;
                txtExecDate.Text = processingDate.Value.ToString("yyyy/MM/dd");
                txtExecDate.BorderBrush = BrushBorderNormal;
                calendarPopup.IsOpen = false;
                UpdateAddButton();
            }
        };

        // アクション
        btnAdd.Click += delegate { ExecuteAdd(); };
        btnSkip.Click += delegate { ExecuteSkip(); };
        resultButton.Click += delegate { ProceedToNext(); };

        // ショートカット
        window.InputBindings.Add(new KeyBinding(new RelayCommand(p => { if (overlayPanel.Visibility == Visibility.Visible) ProceedToNext(); }), new KeyGesture(Key.Escape)));
        window.InputBindings.Add(new KeyBinding(new RelayCommand(p => DoOpenFileDialog(initialPanel.Visibility == Visibility.Visible)), new KeyGesture(Key.O, ModifierKeys.Control)));

        // 空白クリックでフォーカスを外す
        window.MouseDown += delegate(object s, MouseButtonEventArgs me)
        {
            if (me.OriginalSource is System.Windows.Controls.Panel ||
                me.OriginalSource is Border || me.OriginalSource is Window)
            { FocusManager.SetFocusedElement(window, window); Keyboard.ClearFocus(); }
        };

        // スピナーアニメーション: LoadingOverlay の表示/非表示に連動
        loadingOverlay.IsVisibleChanged += delegate(object s, DependencyPropertyChangedEventArgs dpce)
        {
            if ((bool)dpce.NewValue) StartSpinner();
            else StopSpinner();
        };
    }

    private void UpdateAddButton()
    {
        string msg = "";
        if (string.IsNullOrWhiteSpace(txtName.Text)) msg = "氏名を入力してください";
        else if (string.IsNullOrWhiteSpace(txtStaff.Text)) msg = "処分担当を入力してください";
        else if (string.IsNullOrWhiteSpace(txtResidenceAddr.Text)) msg = "住民票住所を入力してください";
        else if (!processingDate.HasValue) msg = "執行日を入力してください";
        // 保険契約無の場合はボタン無効化
        if (string.IsNullOrEmpty(msg) && txtContractExists != null &&
            txtContractExists.Text.Trim() == "無")
            msg = "保険契約がありません";
        btnAdd.IsEnabled = string.IsNullOrEmpty(msg);
        guideText.Text = msg;
    }

    private void ValidateDelivery()
    {
        if (chkDeliveryOutput.IsChecked == true && !string.IsNullOrEmpty(txtDeliveryAddr.Text))
        {
            int len = ("（届出：" + txtDeliveryAddr.Text.Trim() + "）").Length;
            if (len > 50) { deliveryError.Text = "50文字を超えています（現在: " + len + "文字）"; deliveryError.Visibility = Visibility.Visible; return; }
        }
        deliveryError.Visibility = Visibility.Collapsed;
    }

    // ==============================================================
    // アニメーション
    // ==============================================================

    private void StartSpinner()
    {
        if (spinnerRotation == null) return;
        var anim = new DoubleAnimation { By = 360, Duration = new Duration(TimeSpan.FromSeconds(1)), RepeatBehavior = RepeatBehavior.Forever };
        spinnerRotation.BeginAnimation(RotateTransform.AngleProperty, anim);
    }

    private void StopSpinner()
    {
        if (spinnerRotation == null) return;
        spinnerRotation.BeginAnimation(RotateTransform.AngleProperty, null);
    }

    private void FadeInOverlay()
    {
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, null);
        overlayPanel.Opacity = 0;
        overlayPanel.Visibility = Visibility.Visible;
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
    }

    private void FadeOutOverlay(Action onComplete = null)
    {
        var anim = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(150)));
        anim.Completed += delegate
        {
            overlayPanel.Visibility = Visibility.Collapsed;
            overlayPanel.BeginAnimation(UIElement.OpacityProperty, null);
            overlayPanel.Opacity = 1;
            if (onComplete != null) onComplete();
        };
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, anim);
    }

    private void CrossFadeToResult(string type, string title, string detail)
    {
        var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(100)));
        fadeOut.Completed += delegate
        {
            ShowResult(type, title, detail);
            overlayPanel.BeginAnimation(UIElement.OpacityProperty,
                new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(100))));
        };
        overlayPanel.BeginAnimation(UIElement.OpacityProperty, fadeOut);
    }

    // ==============================================================
    // ファイル処理
    // ==============================================================

    private void DoOpenFileDialog(bool isInitial)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Excel (*.xlsm)|*.xlsm", Multiselect = true, Title = "照会結果ファイルを選択" };
        if (!isInitial && !string.IsNullOrEmpty(currentFilePath)) dlg.InitialDirectory = System.IO.Path.GetDirectoryName(currentFilePath);
        else if (activeProfile.DefaultFolder != null && Directory.Exists(activeProfile.DefaultFolder)) dlg.InitialDirectory = activeProfile.DefaultFolder;
        if (dlg.ShowDialog() == true) StartFileProcessing(dlg.FileNames);
    }

    private void StartFileProcessing(string[] paths)
    {
        fileEntries.Clear();
        foreach (var p in paths) fileEntries.Add(new FileEntry { FilePath = p, State = FileProcessState.Pending });
        LoadFileAtIndex(0);
    }

    private void LoadFileAtIndex(int index)
    {
        if (index >= fileEntries.Count)
        {
            if (isFromFileSearch) { Shutdown(); return; }
            var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(150)));
            fadeOut.Completed += delegate
            {
                mainPanel.BeginAnimation(UIElement.OpacityProperty, null);
                mainPanel.Opacity = 1;
                ShowState("initial");
                initialPanel.Opacity = 0;
                initialPanel.BeginAnimation(UIElement.OpacityProperty,
                    new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
            };
            mainPanel.BeginAnimation(UIElement.OpacityProperty, fadeOut);
            return;
        }
        currentFileIndex = index;
        currentFilePath = fileEntries[index].FilePath;
        statusLeft.Text = fileEntries.Count > 1 ? (index + 1) + " / " + fileEntries.Count + " 件目" : "";
        txtExecDate.Text = ""; processingDate = null;
        chkDeliveryOutput.IsChecked = false;
        deliveryError.Visibility = Visibility.Collapsed;
        LoadSingleFile(currentFilePath);
    }

    private void LoadSingleFile(string filePath)
    {
        ShowState("main");
        mainPanel.Visibility = Visibility.Visible;
        loadingOverlay.Visibility = Visibility.Visible;
        resultOverlay.Visibility = Visibility.Collapsed;
        FadeInOverlay();

        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args) { args.Result = ReadExcelFile(filePath); };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            if (args.Error != null) { CrossFadeToResult("error", "読込失敗", args.Error.Message); return; }
            var data = args.Result as Dictionary<string, object>;
            if (data != null && data.ContainsKey("error")) { CrossFadeToResult("error", "読込失敗", data["error"].ToString()); return; }
            PopulateForm(data);
            FadeOutOverlay();
        };
        worker.RunWorkerAsync();
    }

    // シート切替時にデータを再読み込みする
    // 現在のファイルを再度開いて、選択中のシートからデータを取得し直す
    private void ReloadSheetData()
    {
        if (excel == null || string.IsNullOrEmpty(currentFilePath)) return;

        loadingOverlay.Visibility = Visibility.Visible;
        resultOverlay.Visibility = Visibility.Collapsed;
        FadeInOverlay();

        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args)
        {
            dynamic wb = null;
            try
            {
                wb = excel.Workbooks.Open(currentFilePath, 0, true);
                args.Result = ReadSheetData(wb, selectedSheetName);
            }
            finally
            {
                if (wb != null)
                    try { wb.Close(false); } catch { }
            }
        };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            if (args.Error == null)
                ApplySheet(args.Result as InsuranceData);
            FadeOutOverlay();
        };
        worker.RunWorkerAsync();
    }

    // ==============================================================
    // Excel読取り
    // ==============================================================

    // Excel COM でファイルからデータを読み取る
    // BackgroundWorker から呼ばれるためUIスレッドからは分離されている
    private Dictionary<string, object> ReadExcelFile(string filePath)
    {
        var result = new Dictionary<string, object>();

        // Excel COM の初回起動（アプリ生存期間中インスタンスを保持）
        if (excel == null)
        {
            var t = Type.GetTypeFromProgID("Excel.Application");
            if (t == null)
                return new Dictionary<string, object> { { "error", "Excelがインストールされていません" } };

            excel = Activator.CreateInstance(t);
            excel.Visible = false;
            excel.DisplayAlerts = false;

            // マクロ無効化（msoAutomationSecurityForceDisable = 3）
            try { excel.AutomationSecurity = 3; } catch { }
            // 自動計算を無効化（xlCalculationManual = -4135）
            try { excel.Calculation = -4135; } catch { }
            // 画面更新・イベント発火を無効化
            try { excel.ScreenUpdating = false; } catch { }
            try { excel.EnableEvents = false; } catch { }
        }

        dynamic wb = null;
        try
        {
            wb = excel.Workbooks.Open(filePath, 0, true); // 読取専用

            // シートフィルタ: filterCell/filterValues で対象シートを絞り込む
            var sheets = new List<string>();
            for (int i = 1; i <= (int)wb.Worksheets.Count; i++)
            {
                dynamic ws = wb.Worksheets[i];
                try
                {
                    string name = (string)ws.Name;
                    if (!string.IsNullOrEmpty(activeProfile.FilterCell) &&
                        activeProfile.FilterValues != null && activeProfile.FilterValues.Length > 0)
                    {
                        try
                        {
                            string cellValue = Convert.ToString(ws.Range[activeProfile.FilterCell].Value2 ?? "");
                            foreach (var fv in activeProfile.FilterValues)
                            {
                                if (cellValue.IndexOf(fv, StringComparison.OrdinalIgnoreCase) >= 0)
                                { sheets.Add(name); break; }
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        // フィルタ未設定時は Visible シートを全て対象
                        if ((int)ws.Visible == -1) sheets.Add(name);
                    }
                }
                catch { }
            }

            result["sheets"] = sheets;
            result["filePath"] = filePath;
            result["fileName"] = System.IO.Path.GetFileName(filePath);

            if (sheets.Count == 0)
            {
                result["noSheet"] = true;
                try { wb.Close(false); } catch { }
                wb = null;
                return result;
            }

            result["selectedSheet"] = sheets[0];

            try
            {
                result["sheetData"] = ReadSheetData(wb, sheets[0]);
            }
            catch (Exception rex)
            {
                result["error"] = "シートデータの読取りに失敗: " + rex.Message;
            }
        }
        catch (Exception ex) { result["error"] = ex.Message; }
        finally
        {
            if (wb != null)
                try { wb.Close(false); } catch { }
        }
        return result;
    }

    // 指定シートから全情報を読み取る
    // A1:AN75 を1回の COM 呼出しで一括取得し、BulkCellByAddress で各セルにアクセス
    private InsuranceData ReadSheetData(dynamic wb, string sheetName)
    {
        dynamic ws = wb.Worksheets[sheetName];
        var d = new InsuranceData();

        // シート全体を1回の COM 呼出しで一括取得（全セルアドレスが A1:AN75 内に収まる）
        object[,] bulk = null;
        try
        {
            object raw = ws.Range["A1", BULK_READ_END].Value2;
            bulk = raw as object[,];
        }
        catch { }

        if (bulk == null) return d;

        // ── 基本情報 ──
        d.AddressNum = BusinessLogic.BulkCellByAddress(bulk, activeProfile.AddressNumberCell);
        d.Name = BusinessLogic.BulkCellByAddress(bulk, activeProfile.NameCell);
        d.KanaName = BusinessLogic.BulkCellByAddress(bulk, activeProfile.KanaNameCell);
        d.Staff = BusinessLogic.BulkCellByAddress(bulk, activeProfile.StaffCell);
        d.Address = BusinessLogic.BulkCellByAddress(bulk, activeProfile.AddressCell);
        d.Birthday = BusinessLogic.BulkCellByAddress(bulk, activeProfile.BirthdayCell);
        d.InstitutionCode = BusinessLogic.BulkCellByAddress(bulk, activeProfile.InstitutionCodeCell);
        d.ContractExists = BusinessLogic.BulkCellByAddress(bulk, activeProfile.ContractExistsCell);

        // 金融機関名の取得とマッピング補正
        string institution = BusinessLogic.BulkCellByAddress(bulk, activeProfile.InstitutionNameCell).Trim();
        if (institutionNameMapping != null)
        {
            string normalized = BusinessLogic.ToFullWidth(institution);
            foreach (var kv in institutionNameMapping)
            {
                if (BusinessLogic.ToFullWidth(kv.Key) == normalized)
                { institution = kv.Value; break; }
            }
        }
        d.InstitutionName = institution;

        // ── 金融機関側情報 ──
        d.RespKanaName = BusinessLogic.BulkCellByAddress(bulk, activeProfile.RespKanaNameCell);
        d.RespName = BusinessLogic.BulkCellByAddress(bulk, activeProfile.RespNameCell);
        d.RespBirthday = BusinessLogic.BulkCellByAddress(bulk, activeProfile.RespBirthdayCell);
        d.RespAddress = BusinessLogic.BulkCellByAddress(bulk, activeProfile.RespAddressCell);

        // ── 契約情報 ──
        d.SeizureRights = BusinessLogic.BulkCellByAddress(bulk, activeProfile.SeizureRightsCell);
        d.PolicyNumber = BusinessLogic.BulkCellByAddress(bulk, activeProfile.PolicyNumberCell);
        d.ContractStatus = BusinessLogic.BulkCellByAddress(bulk, activeProfile.ContractStatusCell);
        d.ContractType = BusinessLogic.BulkCellByAddress(bulk, activeProfile.ContractTypeCell);
        d.InsuredName = BusinessLogic.BulkCellByAddress(bulk, activeProfile.InsuredNameCell);
        d.PaymentFrequency = BusinessLogic.BulkCellByAddress(bulk, activeProfile.PaymentFrequencyCell);

        // 保険料（数値 + 払込区分から表示文字列を生成）
        d.PremiumValue = BusinessLogic.ParseAmount(
            BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.PremiumCell));
        d.PremiumDisplay = BusinessLogic.FormatPremiumDisplay(d.PremiumValue, d.PaymentFrequency);

        // 契約年月日（Excelシリアル値または文字列に対応）
        var contractDateRaw = BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.ContractDateCell);
        d.ContractDateParsed = BusinessLogic.ParseExcelDate(contractDateRaw);
        d.ContractDate = d.ContractDateParsed.HasValue
            ? d.ContractDateParsed.Value.ToString("yyyy/MM/dd")
            : Convert.ToString(contractDateRaw ?? "");

        // 満期年月日
        var maturityDateRaw = BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.MaturityDateCell);
        var maturityParsed = BusinessLogic.ParseExcelDate(maturityDateRaw);
        d.MaturityDate = maturityParsed.HasValue
            ? maturityParsed.Value.ToString("yyyy/MM/dd")
            : Convert.ToString(maturityDateRaw ?? "");

        // ── 金額情報 ──
        d.SurrenderValue = BusinessLogic.ParseAmount(
            BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.SurrenderValueCell));
        d.DividendExists = BusinessLogic.BulkCellByAddress(bulk, activeProfile.DividendExistsCell);
        d.DividendAmount = BusinessLogic.ParseAmount(
            BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.DividendAmountCell));
        d.LoanExists = BusinessLogic.BulkCellByAddress(bulk, activeProfile.LoanExistsCell);
        d.LoanAmount = BusinessLogic.ParseAmount(
            BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.LoanAmountCell));
        d.UnpaidPremiumExists = BusinessLogic.BulkCellByAddress(bulk, activeProfile.UnpaidPremiumExistsCell);
        d.UnpaidPremiumAmount = BusinessLogic.ParseAmount(
            BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.UnpaidPremiumAmountCell));
        d.UnpaidInterestExists = BusinessLogic.BulkCellByAddress(bulk, activeProfile.UnpaidInterestExistsCell);
        d.UnpaidInterestAmount = BusinessLogic.ParseAmount(
            BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.UnpaidInterestAmountCell));
        d.PrepaidPremiumExists = BusinessLogic.BulkCellByAddress(bulk, activeProfile.PrepaidPremiumExistsCell);
        d.PrepaidPremiumAmount = BusinessLogic.ParseAmount(
            BusinessLogic.BulkCellRawByAddress(bulk, activeProfile.PrepaidPremiumAmountCell));

        // 差引見込額
        d.NetValue = BusinessLogic.CalcNetValue(d);

        // ── 給付金額（内部保持、UI非表示） ──
        if (!string.IsNullOrEmpty(activeProfile.BenefitLabelCol) &&
            !string.IsNullOrEmpty(activeProfile.BenefitAmountCol))
        {
            int lblCol = BusinessLogic.ColToIndex(activeProfile.BenefitLabelCol);
            int amtCol = BusinessLogic.ColToIndex(activeProfile.BenefitAmountCol);
            for (int r = activeProfile.BenefitStartRow; r <= activeProfile.BenefitEndRow; r++)
            {
                string label = BusinessLogic.BulkCell(bulk, r, lblCol).Trim();
                if (!string.IsNullOrEmpty(label))
                    d.Benefits.Add(new[] { label,
                        BusinessLogic.FormatBalance(BusinessLogic.BulkCellRaw(bulk, r, amtCol)) });
            }
        }

        // ── 特約（内部保持、UI非表示） ──
        if (!string.IsNullOrEmpty(activeProfile.RiderLabelCol) &&
            !string.IsNullOrEmpty(activeProfile.RiderAmountCol))
        {
            int lblCol = BusinessLogic.ColToIndex(activeProfile.RiderLabelCol);
            int amtCol = BusinessLogic.ColToIndex(activeProfile.RiderAmountCol);
            for (int r = activeProfile.RiderStartRow; r <= activeProfile.RiderEndRow; r++)
            {
                string label = BusinessLogic.BulkCell(bulk, r, lblCol).Trim();
                if (!string.IsNullOrEmpty(label))
                    d.Riders.Add(new[] { label,
                        BusinessLogic.FormatBalance(BusinessLogic.BulkCellRaw(bulk, r, amtCol)) });
            }
        }

        return d;
    }

    // フォームにデータ反映
    private void PopulateForm(Dictionary<string, object> data)
    {
        if (data == null) return;
        fileLink.Text = (data["fileName"] ?? "").ToString();
        var sheets = data["sheets"] as List<string>;
        sheetCombo.Items.Clear();
        if (sheets != null) foreach (var s in sheets) sheetCombo.Items.Add(s);
        if (data.ContainsKey("noSheet") && (bool)data["noSheet"])
        {
            // 対象シートなし → フォームをクリアしてガイドテキストで通知
            ApplySheet(new InsuranceData());
            guideText.Text = "対象シートがありません（フィルタ条件: " +
                (activeProfile.FilterValues != null ? string.Join(", ", activeProfile.FilterValues) : "") + "）";
            btnAdd.IsEnabled = false;
            return;
        }
        // 初期選択時は SelectionChanged による ReloadSheetData を抑止
        suppressSheetChange = true;
        if (sheetCombo.Items.Count > 0) sheetCombo.SelectedIndex = 0;
        suppressSheetChange = false;
        if (sheetCombo.SelectedItem != null) selectedSheetName = sheetCombo.SelectedItem.ToString();
        if (data.ContainsKey("sheetData")) ApplySheet(data["sheetData"] as InsuranceData);
    }

    // InsuranceData を画面に反映する
    private void ApplySheet(InsuranceData d)
    {
        if (d == null) return;
        // 基本情報
        txtAddressNum.Text = BusinessLogic.FormatAddressNumber(d.AddressNum ?? "");
        txtName.Text = (d.Name ?? "").Trim();
        txtInstitution.Text = (d.InstitutionName ?? "").Trim();
        txtStaff.Text = (d.Staff ?? "").Trim();
        txtResidenceAddr.Text = BusinessLogic.FormatAddress(d.Address ?? "");
        txtDeliveryAddr.Text = BusinessLogic.FormatAddress(d.RespAddress ?? "");
        // 契約情報
        txtContractExists.Text = (d.ContractExists ?? "").Trim();
        txtContractExists.Foreground = (d.ContractExists ?? "").Contains("有") ? BrushSuccessIcon : BrushValidationError;
        txtContractStatus.Text = (d.ContractStatus ?? "").Trim();
        txtPolicyNumber.Text = (d.PolicyNumber ?? "").Trim();
        txtContractType.Text = (d.ContractType ?? "").Trim();
        txtContractDate.Text = (d.ContractDate ?? "").Trim();
        currentContractDate = d.ContractDateParsed;
        txtMaturityDate.Text = (d.MaturityDate ?? "").Trim();
        txtPremium.Text = (d.PremiumDisplay ?? "").Trim();
        txtInsuredName.Text = (d.InsuredName ?? "").Trim();
        txtSeizureRights.Text = (d.SeizureRights ?? "").Trim();
        txtSeizureRights.Foreground = (d.SeizureRights ?? "").Contains("有") ? BrushValidationError
            : (d.SeizureRights ?? "").Contains("無") ? BrushSuccessIcon
            : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#333"));
        // 警告バナー
        UpdateWarningBanner(d);
        // 金額情報
        txtSurrenderValue.Text = BusinessLogic.FormatBalance(d.SurrenderValue);
        SetFinancialRow(lblDividend, txtDividend, "配当金", d.DividendExists, d.DividendAmount, false);
        SetFinancialRow(lblLoan, txtLoan, "貸付金", d.LoanExists, d.LoanAmount, true);
        SetFinancialRow(lblUnpaidPremium, txtUnpaidPremium, "未払い保険料", d.UnpaidPremiumExists, d.UnpaidPremiumAmount, true);
        SetFinancialRow(lblUnpaidInterest, txtUnpaidInterest, "未払い利息", d.UnpaidInterestExists, d.UnpaidInterestAmount, true);
        SetFinancialRow(lblPrepaidPremium, txtPrepaidPremium, "前払い保険料", d.PrepaidPremiumExists, d.PrepaidPremiumAmount, false);
        txtNetValue.Text = string.Format("{0:N0}円", d.NetValue);
        UpdateAddButton();
    }

    // 金額情報の1行を設定する（ラベルに有無表示、金額にマイナス表示）
    private void SetFinancialRow(TextBlock lbl, TextBlock val, string baseName,
        string exists, double amount, bool showNegative)
    {
        string suffix = string.IsNullOrEmpty(exists) ? "" : "（" + exists.Trim() + "）";
        lbl.Text = baseName + suffix;
        if (amount == 0 && (exists ?? "").Contains("無"))
        {
            val.Text = "0円";
            val.Foreground = BrushDetailMuted;
        }
        else if (amount == 0 && string.IsNullOrWhiteSpace(exists))
        {
            val.Text = "\u2014";  // em dash
            val.Foreground = BrushDetailMuted;
        }
        else
        {
            double displayAmount = showNegative ? -Math.Abs(amount) : amount;
            val.Text = string.Format("{0:N0}円", displayAmount);
            val.Foreground = showNegative && amount > 0 ? BrushValidationError : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#333"));
        }
    }

    // 警告バナーの表示制御
    private void UpdateWarningBanner(InsuranceData d)
    {
        string contractExists = (d.ContractExists ?? "").Trim();
        string seizureRights = (d.SeizureRights ?? "").Trim();
        if (contractExists == "無")
        {
            // エラーレベル: 保険契約なし
            warningIcon.Text = "\u2717"; warningIcon.Foreground = BrushValidationError;
            warningText.Text = "保険契約がありません"; warningText.Foreground = BrushValidationError;
            warningBanner.Visibility = Visibility.Visible;
        }
        else if (seizureRights.Contains("有"))
        {
            // 注意レベル: 差押権利者あり
            warningIcon.Text = "\u26A0"; warningIcon.Foreground = BrushValidationError;
            warningText.Text = "差押権利者が存在します"; warningText.Foreground = BrushValidationError;
            warningBanner.Visibility = Visibility.Visible;
        }
        else
        {
            warningBanner.Visibility = Visibility.Collapsed;
        }
    }

    // ==============================================================
    // 処理結果オーバーレイ
    // ==============================================================

    private void ShowResult(string type, string title, string detail)
    {
        overlayPanel.Visibility = Visibility.Visible; loadingOverlay.Visibility = Visibility.Collapsed; resultOverlay.Visibility = Visibility.Visible;
        resultTitle.Text = title;
        resultDetail.Text = detail;
        resultDetail.Visibility = string.IsNullOrEmpty(detail) ? Visibility.Collapsed : Visibility.Visible;
        bool last = currentFileIndex >= fileEntries.Count - 1;
        if (type == "success")
        {
            resultIcon.Text = "\u2713";
            resultIcon.Foreground = BrushSuccessIcon;
            resultIconBg.Background = BrushIconBgSuccess;
            resultDetail.Foreground = BrushAccent;
        }
        else if (type == "skip")
        {
            resultIcon.Text = "\u2192";
            resultIcon.Foreground = BrushAccent;
            resultIconBg.Background = BrushIconBgSkip;
            resultDetail.Foreground = BrushDetailMuted;
        }
        else
        {
            resultIcon.Text = "\u2717";
            resultIcon.Foreground = BrushValidationError;
            resultIconBg.Background = BrushIconBgError;
            resultDetail.Foreground = BrushValidationError;
        }
        resultButton.Content = last ? "完了" : "次のファイルへ \u2192";
        resultSub.Text = last ? "" : ((currentFileIndex + 2) + " / " + fileEntries.Count + " 件目へ進みます");
        resultSub.Visibility = last ? Visibility.Collapsed : Visibility.Visible;
        resultButton.Focus();
    }

    private void ProceedToNext()
    {
        FadeOutOverlay(delegate { LoadFileAtIndex(currentFileIndex + 1); });
    }

    private void ExecuteSkip()
    {
        if (currentFileIndex >= 0 && currentFileIndex < fileEntries.Count)
            fileEntries[currentFileIndex].State = FileProcessState.Skipped;
        ShowResult("skip", "スキップしました", System.IO.Path.GetFileName(currentFilePath));
        FadeInOverlay();
    }

    // ==============================================================
    // 処理実行（一覧に追加・スキップ）
    // ==============================================================

    // 「一覧に追加」ボタン押下時の処理
    // バリデーション → UI入力値をDictionaryに収集 → BackgroundWorkerで非同期処理
    private void ExecuteAdd()
    {
        if (!processingDate.HasValue) return;

        // 届出住所50文字チェック（チェックONの場合のみ）
        if (chkDeliveryOutput.IsChecked == true)
        {
            string deliveryFull = "（届出：" + txtDeliveryAddr.Text.Trim() + "）";
            if (deliveryFull.Length > 50)
            { deliveryError.Visibility = Visibility.Visible; return; }
        }

        // オーバーレイ表示（処理中スピナー）
        loadingOverlay.Visibility = Visibility.Visible;
        resultOverlay.Visibility = Visibility.Collapsed;
        FadeInOverlay();

        // UI入力値をDictionaryに収集（BackgroundWorkerに渡すため）
        string deliveryAddr = (chkDeliveryOutput.IsChecked == true)
            ? "（届出：" + txtDeliveryAddr.Text.Trim() + "）" : "";
        var addData = new Dictionary<string, string>
        {
            { "addressNum",    txtAddressNum.Text.Trim() },
            { "name",          txtName.Text.Trim() },
            { "staff",         txtStaff.Text.Trim() },
            { "institution",   txtInstitution.Text.Trim() },
            { "residenceAddr", txtResidenceAddr.Text.Trim() },
            { "deliveryAddr",  deliveryAddr },
            { "execDate",      BusinessLogic.DateToWareki(processingDate.Value, eraMapping) },
            { "policyNumber",  txtPolicyNumber.Text.Trim() },
            { "contractType",  txtContractType.Text.Trim() },
            { "insuredName",   txtInsuredName.Text.Trim() },
            { "filePath",      currentFilePath },
            { "fileName",      System.IO.Path.GetFileName(currentFilePath) }
        };
        var contractDate = currentContractDate;

        var worker = new BackgroundWorker();
        worker.DoWork += delegate(object s, DoWorkEventArgs args)
        {
            args.Result = ProcessAdd(addData, contractDate);
        };
        worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
        {
            if (args.Error != null)
            { CrossFadeToResult("error", "処理失敗", args.Error.Message); return; }
            var result = args.Result as Dictionary<string, string>;
            if (result["status"] == "ok")
            {
                fileEntries[currentFileIndex].State = FileProcessState.Added;
                CrossFadeToResult("success", "一覧に追加しました",
                    "文書番号: " + result["docNumber"]);
            }
            else
            { CrossFadeToResult("error", "処理失敗", result["message"]); }
        };
        worker.RunWorkerAsync();
    }

    // 一覧への追加処理を実行する（BackgroundWorker から呼ばれる）
    // 処理フロー:
    //   1. 文書番号の排他ロック採番
    //   2. 差押文言6フィールドの生成
    //   3. CSV行の構築（21列、RFC 4180準拠エスケープ）
    //   4. CSV追記（排他ロック付き FileStream、リトライ最大5回）
    //   5. 印刷用ファイル保存
    private Dictionary<string, string> ProcessAdd(
        Dictionary<string, string> addData, DateTime? contractDate)
    {
        var result = new Dictionary<string, string>();

        // 1. 文書番号の排他ロック採番
        string docNumber;
        if (!AllocateDocNumber(out docNumber))
        {
            result["status"] = "error";
            result["message"] = "文書番号の取得に失敗";
            return result;
        }

        // 2. 差押文言の生成
        var seizureText = GenerateSeizureText(
            addData["institution"], addData["contractType"],
            addData["policyNumber"], contractDate,
            addData["name"], addData["insuredName"]);

        // 3. CSV行の構築（21列）
        string printFileName = docNumber + ".xlsm";
        var csvFields = new[]
        {
            DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),  // 登録日時
            addData["addressNum"],                           // 宛名番号
            addData["name"],                                 // 氏名
            addData["staff"],                                // 職員名
            addData["execDate"],                             // 執行日（7桁和暦）
            addData["residenceAddr"],                        // 住民票住所
            addData["deliveryAddr"],                          // 届出住所
            addData["institution"],                           // 保険会社名
            addData["policyNumber"],                          // 証券番号
            addData["contractType"],                          // 契約種類
            seizureText["Line1"],                             // 差押文言1（宣言文）
            seizureText["Line2"],                             // 差押文言2（保険の種類）
            seizureText["Line3"],                             // 差押文言3（証券番号）
            seizureText["Line4"],                             // 差押文言4（契約年月日）
            seizureText["Line5"],                             // 差押文言5（契約者）
            seizureText["Line6"],                             // 差押文言6（被保険者）
            docNumber,                                        // 文書番号
            printFileName,                                    // 照会結果ファイル名
            "",                                               // 処理済フラグ1
            "",                                               // 処理済フラグ2
            ""                                                // 処理済フラグ3
        };
        string csvLine = string.Join(",", csvFields.Select(f => BusinessLogic.CsvEscape(f)));

        // 4. CSV追記
        string csvPath = System.IO.Path.Combine(activeProfile.OutputFolder, CSV_FILENAME);
        if (!WriteCsvLine(csvPath, csvLine))
        {
            RollbackDocNumber();
            result["status"] = "error";
            result["message"] = "CSV書き込み失敗";
            return result;
        }

        // 5. 印刷用ファイル保存
        string printFilePath = System.IO.Path.Combine(activeProfile.PrintFolder, printFileName);
        SavePrintFile(addData["filePath"], printFilePath);

        result["status"] = "ok";
        result["docNumber"] = docNumber;
        return result;
    }

    // ==============================================================
    // 差押文言生成
    // ==============================================================

    // 契約情報から差押文言6フィールドを生成する
    //
    // Line1: 宣言文（保険会社名を全角で埋め込み）
    // Line2: 保険の種類（契約種類を全角変換）
    // Line3: 証券番号（全角変換）
    // Line4: 保険契約年月日（和暦全角、ゼロ埋めなし）
    // Line5: 契約者名
    // Line6: 被保険者名
    private Dictionary<string, string> GenerateSeizureText(
        string institutionName, string contractType,
        string policyNumber, DateTime? contractDate,
        string contractorName, string insuredName)
    {
        string normalizedInst = BusinessLogic.ToFullWidth((institutionName ?? "").Trim());

        string line1 = "\u3000上記滞納者が、債務者である" + normalizedInst
            + "に対して有する下記生命保険契約に基づく一切の支払請求権。"
            + "なお、保険事故発生又は契約期間の終了しないうちに解約されたときは"
            + "解約返戻金全額の支払請求権。";

        string line2 = BusinessLogic.ToFullWidth((contractType ?? "").Trim());
        string line3 = BusinessLogic.ToFullWidth((policyNumber ?? "").Trim());
        string line4 = contractDate.HasValue
            ? BusinessLogic.FormatWarekiDisplay(contractDate.Value, eraMapping) : "";
        string line5 = (contractorName ?? "").Trim();
        string line6 = (insuredName ?? "").Trim();

        return new Dictionary<string, string>
        {
            { "Line1", line1 }, { "Line2", line2 }, { "Line3", line3 },
            { "Line4", line4 }, { "Line5", line5 }, { "Line6", line6 }
        };
    }

    // ==============================================================
    // 印刷用ファイル保存
    // ==============================================================

    // 印刷用ファイルを保存する
    // 元ファイルを SaveAs で複製し、選択シート以外を削除（VeryHidden は保持）
    private void SavePrintFile(string sourcePath, string destPath)
    {
        if (destPath.Length > MAX_PATH) return;
        var destDir = System.IO.Path.GetDirectoryName(destPath);
        if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

        dynamic printWorkbook = null;
        try
        {
            excel.Visible = false;
            printWorkbook = excel.Workbooks.Open(sourcePath, 0, false); // 読み書き可能

            // SaveAs 前にブック保護を解除
            try { if ((bool)printWorkbook.ProtectStructure) printWorkbook.Unprotect(); } catch { }

            // SaveAs（xlOpenXMLWorkbookMacroEnabled = 52）
            excel.DisplayAlerts = false;
            printWorkbook.SaveAs(destPath, 52);
            excel.DisplayAlerts = true;

            try { if ((bool)printWorkbook.ProtectStructure) printWorkbook.Unprotect(); } catch { }

            // 選択シート以外を削除（VeryHidden シートは残す）
            excel.DisplayAlerts = false;
            for (int i = (int)printWorkbook.Worksheets.Count; i >= 1; i--)
            {
                dynamic ws = printWorkbook.Worksheets[i];
                try
                {
                    // VeryHidden（Visible=2）以外で、選択シート名と異なるシートを削除
                    if ((int)ws.Visible != 2 && (string)ws.Name != selectedSheetName)
                        ws.Delete();
                }
                catch { }
            }
            excel.DisplayAlerts = true;

            // 表示位置を A1・スクロール先頭にリセット
            try
            {
                dynamic ps = printWorkbook.Worksheets[selectedSheetName];
                ps.Activate();
                printWorkbook.Application.ActiveWindow.ScrollRow = 1;
                printWorkbook.Application.ActiveWindow.ScrollColumn = 1;
                ps.Range["A1"].Select();
            }
            catch { }

            printWorkbook.Save();
        }
        catch { }
        finally
        {
            if (printWorkbook != null)
                try { printWorkbook.Close(false); } catch { }
        }
    }

    // ==============================================================
    // CSV 操作
    // ==============================================================

    private bool WriteCsvLine(string csvPath, string csvLine)
    {
        var dir = System.IO.Path.GetDirectoryName(csvPath);
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        bool isNewFile = !File.Exists(csvPath);
        for (int retry = 1; retry <= CSV_WRITE_MAX_RETRY; retry++)
        {
            try
            {
                using (var fs = new FileStream(csvPath, FileMode.Append, FileAccess.Write, FileShare.None))
                using (var w = new StreamWriter(fs, new UTF8Encoding(true)))
                {
                    if (isNewFile) w.WriteLine(CSV_HEADER);
                    w.WriteLine(csvLine);
                    w.Flush();
                }
                return true;
            }
            catch { if (retry < CSV_WRITE_MAX_RETRY) System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS); }
        }
        return false;
    }

    // ==============================================================
    // 文書番号管理
    // ==============================================================

    private bool AllocateDocNumber(out string docNum)
    {
        docNum = "";
        string counterPath = System.IO.Path.Combine(exeDir, DOC_NUMBER_FILE);
        for (int retry = 1; retry <= CSV_WRITE_MAX_RETRY; retry++)
        {
            try
            {
                using (var fs = new FileStream(counterPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    var bytes = new byte[fs.Length];
                    fs.Read(bytes, 0, bytes.Length);
                    var json = Encoding.UTF8.GetString(bytes).TrimStart('\uFEFF');
                    int nextNumber = JsonHelper.GetInt(json, "nextNumber", 1);
                    docNum = nextNumber.ToString();
                    lastDocNumber = docNum;
                    fs.Seek(0, SeekOrigin.Begin); fs.SetLength(0);
                    var newBytes = Encoding.UTF8.GetBytes("{\n    \"nextNumber\":  " + (nextNumber + 1) + "\n}");
                    fs.Write(newBytes, 0, newBytes.Length);
                }
                return true;
            }
            catch { if (retry < CSV_WRITE_MAX_RETRY) System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS); }
        }
        return false;
    }

    private void RollbackDocNumber()
    {
        if (string.IsNullOrEmpty(lastDocNumber)) return;
        string counterPath = System.IO.Path.Combine(exeDir, DOC_NUMBER_FILE);
        for (int retry = 1; retry <= CSV_WRITE_MAX_RETRY; retry++)
        {
            try
            {
                using (var fs = new FileStream(counterPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    fs.Seek(0, SeekOrigin.Begin); fs.SetLength(0);
                    var bytes = Encoding.UTF8.GetBytes("{\n    \"nextNumber\":  " + lastDocNumber + "\n}");
                    fs.Write(bytes, 0, bytes.Length);
                }
                return;
            }
            catch { if (retry < CSV_WRITE_MAX_RETRY) System.Threading.Thread.Sleep(CSV_WRITE_RETRY_INTERVAL_MS); }
        }
    }

    // ==============================================================
    // Excel COM クリーンアップ
    // ==============================================================

    private void CleanupExcel()
    {
        if (excel == null) return;
        try
        {
            try
            {
                while ((int)excel.Workbooks.Count > 0)
                { dynamic wb = excel.Workbooks[1]; try { wb.Close(false); } catch { } }
            }
            catch { }
            try { excel.ScreenUpdating = true; } catch { }
            try { excel.DisplayAlerts = true; } catch { }
            excel.Quit();
            try { Marshal.ReleaseComObject(excel); } catch { }
        }
        catch { }
        finally
        {
            excel = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }

    // ==============================================================
    // XAML 定義
    // ==============================================================

    private Window BuildWindow()
    {
        string xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='生命保険差押予定一覧 作成ツール' Width='1000' Height='700' MinWidth='900' MinHeight='520'
    WindowStartupLocation='CenterScreen' Background='#F9F9F9' FontFamily='Meiryo UI'
    UseLayoutRounding='True' SnapsToDevicePixels='True'>
<Window.Resources>
    <Style TargetType='TextBox'>
        <Setter Property='Foreground' Value='#333'/>
        <Setter Property='Template'><Setter.Value>
            <ControlTemplate TargetType='TextBox'>
                <Border Background='{TemplateBinding Background}'
                        BorderBrush='{TemplateBinding BorderBrush}'
                        BorderThickness='{TemplateBinding BorderThickness}'
                        CornerRadius='4' Padding='{TemplateBinding Padding}'
                        SnapsToDevicePixels='True'>
                    <ScrollViewer x:Name='PART_ContentHost' Focusable='False'/></Border>
            </ControlTemplate>
        </Setter.Value></Setter>
    </Style>
    <Style TargetType='ComboBox'>
        <Setter Property='Foreground' Value='#333'/><Setter Property='Background' Value='White'/>
        <Setter Property='BorderBrush' Value='#D0D0D0'/><Setter Property='BorderThickness' Value='1'/>
        <Setter Property='Padding' Value='6,5'/><Setter Property='Cursor' Value='Hand'/>
        <Setter Property='Template'><Setter.Value>
            <ControlTemplate TargetType='ComboBox'>
                <Grid x:Name='comboRoot'>
                    <ToggleButton BorderThickness='0' Background='Transparent' Focusable='False' ClickMode='Press'
                        IsChecked='{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}'>
                        <ToggleButton.Template><ControlTemplate TargetType='ToggleButton'>
                            <Border Background='Transparent'/></ControlTemplate></ToggleButton.Template>
                    </ToggleButton>
                    <Border x:Name='bd' Background='{TemplateBinding Background}'
                            BorderBrush='{TemplateBinding BorderBrush}'
                            BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='4' IsHitTestVisible='False'>
                        <Grid Margin='{TemplateBinding Padding}'>
                            <Grid.ColumnDefinitions><ColumnDefinition Width='*'/><ColumnDefinition Width='20'/></Grid.ColumnDefinitions>
                            <ContentPresenter Content='{TemplateBinding SelectionBoxItem}'
                                ContentTemplate='{TemplateBinding SelectionBoxItemTemplate}'
                                HorizontalAlignment='Left' VerticalAlignment='Center'/>
                            <Path Grid.Column='1' Data='M0,0 L4,4 8,0' Stroke='#888' StrokeThickness='1.5'
                                  VerticalAlignment='Center' HorizontalAlignment='Center'/>
                        </Grid></Border>
                    <Popup x:Name='PART_Popup' AllowsTransparency='True' Placement='Bottom'
                           IsOpen='{TemplateBinding IsDropDownOpen}'>
                        <Border Background='White' BorderBrush='#D0D0D0' BorderThickness='1'
                                CornerRadius='4' Margin='0,2,0,0' Padding='0,4'
                                MinWidth='{Binding ActualWidth, ElementName=comboRoot}'>
                            <Border.Effect><DropShadowEffect BlurRadius='8' ShadowDepth='2' Opacity='0.12'/></Border.Effect>
                            <ScrollViewer MaxHeight='200'><StackPanel IsItemsHost='True'/></ScrollViewer>
                        </Border></Popup>
                </Grid>
                <ControlTemplate.Triggers>
                    <Trigger Property='IsMouseOver' Value='True'>
                        <Setter TargetName='bd' Property='BorderBrush' Value='#E69500'/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value></Setter>
    </Style>
    <Style x:Key='AB' TargetType='Button'>
        <Setter Property='Background' Value='#E69500'/><Setter Property='Foreground' Value='White'/>
        <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
        <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderThickness' Value='0'/>
        <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='Button'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' CornerRadius='4' Padding='{TemplateBinding Padding}'>
                <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='bd' Property='Background' Value='#CC8400'/></Trigger>
                <Trigger Property='IsEnabled' Value='False'><Setter TargetName='bd' Property='Background' Value='#CCC'/><Setter Property='Foreground' Value='#999'/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
    <Style x:Key='GB' TargetType='Button'>
        <Setter Property='Background' Value='White'/><Setter Property='Foreground' Value='#555'/>
        <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
        <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderBrush' Value='#D0D0D0'/><Setter Property='BorderThickness' Value='1'/>
        <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='Button'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}'
                BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='4' Padding='{TemplateBinding Padding}'>
                <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='bd' Property='Background' Value='#FFF6E8'/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
    <Style TargetType='CheckBox'>
        <Setter Property='Foreground' Value='#333'/><Setter Property='Cursor' Value='Hand'/>
        <Setter Property='Template'><Setter.Value>
            <ControlTemplate TargetType='CheckBox'>
                <StackPanel Orientation='Horizontal'>
                    <Border x:Name='cbBox' Width='16' Height='16' CornerRadius='3'
                            Background='White' BorderBrush='#C8C8C8' BorderThickness='1'
                            VerticalAlignment='Center' Margin='0,0,6,0'>
                        <Path x:Name='cbCheck' Data='M2.5,7 L5.5,10 L11.5,3.5' Stroke='White'
                              StrokeThickness='2' Visibility='Collapsed'/></Border>
                    <ContentPresenter VerticalAlignment='Center'/></StackPanel>
                <ControlTemplate.Triggers>
                    <Trigger Property='IsChecked' Value='True'>
                        <Setter TargetName='cbBox' Property='Background' Value='#E69500'/>
                        <Setter TargetName='cbBox' Property='BorderBrush' Value='#E69500'/>
                        <Setter TargetName='cbCheck' Property='Visibility' Value='Visible'/></Trigger>
                    <Trigger Property='IsMouseOver' Value='True'>
                        <Setter TargetName='cbBox' Property='BorderBrush' Value='#E69500'/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value></Setter>
    </Style>
</Window.Resources>
<DockPanel>
    <Border DockPanel.Dock='Top' Background='#E69500' Padding='18,10'>
        <TextBlock Text='生命保険差押予定一覧 作成ツール' FontSize='13' FontWeight='Medium' Foreground='White'/></Border>
    <Border DockPanel.Dock='Bottom' Background='#F0F0F0' BorderBrush='#E0E0E0' BorderThickness='0,1,0,0' Padding='18,4'>
        <DockPanel>
                <StackPanel DockPanel.Dock='Right' Orientation='Horizontal'>
                    <TextBlock Text='出力先: ' FontSize='11' Foreground='#666'/>
                    <TextBlock x:Name='StatusRight' FontSize='11' Foreground='#666'/></StackPanel>
                <TextBlock x:Name='StatusLeft' FontSize='11' Foreground='#666'/></DockPanel></Border>
    <Grid>
        <!-- 初期画面 -->
        <Grid x:Name='InitialPanel'>
            <Border Background='White' BorderBrush='#D4B88A' BorderThickness='2' CornerRadius='8'
                    Margin='80,60' VerticalAlignment='Center' HorizontalAlignment='Center' Padding='60,40'>
                <StackPanel HorizontalAlignment='Center'>
                    <TextBlock Text='&#x1F4C2;' FontSize='36' HorizontalAlignment='Center' Margin='0,0,0,12'/>
                    <TextBlock Text='ここにファイルをドラッグ＆ドロップ' FontSize='14' Foreground='#666' HorizontalAlignment='Center' Margin='0,0,0,16'/>
                    <Button x:Name='BtnSelectFile' Style='{StaticResource AB}' HorizontalAlignment='Center'>
                        <TextBlock Text='ファイルを選択' FontSize='13'/></Button>
                </StackPanel></Border></Grid>
        <!-- メインフォーム -->
        <Grid x:Name='MainPanel' Visibility='Collapsed' Margin='18,14,18,12'>
            <Grid.RowDefinitions>
                <RowDefinition Height='Auto'/><RowDefinition Height='Auto'/>
                <RowDefinition Height='*'/><RowDefinition Height='Auto'/></Grid.RowDefinitions>
            <!-- シート選択 -->
            <DockPanel Grid.Row='0' Margin='0,0,0,10'>
                <TextBlock Text='シート:' VerticalAlignment='Center' Foreground='#555' FontSize='11' Margin='0,0,6,0'/>
                <ComboBox x:Name='SheetCombo' MinWidth='180' FontSize='12'/>
                <StackPanel DockPanel.Dock='Right' Orientation='Horizontal' HorizontalAlignment='Right'>
                    <TextBlock Text='&#x1F4C4; ' FontSize='11' Foreground='#E69500' VerticalAlignment='Center'/>
                    <TextBlock x:Name='FileLink' FontSize='11' Foreground='#E69500'
                               Cursor='Hand' TextDecorations='Underline' VerticalAlignment='Center'/>
                    <Button x:Name='BtnReload' Style='{StaticResource GB}' Padding='8,4' Margin='8,0,0,0' FontSize='11'>
                        <TextBlock Text='&#x1F504; 再読み込み'/></Button>
                </StackPanel>
            </DockPanel>
            <!-- 基本情報 -->
            <Border Grid.Row='1' Background='White' BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='6' Padding='16,14' Margin='0,0,0,10'>
                <StackPanel><TextBlock Text='&#x1F464; 基本情報' FontSize='13' Foreground='#E69500' FontWeight='Medium' Margin='0,0,0,10'/>
                <Grid><Grid.ColumnDefinitions><ColumnDefinition Width='*'/><ColumnDefinition Width='16'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                    <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='6'/>
                        <RowDefinition Height='Auto'/><RowDefinition Height='6'/><RowDefinition Height='Auto'/></Grid.RowDefinitions>
                    <Grid Grid.Row='0' Grid.Column='0'><Grid.ColumnDefinitions><ColumnDefinition Width='120'/><ColumnDefinition Width='16'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                        <StackPanel><TextBlock Text='宛名番号' FontSize='11' Foreground='#777' Margin='0,0,0,3'/>
                            <TextBox x:Name='TxtAddressNum' IsReadOnly='True' Background='#F3F3F3' BorderBrush='#E8E8E8' FontFamily='Consolas' FontSize='12' Padding='5,4'/></StackPanel>
                        <StackPanel Grid.Column='2'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>氏名 &#x270E;</TextBlock>
                            <TextBox x:Name='TxtName' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel></Grid>
                    <Grid Grid.Row='0' Grid.Column='2'><Grid.ColumnDefinitions><ColumnDefinition Width='2*'/><ColumnDefinition Width='16'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                        <StackPanel><TextBlock Text='保険会社' FontSize='11' Foreground='#777' Margin='0,0,0,3'/>
                            <TextBox x:Name='TxtInstitution' IsReadOnly='True' Background='#F3F3F3' BorderBrush='#E8E8E8' FontSize='12' Padding='5,4'/></StackPanel>
                        <StackPanel Grid.Column='2'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>処分担当 &#x270E;</TextBlock>
                            <TextBox x:Name='TxtStaff' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel></Grid>
                    <StackPanel Grid.Row='2' Grid.Column='0'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>住民票住所 &#x270E;</TextBlock>
                        <TextBox x:Name='TxtResidenceAddr' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/></StackPanel>
                    <StackPanel Grid.Row='2' Grid.Column='2'><TextBlock FontSize='11' Foreground='#777' Margin='0,0,0,3'>届出住所 &#x270E;</TextBlock>
                        <TextBox x:Name='TxtDeliveryAddr' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0'/>
                        <TextBlock x:Name='DeliveryError' Foreground='#D32F2F' FontSize='10' Visibility='Collapsed' Margin='0,1,0,0'/></StackPanel>
                    <StackPanel Grid.Row='4' Grid.Column='0'><TextBlock Text='執行日' FontSize='11' Foreground='#777' Margin='0,0,0,3'/>
                        <StackPanel Orientation='Horizontal'>
                            <TextBox x:Name='TxtExecDate' FontSize='12' Padding='5,4' BorderBrush='#D0D0D0' FontFamily='Consolas' Width='120'/>
                            <Button x:Name='BtnCalendar' Style='{StaticResource GB}' Padding='6,4' Margin='4,0,0,0'>
                                <TextBlock Text='&#x1F4C5;' FontSize='13'/></Button>
                            <Popup x:Name='CalendarPopup' StaysOpen='False' Placement='Bottom' AllowsTransparency='True'>
                                <Border Background='White' BorderBrush='#D0D0D0' BorderThickness='1'
                                        CornerRadius='6' Padding='8' Margin='0,4,0,0'>
                                    <Border.Effect><DropShadowEffect BlurRadius='12' ShadowDepth='3' Opacity='0.15'/></Border.Effect>
                                    <Calendar x:Name='DateCalendar' SelectionMode='SingleDate'/></Border>
                            </Popup>
                        </StackPanel></StackPanel>
                    <StackPanel Grid.Row='4' Grid.Column='2' VerticalAlignment='Top' Margin='0,18,0,0'>
                        <CheckBox x:Name='ChkDeliveryOutput' Content='届出住所を差押通知書に出力する' FontSize='12'/></StackPanel>
                </Grid></StackPanel></Border>
            <!-- 契約情報 + 金額情報 -->
            <Grid Grid.Row='2'>
                <Grid.ColumnDefinitions><ColumnDefinition Width='*'/><ColumnDefinition Width='8'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>
                <Border Grid.Column='0' Background='White' BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='6' Padding='14,12' Margin='0,0,0,10'>
                    <StackPanel>
                        <DockPanel Margin='0,0,0,8'>
                            <StackPanel x:Name='WarningBanner' DockPanel.Dock='Right' Orientation='Horizontal' Visibility='Collapsed'>
                                <TextBlock x:Name='WarningIcon' FontSize='11' Margin='0,0,4,0'/>
                                <TextBlock x:Name='WarningText' FontSize='11'/></StackPanel>
                            <TextBlock Text='&#x1F4CB; 契約情報' FontSize='13' Foreground='#E69500' FontWeight='Medium'/></DockPanel>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='保険契約の有無' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtContractExists' Foreground='#333' FontSize='11.5' FontWeight='Medium' HorizontalAlignment='Right'/></DockPanel></Border>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='契約の状態' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtContractStatus' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='証券番号' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtPolicyNumber' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='契約種類' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtContractType' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='契約年月日' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtContractDate' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='満期年月日' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtMaturityDate' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='保険料' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtPremium' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                        <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='被保険者' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtInsuredName' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                        <DockPanel Margin='0,4,0,0'><TextBlock Text='差押権利者' Foreground='#777' FontSize='11.5'/>
                            <TextBlock x:Name='TxtSeizureRights' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel>
                    </StackPanel></Border>
                <Border Grid.Column='2' Background='White' BorderBrush='#E0E0E0' BorderThickness='1' CornerRadius='6' Padding='14,12' Margin='0,0,0,10'>
                    <DockPanel>
                        <TextBlock DockPanel.Dock='Top' Text='&#x1F4B0; 金額情報' FontSize='13' Foreground='#E69500' FontWeight='Medium' Margin='0,0,0,8'/>
                        <Border DockPanel.Dock='Bottom' BorderBrush='#E69500' BorderThickness='0,2,0,0' Padding='0,6,0,0' Margin='0,6,0,0'>
                            <DockPanel><TextBlock Text='参考: 差引見込額' FontSize='11.5' Foreground='#777' VerticalAlignment='Center'/>
                                <TextBlock x:Name='TxtNetValue' FontSize='11.5' Foreground='#E69500' FontWeight='Medium' HorizontalAlignment='Right'/></DockPanel></Border>
                        <StackPanel>
                            <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock Text='解約返戻金' Foreground='#777' FontSize='11.5'/>
                                <TextBlock x:Name='TxtSurrenderValue' Foreground='#333' FontSize='11.5' FontWeight='Medium' HorizontalAlignment='Right'/></DockPanel></Border>
                            <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock x:Name='LblDividend' Text='配当金' Foreground='#777' FontSize='11.5'/>
                                <TextBlock x:Name='TxtDividend' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                            <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock x:Name='LblLoan' Text='貸付金' Foreground='#777' FontSize='11.5'/>
                                <TextBlock x:Name='TxtLoan' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                            <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock x:Name='LblUnpaidPremium' Text='未払い保険料' Foreground='#777' FontSize='11.5'/>
                                <TextBlock x:Name='TxtUnpaidPremium' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                            <Border BorderBrush='#F0F0F0' BorderThickness='0,0,0,1' Padding='0,4'><DockPanel><TextBlock x:Name='LblUnpaidInterest' Text='未払い利息' Foreground='#777' FontSize='11.5'/>
                                <TextBlock x:Name='TxtUnpaidInterest' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel></Border>
                            <DockPanel Margin='0,4,0,0'><TextBlock x:Name='LblPrepaidPremium' Text='前払い保険料' Foreground='#777' FontSize='11.5'/>
                                <TextBlock x:Name='TxtPrepaidPremium' Foreground='#333' FontSize='11.5' HorizontalAlignment='Right'/></DockPanel>
                        </StackPanel>
                    </DockPanel></Border>
            </Grid>
            <!-- アクションボタン -->
            <DockPanel Grid.Row='3'>
                <Button x:Name='BtnLoadFile' DockPanel.Dock='Left' Style='{StaticResource GB}'><TextBlock Text='ファイルを読み込む'/></Button>
                <StackPanel DockPanel.Dock='Right' Orientation='Horizontal' HorizontalAlignment='Right'>
                    <TextBlock x:Name='GuideText' VerticalAlignment='Center' FontSize='11' Foreground='#D32F2F' Margin='0,0,12,0'/>
                    <Button x:Name='BtnSkip' Style='{StaticResource GB}' Margin='0,0,8,0'><TextBlock Text='&#x2192; スキップ'/></Button>
                    <Button x:Name='BtnAdd' Style='{StaticResource AB}' IsEnabled='False'><TextBlock Text='&#xFF0B; 一覧に追加'/></Button>
                </StackPanel></DockPanel>
        </Grid>
        <!-- オーバーレイ -->
        <Grid x:Name='OverlayPanel' Visibility='Collapsed' Background='#CCFFFFFF'>
            <Grid x:Name='LoadingOverlay' Visibility='Collapsed' HorizontalAlignment='Center' VerticalAlignment='Center'>
                <Path x:Name='SpinnerPath' Data='M 20,2 A 18,18 0 1 1 2,20'
                      Stroke='#E69500' StrokeThickness='3'
                      StrokeStartLineCap='Round' StrokeEndLineCap='Round'
                      Width='40' Height='40' Stretch='None'
                      RenderTransformOrigin='0.5,0.5'>
                    <Path.RenderTransform><RotateTransform/></Path.RenderTransform>
                </Path></Grid>
            <Grid x:Name='ResultOverlay' Visibility='Collapsed' HorizontalAlignment='Center' VerticalAlignment='Center'>
                <Border Background='White' CornerRadius='10' Padding='48,36' MinWidth='350'>
                    <Border.Effect><DropShadowEffect BlurRadius='16' ShadowDepth='4' Opacity='0.12'/></Border.Effect>
                    <StackPanel HorizontalAlignment='Center'>
                        <Border x:Name='ResultIconBg' Width='64' Height='64' CornerRadius='32' Background='#E8F5ED' HorizontalAlignment='Center' Margin='0,0,0,20'>
                            <TextBlock x:Name='ResultIcon' Text='&#x2713;' FontSize='32' HorizontalAlignment='Center' VerticalAlignment='Center' Foreground='#107C41'/></Border>
                        <TextBlock x:Name='ResultTitle' Text='' FontSize='17' FontWeight='Medium' HorizontalAlignment='Center' Margin='0,0,0,8'/>
                        <TextBlock x:Name='ResultDetail' FontSize='12' Foreground='#999' HorizontalAlignment='Center'/>
                        <Button x:Name='ResultButton' Style='{StaticResource AB}' HorizontalAlignment='Center' Padding='24,10' Margin='0,24,0,0'>
                            <TextBlock Text='次のファイルへ' FontSize='13'/></Button>
                        <TextBlock x:Name='ResultSub' FontSize='11' Foreground='#999' HorizontalAlignment='Center' Margin='0,12,0,0'/>
                    </StackPanel></Border></Grid>
        </Grid>
    </Grid>
</DockPanel></Window>";
        using (var reader = XmlReader.Create(new StringReader(xaml))) { return (Window)XamlReader.Load(reader); }
    }
}
// ==============================================================
// RelayCommand
// ==============================================================

public class RelayCommand : ICommand
{
    private Action<object> execute;
    public RelayCommand(Action<object> execute) { this.execute = execute; }
    public event EventHandler CanExecuteChanged { add {} remove {} }
    public bool CanExecute(object p) { return true; }
    public void Execute(object p) { execute(p); }
}