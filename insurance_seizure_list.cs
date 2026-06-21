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

    // --- キャッシュ済みブラシ（アンバー #E69500 ベース） ---
    private static readonly SolidColorBrush BrushBorderNormal    = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D0D0D0")));
    private static readonly SolidColorBrush BrushValidationError = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D32F2F")));
    private static readonly SolidColorBrush BrushSuccessIcon     = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#107C41")));
    private static readonly SolidColorBrush BrushAccent          = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E69500")));
    private static readonly SolidColorBrush BrushIconBgSuccess   = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F5ED")));
    private static readonly SolidColorBrush BrushIconBgError     = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FCEBEB")));
    private static readonly SolidColorBrush BrushIconBgSkip      = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF3E0")));
    private static readonly SolidColorBrush BrushDetailMuted     = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#999")));
    private static readonly SolidColorBrush BrushWarningBanner   = Frozen(new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CC8400")));

    // ブラシを Freeze して描画パフォーマンスを向上するヘルパー
    private static SolidColorBrush Frozen(SolidColorBrush brush) { brush.Freeze(); return brush; }

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

    // --- UI要素（Phase 2 で定義） ---
    private Window window;

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
        var validationErrors = new List<string>();
        if (string.IsNullOrWhiteSpace(activeProfile.AddressNumberCell))
            validationErrors.Add("addressNumberCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.NameCell))
            validationErrors.Add("nameCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.StaffCell))
            validationErrors.Add("staffCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.AddressCell))
            validationErrors.Add("addressCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.InstitutionNameCell))
            validationErrors.Add("institutionNameCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.PolicyNumberCell))
            validationErrors.Add("policyNumberCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.SurrenderValueCell))
            validationErrors.Add("surrenderValueCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.ContractExistsCell))
            validationErrors.Add("contractExistsCell が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.OutputFolder))
            validationErrors.Add("outputFolder が未設定です");
        if (string.IsNullOrWhiteSpace(activeProfile.PrintFolder))
            validationErrors.Add("printFolder が未設定です");
        if (validationErrors.Count > 0)
        {
            MessageBox.Show("プロファイル設定エラー:\n\n" + string.Join("\n", validationErrors),
                "設定エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        // --- 出力先フォルダの自動作成 ---
        try
        {
            if (!Directory.Exists(activeProfile.OutputFolder))
                Directory.CreateDirectory(activeProfile.OutputFolder);
            if (!Directory.Exists(activeProfile.PrintFolder))
                Directory.CreateDirectory(activeProfile.PrintFolder);
        }
        catch (Exception dirEx)
        {
            MessageBox.Show(
                "出力先フォルダの作成に失敗しました。\n\n" + dirEx.Message,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        // --- 起動モード判定 ---
        if (StartupArgs != null && StartupArgs.Length > 0)
        {
            isFromFileSearch = true;
            foreach (var arg in StartupArgs)
                if (File.Exists(arg)) fileEntries.Add(new FileEntry { FilePath = arg, State = FileProcessState.Pending });
        }

        // --- Phase 1 確認用: 設定読込成功 ---
        MessageBox.Show(
            "Phase 1: 設定読込テスト完了\n\n"
            + "プロファイル: " + activeProfile.Name + "\n"
            + "元号マッピング: " + eraMapping.Count + " 件\n"
            + "保険会社名マッピング: " + (institutionNameMapping != null ? institutionNameMapping.Count + " 件" : "なし") + "\n"
            + "出力先: " + activeProfile.OutputFolder,
            "Phase 1", MessageBoxButton.OK, MessageBoxImage.Information);
        Shutdown();
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