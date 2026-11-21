using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace MultiReplace;

public sealed record Rule(string Find, string Replace);

public enum Status { Changed, Unchanged, Skipped, Error }

/// <param name="Count">内容の置換箇所数（ファイル名置換は含まない）</param>
/// <param name="RenamedTo">ファイル名を変更した（する予定の）フルパス</param>
public sealed record FileResult(string Path, Status Status, int Count, string? Message, string? RenamedTo = null);

/// <param name="Cells">セルの値（共有文字列・インライン文字列・テーブル見出し）</param>
/// <param name="SheetNames">シート名（数式・定義名・グラフ等の参照も追従）</param>
/// <param name="Objects">コメント・図形・グラフ内のテキスト</param>
/// <param name="Numbers">数値セル（保存されている値の文字列に対して置換。例: 123→456）</param>
public sealed record ExcelOptions(bool Cells, bool SheetNames, bool Objects, bool Numbers);

public sealed record EngineOptions(
    bool Apply, bool Backup,
    bool ReplaceContent, bool ReplaceName, bool NameWithExt,
    ExcelOptions Excel,
    bool ReplaceDir = false);

/// <summary>
/// 複数の置換ワードを「同時に」適用する（A→B と B→A の入れ替えも可能）。
/// regex=false: 文字列そのまま（長い語を優先）。regex=true: 正規表現（先に書いた行を優先、置換後は $1 などが使える）。
/// </summary>
public sealed class Replacer
{
    private readonly Regex? _re;
    private readonly Regex[]? _each; // 正規表現モード: ルールごとの \G 付きRegex（置換後の $1 展開用）
    private readonly string[] _repl;

    public bool IsRegex { get; }

    public Replacer(IEnumerable<Rule> rules, bool caseSensitive, bool regex = false)
    {
        IsRegex = regex;
        var cmp = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        var q = rules
            .Where(r => !string.IsNullOrEmpty(r.Find))
            .GroupBy(r => r.Find, cmp)
            .Select(g => g.First());
        var list = (regex ? q : q.OrderByDescending(r => r.Find.Length)).ToList(); // 文字列モードは長い語を優先
        _repl = list.Select(r => r.Replace ?? "").ToArray();
        if (list.Count == 0) return;

        var opt = RegexOptions.CultureInvariant | (caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);
        if (!regex)
        {
            var pattern = string.Join("|", list.Select((r, i) => $"(?<r{i}>{Regex.Escape(r.Find)})"));
            _re = new Regex(pattern, opt);
            return;
        }

        opt |= RegexOptions.Multiline; // ^ $ を行単位にする
        var timeout = TimeSpan.FromSeconds(10);  // バックトラック暴走対策
        _each = new Regex[list.Count];
        for (int i = 0; i < list.Count; i++)
        {
            try
            {
                var plain = new Regex(list[i].Find, opt, timeout);
                if (plain.IsMatch(""))
                    throw new ArgumentException("空の文字列にマッチするパターンは使えません");
                _each[i] = new Regex(@"\G(?:" + list[i].Find + ")", opt, timeout);
            }
            catch (ArgumentException ex)
            {
                throw new ArgumentException($"正規表現が不正です: \"{list[i].Find}\"（{ex.Message}）");
            }
        }
        _re = new Regex(string.Join("|", list.Select((r, i) => $"(?<r{i}>{r.Find})")), opt, timeout);
    }

    public bool IsEmpty => _re == null;

    public Match[] Matches(string s) =>
        _re == null ? Array.Empty<Match>() : _re.Matches(s).Where(m => m.Length > 0).ToArray();

    public bool IsMatch(string s) => _re != null && _re.IsMatch(s);

    /// <param name="input">Match の元になった文字列</param>
    public string ReplacementFor(Match m, string input)
    {
        for (int i = 0; i < _repl.Length; i++)
        {
            if (!m.Groups["r" + i].Success) continue;
            if (_each == null) return _repl[i];
            var em = _each[i].Match(input, m.Index);
            return em.Success ? em.Result(_repl[i]) : m.Value;
        }
        return m.Value;
    }

    public string Replace(string s, out int count)
    {
        int c = 0;
        if (_re == null) { count = 0; return s; }
        var r = _re.Replace(s, m =>
        {
            if (m.Length == 0) return "";
            c++;
            return ReplacementFor(m, s);
        });
        count = c;
        return r;
    }

    /// <summary>\n \r \t \\ のエスケープ記法を実際の文字に変換する。</summary>
    public static string Unescape(string s)
    {
        if (s.IndexOf('\\') < 0) return s;
        var sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '\\' && i + 1 < s.Length)
            {
                switch (s[i + 1])
                {
                    case 'n': sb.Append('\n'); i++; continue;
                    case 'r': sb.Append('\r'); i++; continue;
                    case 't': sb.Append('\t'); i++; continue;
                    case '\\': sb.Append('\\'); i++; continue;
                }
            }
            sb.Append(s[i]);
        }
        return sb.ToString();
    }
}

public static class Engine
{
    /// <param name="excludes">
    /// 除外パターン（ワイルドカード可）。ファイル名に対して判定する。
    /// 末尾が \ または / のものはフォルダ名の除外（対象フォルダ以下のどの階層のフォルダ名でも判定）。
    /// </param>
    public static List<string> Enumerate(string dir, string[] patterns, bool recurse, string[]? excludes = null)
    {
        var opt = new EnumerationOptions
        {
            RecurseSubdirectories = recurse,
            IgnoreInaccessible = true,
            MatchType = MatchType.Simple, // "*.xls" が .xlsx に化ける旧仕様を避ける
            MatchCasing = MatchCasing.CaseInsensitive,
        };
        var ex = new ExcludeFilter(dir, excludes);
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in patterns.Length == 0 ? new[] { "*" } : patterns)
            foreach (var f in Directory.EnumerateFiles(dir, p, opt))
                if (!Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal) && !ex.File(f))
                    set.Add(f);
        var res = set.ToList();
        res.Sort(StringComparer.OrdinalIgnoreCase);
        return res;
    }

    /// <summary>
    /// ディレクトリ名の置換。対象フォルダそのものは対象外。
    /// 深い階層から順に処理する（親の名前が変わっても、先に処理した子のパスが無効にならないように）。
    /// ファイルの処理（中身・ファイル名）はこれより先に済ませておくこと。
    /// </summary>
    public static List<FileResult> RenameDirectories(string baseDir, Replacer rp, EngineOptions o, bool recurse,
        string[]? excludes, HashSet<string> planned, CancellationToken ct = default)
    {
        var opt = new EnumerationOptions
        {
            RecurseSubdirectories = recurse,
            IgnoreInaccessible = true,
            MatchType = MatchType.Simple,
            MatchCasing = MatchCasing.CaseInsensitive,
        };
        var ex = new ExcludeFilter(baseDir, excludes);
        var dirs = Directory.EnumerateDirectories(baseDir, "*", opt)
            .Where(d => !ex.Dir(d))
            .OrderByDescending(d => d.Count(ch => ch == Path.DirectorySeparatorChar || ch == Path.AltDirectorySeparatorChar))
            .ThenBy(d => d, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var results = new List<FileResult>();
        foreach (var path in dirs)
        {
            if (ct.IsCancellationRequested) break;
            try
            {
                var name = Path.GetFileName(path);
                var newName = rp.Replace(name, out int c);
                if (c == 0 || newName == name) continue;

                if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
                {
                    results.Add(new(path, Status.Skipped, 0, "ジャンクション/シンボリックリンクのためスキップ"));
                    continue;
                }
                if (newName.Length == 0 || newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
                    || newName.EndsWith(' ') || newName.EndsWith('.'))
                {
                    results.Add(new(path, Status.Error, 0, $"変更後のフォルダ名が不正です: \"{newName}\""));
                    continue;
                }
                var newPath = Path.Combine(Path.GetDirectoryName(path)!, newName);
                bool sameIgnoringCase = string.Equals(path, newPath, StringComparison.OrdinalIgnoreCase);
                if (!sameIgnoringCase && (Directory.Exists(newPath) || File.Exists(newPath)))
                {
                    results.Add(new(path, Status.Error, 0, $"同名のフォルダ/ファイルが既に存在するため変更しません: {newName}"));
                    continue;
                }
                if (!planned.Add(newPath.ToUpperInvariant()))
                {
                    results.Add(new(path, Status.Error, 0, $"他の変更後の名前と重複するため変更しません: {newName}"));
                    continue;
                }
                if (o.Apply) MoveCaseSafe(path, newPath, isDir: true);
                results.Add(new(path, Status.Changed, 0, null, newPath));
            }
            catch (Exception e)
            {
                results.Add(new(path, Status.Error, 0, "名前変更に失敗: " + e.Message));
            }
        }
        return results;
    }

    /// <summary>大文字小文字だけの変更（Windowsでは同一パス扱い）は、一時名を経由して移動する。</summary>
    internal static void MoveCaseSafe(string from, string to, bool isDir)
    {
        if (string.Equals(from, to, StringComparison.OrdinalIgnoreCase))
        {
            var tmp = from + ".rn-" + Guid.NewGuid().ToString("N");
            if (isDir) { Directory.Move(from, tmp); Directory.Move(tmp, to); }
            else { File.Move(from, tmp); File.Move(tmp, to); }
        }
        else if (isDir) Directory.Move(from, to);
        else File.Move(from, to);
    }

    private sealed class ExcludeFilter
    {
        private readonly string _base;
        private readonly List<string> _file = new();
        private readonly List<string> _dir = new();

        public ExcludeFilter(string baseDir, string[]? excludes)
        {
            _base = baseDir;
            foreach (var raw in excludes ?? Array.Empty<string>())
            {
                var p = raw.Trim();
                if (p.Length == 0) continue;
                if (p.EndsWith('\\') || p.EndsWith('/')) { p = p.TrimEnd('\\', '/'); if (p.Length > 0) _dir.Add(p); }
                else _file.Add(p);
            }
        }

        private static bool M(string pattern, string name) =>
            System.IO.Enumeration.FileSystemName.MatchesSimpleExpression(pattern, name, true);

        private bool AncestorExcluded(string? parent)
        {
            if (_dir.Count == 0 || parent == null) return false;
            var rel = Path.GetRelativePath(_base, parent);
            if (rel == ".") return false;
            foreach (var seg in rel.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
                foreach (var p in _dir)
                    if (M(p, seg)) return true;
            return false;
        }

        public bool File(string full) =>
            _file.Any(p => M(p, Path.GetFileName(full))) || AncestorExcluded(Path.GetDirectoryName(full));

        public bool Dir(string full)
        {
            var name = Path.GetFileName(full);
            return _file.Any(p => M(p, name)) || _dir.Any(p => M(p, name)) || AncestorExcluded(Path.GetDirectoryName(full));
        }
    }

    /// <param name="planned">ファイル名変更の衝突検出用（呼び出し側が1回の実行で使い回す）</param>
    public static FileResult ProcessFile(string path, Replacer rp, EngineOptions o, HashSet<string> planned)
    {
        FileResult content = new(path, Status.Unchanged, 0, null);
        if (o.ReplaceContent)
        {
            try { content = ProcessContent(path, rp, o); }
            catch (Exception ex) { content = new(path, Status.Error, 0, ex.Message); }
            if (content.Status == Status.Error) return content;
        }
        if (!o.ReplaceName) return content;
        try { return Rename(path, rp, o, planned, content); }
        catch (Exception ex) { return new(path, Status.Error, content.Count, "名前変更に失敗: " + ex.Message); }
    }

    private static FileResult ProcessContent(string path, Replacer rp, EngineOptions o)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is ".xlsx" or ".xlsm" or ".xltx" or ".xltm")
            return ExcelProcessor.Process(path, rp, o.Apply, o.Backup, o.Excel);
        if (ext is ".xls" or ".xlsb")
            return new(path, Status.Skipped, 0, "未対応の形式（xlsx / xlsm に変換してください）");
        return TextProcessor.Process(path, rp, o.Apply, o.Backup);
    }

    private static FileResult Rename(string path, Replacer rp, EngineOptions o, HashSet<string> planned, FileResult content)
    {
        string dir = Path.GetDirectoryName(path)!;
        string fname = Path.GetFileName(path);
        string stem = o.NameWithExt ? fname : Path.GetFileNameWithoutExtension(fname);
        string ext = o.NameWithExt ? "" : Path.GetExtension(fname);

        string newStem = rp.Replace(stem, out int c);
        if (c == 0 || newStem == stem) return content;

        string newName = newStem + ext;
        if (newName.Length == 0 || newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
            || newName.EndsWith(' ') || newName.EndsWith('.'))
            return new(path, Status.Error, content.Count, $"変更後のファイル名が不正です: \"{newName}\"");

        string newPath = Path.Combine(dir, newName);
        bool sameIgnoringCase = string.Equals(path, newPath, StringComparison.OrdinalIgnoreCase);
        if (!sameIgnoringCase)
        {
            if (File.Exists(newPath) || Directory.Exists(newPath))
                return new(path, Status.Error, content.Count, $"同名のファイルが既に存在するため変更しません: {newName}");
        }
        if (!planned.Add(newPath.ToUpperInvariant()))
            return new(path, Status.Error, content.Count, $"他のファイルの変更後の名前と重複するため変更しません: {newName}");

        if (o.Apply) MoveCaseSafe(path, newPath, isDir: false);
        return new(path, Status.Changed, content.Count, content.Message, newPath);
    }

    internal static void MakeBackup(string path)
    {
        File.Copy(path, path + ".bak", overwrite: true);
    }
}

/// <summary>
/// テキストファイルの置換。文字コード（UTF-8 / BOM付きUTF-8 / UTF-16 / Shift_JIS(CP932)）は
/// 自動判別して同じコードで書き戻す。文字列のまま置換するので改行コード（CRLF/LF/CR混在含む）も維持される。
/// </summary>
internal static class TextProcessor
{
    private static readonly UTF8Encoding Utf8Strict = new(false, true);

    public static FileResult Process(string path, Replacer rp, bool apply, bool backup)
    {
        byte[] bytes = File.ReadAllBytes(path);
        Encoding enc;
        int bom = 0;
        byte[] preamble = Array.Empty<byte>();

        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            enc = Utf8Strict; bom = 3; preamble = new byte[] { 0xEF, 0xBB, 0xBF };
        }
        else if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            enc = new UnicodeEncoding(false, false, true); bom = 2; preamble = new byte[] { 0xFF, 0xFE };
        }
        else if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            enc = new UnicodeEncoding(true, false, true); bom = 2; preamble = new byte[] { 0xFE, 0xFF };
        }
        else
        {
            int n = Math.Min(bytes.Length, 8192);
            for (int i = 0; i < n; i++)
                if (bytes[i] == 0)
                    return new(path, Status.Skipped, 0, "バイナリファイルのためスキップ");
            try { Utf8Strict.GetString(bytes); enc = Utf8Strict; }
            catch (DecoderFallbackException)
            {
                enc = Encoding.GetEncoding(932, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            }
        }

        string text;
        try { text = enc.GetString(bytes, bom, bytes.Length - bom); }
        catch (DecoderFallbackException)
        {
            return new(path, Status.Skipped, 0, "文字コードを判別できないためスキップ");
        }

        var newText = rp.Replace(text, out int cnt);
        if (cnt == 0 || newText == text) return new(path, Status.Unchanged, 0, null);

        if (apply)
        {
            byte[] body;
            try { body = enc.GetBytes(newText); }
            catch (EncoderFallbackException)
            {
                return new(path, Status.Error, 0, "置換後の文字を元の文字コードで表現できません");
            }
            if (backup) Engine.MakeBackup(path);
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            fs.Write(preamble, 0, preamble.Length);
            fs.Write(body, 0, body.Length);
        }
        return new(path, Status.Changed, cnt, null);
    }
}

internal static class ExcelProcessor
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    private static readonly HashSet<string> FormulaNames = new()
    {
        "f", "formula", "formula1", "formula2", "definedName",
        "calculatedColumnFormula", "totalsRowFormula",
    };

    private static bool IsXml(string n, string prefix) =>
        n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    private sealed class Flags
    {
        public bool CellStrings;   // <si>(共有文字列) / <is>(インライン文字列)
        public bool CommentText;   // コメント <text>
        public bool DrawingText;   // 図形・グラフ <a:p>
        public bool TableCols;     // テーブル列見出し
        public bool Numbers;       // 数値セル <c><v>123</v></c>
        public bool Refs;          // シート名変更に伴う参照の書き換え
        public bool Workbook;      // workbook.xml のシート名
        public bool App;           // docProps/app.xml のシート名一覧
    }

    public static FileResult Process(string path, Replacer rp, bool apply, bool backup, ExcelOptions xo)
    {
        if (!xo.Cells && !xo.SheetNames && !xo.Objects && !xo.Numbers)
            return new(path, Status.Skipped, 0, "Excelの置換対象が1つも選ばれていません");

        int total = 0;
        var ctx = new Ctx();
        var modified = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        string? tmp = null;

        try
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (xo.SheetNames)
                {
                    var err = BuildSheetMap(zip, rp, map);
                    if (err != null) return new(path, Status.Error, 0, err);
                }

                foreach (var e in zip.Entries)
                {
                    var n = e.FullName;
                    bool isShared = n.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase);
                    bool isSheet = IsXml(n, "xl/worksheets/");
                    bool isComment = IsXml(n, "xl/comments");
                    bool isDrawing = IsXml(n, "xl/drawings/");
                    bool isChart = IsXml(n, "xl/charts/");
                    bool isTable = IsXml(n, "xl/tables/");
                    bool isWb = n.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase);
                    bool isApp = n.Equals("docProps/app.xml", StringComparison.OrdinalIgnoreCase);

                    var f = new Flags
                    {
                        CellStrings = xo.Cells && (isShared || isSheet),
                        CommentText = xo.Objects && isComment,
                        DrawingText = xo.Objects && (isDrawing || isChart),
                        TableCols = xo.Cells && isTable,
                        Numbers = xo.Numbers && isSheet,
                        Refs = map.Count > 0 && (isSheet || isChart || isTable || isWb),
                        Workbook = map.Count > 0 && isWb,
                        App = map.Count > 0 && isApp,
                    };
                    if (!(f.CellStrings || f.CommentText || f.DrawingText || f.TableCols || f.Numbers || f.Refs || f.Workbook || f.App))
                        continue;

                    byte[] data;
                    using (var s = e.Open())
                    using (var ms = new MemoryStream())
                    {
                        s.CopyTo(ms);
                        data = ms.ToArray();
                    }

                    if (isSheet)
                    {
                        // 対象になり得る要素が無いシートは解析を省略（大きなシートの高速化）
                        var head = Encoding.UTF8.GetString(data);
                        bool needInline = f.CellStrings && (head.Contains("<is>") || head.Contains("<is "));
                        bool needRef = f.Refs && head.Contains('!');
                        // 数値はXML上にそのままの文字列で入っているので、一致が無ければ解析不要
                        bool needNum = f.Numbers && (rp.IsRegex || rp.IsMatch(head)); // 正規表現は ^$ などXML全体に一致しないことがあるので省略しない
                        if (!needInline && !needRef && !needNum) continue;
                    }

                    XDocument doc;
                    using (var ms = new MemoryStream(data))
                        doc = XDocument.Load(ms, LoadOptions.PreserveWhitespace);

                    var (c, changed) = ProcessDoc(f, doc, rp, map, ctx);
                    if (!changed) continue;
                    total += c;
                    modified[e.FullName] = Serialize(doc);
                }

                if (modified.Count == 0) return new(path, Status.Unchanged, 0, null);
                if (!apply) return new(path, Status.Changed, total, null);

                if (backup) Engine.MakeBackup(path);
                tmp = path + ".tmp-" + Guid.NewGuid().ToString("N");
                using (var outFs = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                using (var outZip = new ZipArchive(outFs, ZipArchiveMode.Create))
                {
                    foreach (var e in zip.Entries)
                    {
                        var ne = outZip.CreateEntry(e.FullName, CompressionLevel.Optimal);
                        ne.LastWriteTime = e.LastWriteTime;
                        using var dst = ne.Open();
                        if (modified.TryGetValue(e.FullName, out var nd))
                            dst.Write(nd, 0, nd.Length);
                        else
                        {
                            using var src = e.Open();
                            src.CopyTo(dst);
                        }
                    }
                }
            }

            File.Move(tmp, path, overwrite: true);
            tmp = null;
            return new(path, Status.Changed, total, null);
        }
        catch (InvalidDataException)
        {
            return new(path, Status.Skipped, 0, "xlsxとして開けません（パスワード保護 or 破損の可能性）");
        }
        finally
        {
            if (tmp != null && File.Exists(tmp)) { try { File.Delete(tmp); } catch { } }
        }
    }

    private static string? ValidateSheetName(string n)
    {
        if (n.Length == 0) return "空になります";
        if (n.Length > 31) return "31文字を超えます";
        if (n.IndexOfAny(new[] { '\\', '/', '?', '*', '[', ']', ':' }) >= 0) return "使用できない文字( \\ / ? * [ ] : )が含まれます";
        if (n[0] == '\'' || n[^1] == '\'') return "先頭・末尾にアポストロフィは使えません";
        return null;
    }

    /// <summary>旧シート名→新シート名の対応表を作る。不正な名前や重複があればエラー文字列を返す。</summary>
    private static string? BuildSheetMap(ZipArchive zip, Replacer rp, Dictionary<string, string> map)
    {
        var entry = zip.GetEntry("xl/workbook.xml");
        if (entry == null) return null;
        XDocument doc;
        using (var s = entry.Open())
            doc = XDocument.Load(s);

        var finalNames = new List<string>();
        foreach (var sh in doc.Descendants(Main + "sheet"))
        {
            var old = (string?)sh.Attribute("name");
            if (old == null) continue;
            var nw = rp.Replace(old, out int c);
            if (c == 0 || nw == old) { finalNames.Add(old); continue; }
            var why = ValidateSheetName(nw);
            if (why != null) return $"シート名 \"{old}\" → \"{nw}\" は不正です（{why}）";
            map[old] = nw;
            finalNames.Add(nw);
        }
        if (finalNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() != finalNames.Count)
        {
            map.Clear();
            return "シート名を置換すると同名のシートができてしまうため、このファイルは変更しません";
        }
        return null;
    }

    private static byte[] Serialize(XDocument doc)
    {
        using var ms = new MemoryStream();
        using (var xw = XmlWriter.Create(ms, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            Indent = false,
        }))
        {
            doc.Save(xw);
        }
        return ms.ToArray();
    }

    private sealed class Ctx { }

    private static (int count, bool changed) ProcessDoc(Flags f, XDocument doc, Replacer rp, Dictionary<string, string> map, Ctx ctx)
    {
        int total = 0;
        bool changed = false;
        if (doc.Root == null) return (0, false);

        // セルの文字列: <si>(共有文字列) / <is>(インライン文字列)
        if (f.CellStrings)
        {
            foreach (var el in doc.Descendants().Where(x => x.Name.Namespace == Main &&
                         (x.Name.LocalName == "si" || x.Name.LocalName == "is")).ToList())
            {
                var runs = el.Elements(Main + "t").Concat(el.Elements(Main + "r").Elements(Main + "t")).ToList();
                int c = ReplaceInRuns(runs, rp, true);
                if (c > 0) { total += c; changed = true; }
            }
        }

        // 数値セル: <c><v>123</v></c>（t属性なし or t="n"、数式セルは除く）
        // 保存されている値の文字列（例 "123", "0.5"）に対して置換する。文字列セルの処理より後に行う（変換後のセルを二重に置換しない）
        if (f.Numbers)
        {
            foreach (var cell in doc.Descendants(Main + "c").ToList())
            {
                var t = (string?)cell.Attribute("t");
                if (t != null && t != "n") continue;
                if (cell.Element(Main + "f") != null) continue;
                var v = cell.Element(Main + "v");
                if (v == null || v.HasElements) continue;
                var nv = rp.Replace(v.Value, out int c);
                if (c == 0 || nv == v.Value) continue;
                if (double.TryParse(nv, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var d) && double.IsFinite(d))
                {
                    v.Value = nv; // 数値のまま
                }
                else
                {
                    // 数値として不正になる場合（123→ABC など）は文字列セルに変換する
                    var tEl = new XElement(Main + "t", nv);
                    if (nv.Length > 0 && (char.IsWhiteSpace(nv[0]) || char.IsWhiteSpace(nv[^1])))
                        tEl.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                    cell.SetAttributeValue("t", "inlineStr");
                    v.Remove();
                    cell.Add(new XElement(Main + "is", tEl));
                }
                total += c; changed = true;
            }
        }

        // コメント
        if (f.CommentText)
        {
            foreach (var el in doc.Descendants(Main + "text").ToList())
            {
                var runs = el.Elements(Main + "t").Concat(el.Elements(Main + "r").Elements(Main + "t")).ToList();
                int c = ReplaceInRuns(runs, rp, true);
                if (c > 0) { total += c; changed = true; }
            }
        }

        // 図形・グラフ内テキスト: <a:p> 内の <a:r>/<a:t>
        if (f.DrawingText)
        {
            foreach (var p in doc.Descendants(A + "p").ToList())
            {
                var runs = p.Elements(A + "r").Elements(A + "t").ToList();
                int c = ReplaceInRuns(runs, rp, false);
                if (c > 0) { total += c; changed = true; }
            }
        }

        // テーブルの列見出し（セルの見出し文字列と一致している必要がある）
        if (f.TableCols)
        {
            foreach (var col in doc.Descendants(Main + "tableColumn").ToList())
            {
                var a = col.Attribute("name");
                if (a == null) continue;
                var nv = rp.Replace(a.Value, out int c);
                if (c > 0 && nv != a.Value) { a.Value = nv; total += c; changed = true; }
            }
        }

        // シート名そのもの（workbook.xml）
        if (f.Workbook)
        {
            foreach (var sh in doc.Descendants(Main + "sheet").ToList())
            {
                var a = sh.Attribute("name");
                if (a != null && map.TryGetValue(a.Value, out var nn))
                {
                    a.Value = nn; total++; changed = true;
                }
            }
        }

        // シート名を参照している数式・定義名・ハイパーリンクなど
        if (f.Refs)
        {
            foreach (var el in doc.Descendants().ToList())
            {
                if (FormulaNames.Contains(el.Name.LocalName) && !el.HasElements && el.Value.Contains('!'))
                {
                    var nv = RewriteRefs(el.Value, map);
                    if (nv != el.Value) { el.Value = nv; changed = true; }
                }
                else if (el.Name.LocalName == "hyperlink")
                {
                    var loc = el.Attribute("location");
                    if (loc != null && loc.Value.Contains('!'))
                    {
                        var nv = RewriteRefs(loc.Value, map);
                        if (nv != loc.Value) { loc.Value = nv; changed = true; }
                    }
                }
            }
        }

        // docProps/app.xml のシート名一覧（Excelが再生成するが整合させておく）
        if (f.App)
        {
            foreach (var el in doc.Descendants(VT + "lpstr").ToList())
            {
                if (!el.HasElements && map.TryGetValue(el.Value, out var nn))
                {
                    el.Value = nn; changed = true;
                }
            }
        }

        return (total, changed);
    }

    private static readonly Regex StrLit = new("\"(?:[^\"]|\"\")*\"", RegexOptions.Compiled);
    private static readonly Regex SheetRef = new(
        @"(?<q>'(?:[^']|'')+')!|(?<![\w.$'])(?<u>[A-Za-z_\u0080-￿][\w.\u0080-￿]*)!",
        RegexOptions.Compiled);

    /// <summary>数式中の「シート名!」を新しいシート名に書き換える（文字列リテラル内は触らない）。</summary>
    internal static string RewriteRefs(string formula, Dictionary<string, string> map)
    {
        var sb = new StringBuilder();
        int last = 0;
        foreach (Match lit in StrLit.Matches(formula))
        {
            sb.Append(RewriteSeg(formula.Substring(last, lit.Index - last), map));
            sb.Append(lit.Value);
            last = lit.Index + lit.Length;
        }
        sb.Append(RewriteSeg(formula.Substring(last), map));
        return sb.ToString();
    }

    private static string RewriteSeg(string seg, Dictionary<string, string> map) =>
        SheetRef.Replace(seg, m =>
        {
            string old = m.Groups["q"].Success
                ? m.Groups["q"].Value[1..^1].Replace("''", "'")
                : m.Groups["u"].Value;
            if (!map.TryGetValue(old, out var nn)) return m.Value;
            return "'" + nn.Replace("'", "''") + "'!";
        });

    /// <summary>
    /// 1つの文字列を構成する複数の書式ラン（リッチテキスト）をまたいだ一致も置換する。
    /// 置換後の文字列は一致の開始位置のランに入るので、書式はできるだけ保たれる。
    /// </summary>
    private static int ReplaceInRuns(List<XElement> runs, Replacer rp, bool setSpace)
    {
        if (runs.Count == 0) return 0;
        var texts = runs.Select(r => r.Value).ToArray();
        string full = string.Concat(texts);
        var ms = rp.Matches(full);
        if (ms.Length == 0) return 0;

        bool anyChange = false;
        int pos = 0, k = 0;
        for (int i = 0; i < runs.Count; i++)
        {
            int s = pos, e = pos + texts[i].Length;
            pos = e;
            var sb = new StringBuilder();
            int p = s;
            while (p < e)
            {
                while (k < ms.Length && ms[k].Index + ms[k].Length <= p) k++;
                if (k < ms.Length && ms[k].Index <= p)
                {
                    if (ms[k].Index == p) sb.Append(rp.ReplacementFor(ms[k], full));
                    p = Math.Min(e, ms[k].Index + ms[k].Length);
                }
                else
                {
                    int next = k < ms.Length ? Math.Min(e, ms[k].Index) : e;
                    sb.Append(full, p, next - p);
                    p = next;
                }
            }
            var nt = sb.ToString();
            if (nt != texts[i])
            {
                anyChange = true;
                runs[i].Value = nt;
                if (setSpace && nt.Length > 0 && (char.IsWhiteSpace(nt[0]) || char.IsWhiteSpace(nt[^1])))
                    runs[i].SetAttributeValue(XNamespace.Xml + "space", "preserve");
            }
        }
        return anyChange ? ms.Length : 0;
    }
}
