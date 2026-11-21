using System.Text.Json;

namespace MultiReplace;

public sealed class Settings
{
    public string Dir { get; set; } = "";
    public string Pattern { get; set; } = "*.txt;*.xlsx";
    public string Exclude { get; set; } = "*.bak";
    public bool Sub { get; set; } = true;
    public bool CaseSensitive { get; set; } = true;
    public bool Escape { get; set; } = false;
    public bool Regex { get; set; } = false;
    public bool Backup { get; set; } = true;
    public bool DoContent { get; set; } = true;
    public bool DoName { get; set; } = false;
    public bool DoDir { get; set; } = false;
    public bool NameWithExt { get; set; } = false;
    public bool XlCells { get; set; } = true;
    public bool XlSheetNames { get; set; } = false;
    public bool XlObjects { get; set; } = true;
    public bool XlNumbers { get; set; } = true;
    public List<Rule> Rules { get; set; } = new();
}

public sealed class MainForm : Form
{
    private readonly TextBox txtDir = new() { Dock = DockStyle.Fill };
    private readonly TextBox txtPattern = new() { Dock = DockStyle.Fill };
    private readonly TextBox txtExclude = new() { Dock = DockStyle.Fill, Text = "*.bak" };
    private readonly TextBox txtLog = new()
    {
        Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
        WordWrap = false, Font = new Font("Consolas", 9F), BackColor = Color.White,
    };
    private readonly CheckBox chkSub = new() { Text = "サブフォルダも対象にする", AutoSize = true, Checked = true };
    private readonly CheckBox chkCase = new() { Text = "大文字/小文字を区別する", AutoSize = true, Checked = true };
    private readonly CheckBox chkRegex = new() { Text = "正規表現を使う（置換後は $1 ${名前}）", AutoSize = true };
    private readonly CheckBox chkEscape = new() { Text = @"\n \t などのエスケープを使う", AutoSize = true };
    private readonly CheckBox chkBackup = new() { Text = "バックアップ(.bak)を作る", AutoSize = true, Checked = true };
    private readonly CheckBox chkContent = new() { Text = "ファイルの中身", AutoSize = true, Checked = true };
    private readonly CheckBox chkName = new() { Text = "ファイル名", AutoSize = true };
    private readonly CheckBox chkDir = new() { Text = "ディレクトリ名（フォルダ名）", AutoSize = true };
    private readonly CheckBox chkNameExt = new() { Text = "拡張子もファイル名置換の対象にする", AutoSize = true };
    private readonly CheckBox chkXlCells = new() { Text = "セルの値", AutoSize = true, Checked = true };
    private readonly CheckBox chkXlNum = new() { Text = "数値セル（123 → 456 など）", AutoSize = true, Checked = true };
    private readonly CheckBox chkXlSheet = new() { Text = "シート名（数式などの参照も追従）", AutoSize = true };
    private readonly CheckBox chkXlObj = new() { Text = "コメント・図形・グラフ内の文字", AutoSize = true, Checked = true };
    private readonly GroupBox gbExcel = new() { Text = "Excel（.xlsx/.xlsm）の置換対象", AutoSize = true, Dock = DockStyle.Fill };
    private readonly GroupBox gbMode = new() { Text = "置換するもの", AutoSize = true, Dock = DockStyle.Fill };
    private readonly DataGridView grid = new();
    private readonly Button btnBrowse = new() { Text = "参照...", AutoSize = true };
    private readonly Button btnPreview = new() { Text = "プレビュー（置換しない）", AutoSize = true };
    private readonly Button btnRun = new() { Text = "置換を実行", AutoSize = true };
    private readonly Button btnCancel = new() { Text = "中止", AutoSize = true, Enabled = false };
    private readonly ProgressBar bar = new() { Width = 220, Height = 22 };

    private CancellationTokenSource? _cts;

    private static string SettingsPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MultiReplace", "settings.json");

    private sealed record Prog(int Done, int Total, string? Line);

    public MainForm()
    {
        Text = "MultiReplace - 一括置換ツール";
        Font = new Font("Yu Gothic UI", 9F);
        ClientSize = new Size(860, 760);
        MinimumSize = new Size(700, 620);
        StartPosition = FormStartPosition.CenterScreen;
        AllowDrop = true;
        DragEnter += (_, e) => { if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy; };
        DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is string[] a && a.Length > 0)
            {
                var p = a[0];
                if (File.Exists(p)) p = Path.GetDirectoryName(p) ?? p;
                if (Directory.Exists(p)) txtDir.Text = p;
            }
        };

        BuildLayout();
        WireEvents();
        LoadSettings();
        UpdateEnabled();
        FormClosing += (_, _) => SaveSettings();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(10) };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // フォルダ
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // ファイル名
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // 除外
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // モード/Excel
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 38));    // 置換ワード
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // オプション
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // ボタン
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 62));    // ログ

        // 1行目: フォルダ
        var r1 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true };
        r1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        r1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        r1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        r1.Controls.Add(new Label { Text = "対象フォルダ", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, 0);
        r1.Controls.Add(txtDir, 1, 0);
        r1.Controls.Add(btnBrowse, 2, 0);
        root.Controls.Add(r1);

        // 2行目: ファイル名パターン
        var r2 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, AutoSize = true };
        r2.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        r2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        r2.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        r2.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        r2.Controls.Add(new Label { Text = "対象ファイル名", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, 0);
        r2.Controls.Add(txtPattern, 1, 0);
        r2.Controls.Add(new Label { Text = "例: *.txt;*.csv;data_*.xlsx（; 区切り）", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = SystemColors.GrayText, Margin = new Padding(8, 6, 8, 6) }, 2, 0);
        r2.Controls.Add(chkSub, 3, 0);
        root.Controls.Add(r2);

        // 2.5行目: 除外フィルタ
        var r2b = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true };
        r2b.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        r2b.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        r2b.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        r2b.Controls.Add(new Label { Text = "除外ファイル名", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 6) }, 0, 0);
        r2b.Controls.Add(txtExclude, 1, 0);
        r2b.Controls.Add(new Label { Text = @"例: *.bak;*_old.*;.git\ （末尾が \ ならフォルダ名の除外）", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = SystemColors.GrayText, Margin = new Padding(8, 6, 0, 6) }, 2, 0);
        root.Controls.Add(r2b);

        // 3行目: モード + Excel対象
        var r3 = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
        r3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        r3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        var fm = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
        fm.Controls.AddRange(new Control[] { chkContent, chkName, chkDir, chkNameExt });
        gbMode.Controls.Add(fm);
        var fx = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
        fx.Controls.AddRange(new Control[] { chkXlCells, chkXlNum, chkXlSheet, chkXlObj });
        gbExcel.Controls.Add(fx);
        r3.Controls.Add(gbMode, 0, 0);
        r3.Controls.Add(gbExcel, 1, 0);
        root.Controls.Add(r3);

        // 4行目: 置換ワード表
        var gbRules = new GroupBox { Text = "置換ワード（複数登録可。Excelなどからコピーして Ctrl+V で貼り付けもできます）", Dock = DockStyle.Fill, Padding = new Padding(8) };
        grid.Dock = DockStyle.Fill;
        grid.AllowUserToAddRows = true;
        grid.AllowUserToDeleteRows = true;
        grid.AllowUserToResizeRows = false;
        grid.BackgroundColor = Color.White;
        grid.RowHeadersWidth = 28;
        grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "検索する文字列", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "置換後の文字列（空欄なら削除）", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable });
        gbRules.Controls.Add(grid);
        root.Controls.Add(gbRules);

        // 5行目: オプション
        var fo = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
        fo.Controls.AddRange(new Control[] { chkCase, chkRegex, chkEscape, chkBackup });
        root.Controls.Add(fo);

        // 6行目: ボタン
        var fb = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        btnRun.Font = new Font(Font, FontStyle.Bold);
        fb.Controls.AddRange(new Control[] { btnPreview, btnRun, btnCancel, bar });
        bar.Margin = new Padding(12, 6, 3, 3);
        root.Controls.Add(fb);

        // 7行目: ログ
        root.Controls.Add(txtLog);

        Controls.Add(root);
    }

    private void WireEvents()
    {
        btnBrowse.Click += (_, _) =>
        {
            using var d = new FolderBrowserDialog { Description = "置換対象のフォルダを選択", UseDescriptionForTitle = true };
            if (Directory.Exists(txtDir.Text)) d.InitialDirectory = txtDir.Text;
            if (d.ShowDialog(this) == DialogResult.OK) txtDir.Text = d.SelectedPath;
        };
        chkContent.CheckedChanged += (_, _) => UpdateEnabled();
        chkName.CheckedChanged += (_, _) => UpdateEnabled();
        chkDir.CheckedChanged += (_, _) => UpdateEnabled();
        btnPreview.Click += async (_, _) => await RunAsync(false);
        btnRun.Click += async (_, _) => await RunAsync(true);
        btnCancel.Click += (_, _) => _cts?.Cancel();
        grid.KeyDown += Grid_KeyDown;
    }

    private bool DoContent => chkContent.Checked;
    private bool DoName => chkName.Checked;
    private bool DoDir => chkDir.Checked;

    private void UpdateEnabled()
    {
        gbExcel.Enabled = DoContent;
        chkBackup.Enabled = DoContent;
        chkNameExt.Enabled = DoName;
    }

    private void Grid_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.V && !grid.IsCurrentCellInEditMode)
        {
            var text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;
            e.Handled = true;
            foreach (var line in text.Replace("\r\n", "\n").TrimEnd('\n').Split('\n'))
            {
                var cols = line.Split('\t');
                grid.Rows.Add(cols[0], cols.Length > 1 ? cols[1] : "");
            }
        }
    }

    private List<Rule> ReadRules()
    {
        grid.EndEdit();
        var list = new List<Rule>();
        foreach (DataGridViewRow r in grid.Rows)
        {
            if (r.IsNewRow) continue;
            var f = r.Cells[0].Value?.ToString() ?? "";
            var t = r.Cells[1].Value?.ToString() ?? "";
            if (f.Length == 0) continue;
            // 文字列モード: 検索・置換とも \n \t を変換 / 正規表現モード: 検索は正規表現側で解釈するので置換後だけ変換
            if (chkEscape.Checked) { if (!chkRegex.Checked) f = Replacer.Unescape(f); t = Replacer.Unescape(t); }
            list.Add(new Rule(f, t));
        }
        return list;
    }

    private void LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return;
            var s = JsonSerializer.Deserialize<Settings>(File.ReadAllText(SettingsPath));
            if (s == null) return;
            txtDir.Text = s.Dir; txtPattern.Text = s.Pattern; txtExclude.Text = s.Exclude;
            chkSub.Checked = s.Sub; chkCase.Checked = s.CaseSensitive; chkEscape.Checked = s.Escape; chkRegex.Checked = s.Regex; chkBackup.Checked = s.Backup;
            chkContent.Checked = s.DoContent; chkName.Checked = s.DoName; chkDir.Checked = s.DoDir;
            chkNameExt.Checked = s.NameWithExt;
            chkXlCells.Checked = s.XlCells; chkXlSheet.Checked = s.XlSheetNames; chkXlObj.Checked = s.XlObjects; chkXlNum.Checked = s.XlNumbers;
            foreach (var r in s.Rules) grid.Rows.Add(r.Find, r.Replace);
        }
        catch { /* 設定が壊れていても起動は継続 */ }
    }

    private void SaveSettings()
    {
        try
        {
            grid.EndEdit();
            var s = new Settings
            {
                Dir = txtDir.Text, Pattern = txtPattern.Text, Exclude = txtExclude.Text, Sub = chkSub.Checked,
                CaseSensitive = chkCase.Checked, Escape = chkEscape.Checked, Regex = chkRegex.Checked, Backup = chkBackup.Checked,
                DoContent = chkContent.Checked, DoName = chkName.Checked, DoDir = chkDir.Checked,
                NameWithExt = chkNameExt.Checked,
                XlCells = chkXlCells.Checked, XlSheetNames = chkXlSheet.Checked, XlObjects = chkXlObj.Checked, XlNumbers = chkXlNum.Checked,
            };
            // エスケープ変換前の生の入力を保存する
            foreach (DataGridViewRow r in grid.Rows)
            {
                if (r.IsNewRow) continue;
                var f = r.Cells[0].Value?.ToString() ?? "";
                if (f.Length == 0) continue;
                s.Rules.Add(new Rule(f, r.Cells[1].Value?.ToString() ?? ""));
            }
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(s, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private void Log(string line) => txtLog.AppendText(line + Environment.NewLine);

    private void SetBusy(bool busy)
    {
        btnPreview.Enabled = btnRun.Enabled = !busy;
        btnCancel.Enabled = busy;
        foreach (Control c in new Control[] { txtDir, txtPattern, txtExclude, btnBrowse, grid }) c.Enabled = !busy;
    }

    private async Task RunAsync(bool apply)
    {
        var dir = txtDir.Text.Trim().Trim('"');
        if (!Directory.Exists(dir)) { MessageBox.Show(this, "対象フォルダが見つかりません。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        var rules = ReadRules();
        if (rules.Count == 0) { MessageBox.Show(this, "置換ワードを1つ以上入力してください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        var patterns = txtPattern.Text.Split(new[] { ';', ',', '|' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var excludes = txtExclude.Text.Split(new[] { ';', ',', '|' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        Replacer rp;
        try { rp = new Replacer(rules, chkCase.Checked, chkRegex.Checked); }
        catch (ArgumentException ex) { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        if (!DoContent && !DoName && !DoDir)
        {
            MessageBox.Show(this, "「置換するもの」を1つ以上選んでください。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var opt = new EngineOptions(apply, chkBackup.Checked, DoContent, DoName, chkNameExt.Checked,
            new ExcelOptions(chkXlCells.Checked, chkXlSheet.Checked, chkXlObj.Checked, chkXlNum.Checked), DoDir);

        if (apply)
        {
            var what = string.Join("・", new[] { DoContent ? "ファイルの中身" : null, DoName ? "ファイル名" : null, DoDir ? "ディレクトリ名" : null }.Where(x => x != null));
            var bk = DoContent ? (opt.Backup ? "\nバックアップ(.bak)を作成します。" : "\n※バックアップは作成しません。") : "";
            var ok = MessageBox.Show(this, $"{dir}\n内の{what}を書き換えます。{bk}\n\n実行しますか？（先にプレビューで確認するのがおすすめです）",
                Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (ok != DialogResult.OK) return;
        }

        txtLog.Clear();
        Log($"=== {(apply ? "置換実行" : "プレビュー（ファイルは変更しません）")}  {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
        Log($"フォルダ: {dir}  /  ファイル名: {string.Join(" ; ", patterns)}  /  除外: {(excludes.Length == 0 ? "なし" : string.Join(" ; ", excludes))}  /  サブフォルダ: {(chkSub.Checked ? "含む" : "含まない")}");
        Log($"  モード: {(chkRegex.Checked ? "正規表現" : "文字列")} / 大文字小文字: {(chkCase.Checked ? "区別する" : "区別しない")}");
        foreach (var r in rules) Log($"  置換: \"{r.Find}\" → \"{r.Replace}\"");

        SetBusy(true);
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        var progress = new Progress<Prog>(p =>
        {
            if (p.Total > 0) { bar.Maximum = p.Total; bar.Value = Math.Min(p.Done, p.Total); }
            if (p.Line != null) Log(p.Line);
        });
        bool sub = chkSub.Checked;

        try
        {
            var sum = await Task.Run(() => Worker(dir, patterns, excludes, sub, rp, opt, progress, ct));
            Log("");
            Log($"対象 {sum.Files} ファイル / {(apply ? "変更" : "変更予定")} {sum.Changed} ファイル"
                + (DoContent ? $" / 置換 {sum.Replacements} 箇所" : "")
                + (DoName ? $" / ファイル名{(apply ? "変更" : "変更予定")} {sum.Renamed} 件" : "")
                + (DoDir ? $" / フォルダ名{(apply ? "変更" : "変更予定")} {sum.DirsRenamed} 件" : "")
                + $" / スキップ {sum.Skipped} / エラー {sum.Errors}"
                + (sum.Cancelled ? " （中止しました）" : ""));
        }
        catch (Exception ex)
        {
            Log("エラー: " + ex.Message);
        }
        finally
        {
            SetBusy(false);
            _cts.Dispose();
            _cts = null;
        }
    }

    private sealed record Summary(int Files, int Changed, int Replacements, int Renamed, int Skipped, int Errors, bool Cancelled, int DirsRenamed = 0);

    private static Summary Worker(string dir, string[] patterns, string[] excludes, bool sub, Replacer rp, EngineOptions o,
        IProgress<Prog> progress, CancellationToken ct)
    {
        var files = o.ReplaceContent || o.ReplaceName ? Engine.Enumerate(dir, patterns, sub, excludes) : new List<string>();
        progress.Report(new Prog(0, Math.Max(files.Count, 1), $"{files.Count} ファイルが見つかりました。"));
        var planned = new HashSet<string>();
        int changed = 0, repl = 0, renamed = 0, skipped = 0, errors = 0, done = 0;
        bool cancelled = false;
        string verb = o.Apply ? "" : "予定";

        var gate = new object();

        void One(string f)
        {
            var r = Engine.ProcessFile(f, rp, o, planned);
            string rel = Path.GetRelativePath(dir, f);
            string? line = null;
            int d;
            lock (gate)
            {
                d = ++done;
                switch (r.Status)
                {
                    case Status.Changed:
                        changed++;
                        repl += r.Count;
                        var parts = new List<string>();
                        if (o.ReplaceContent && r.Count > 0) parts.Add($"{r.Count}箇所");
                        if (r.RenamedTo != null) { renamed++; parts.Add($"名前: → {Path.GetFileName(r.RenamedTo)}"); }
                        line = $"[置換{verb}] {rel}  ({string.Join(" / ", parts)})";
                        if (r.Message != null) line += $"  ※{r.Message}";
                        break;
                    case Status.Skipped:
                        skipped++;
                        line = $"[スキップ] {rel}  {r.Message}";
                        break;
                    case Status.Error:
                        errors++;
                        line = $"[エラー] {rel}  {r.Message}";
                        break;
                }
            }
            progress.Report(new Prog(d, Math.Max(files.Count, 1), line));
        }

        if (o.ReplaceName)
        {
            // ファイル名の衝突判定が順序に依存するため、名前変更を伴うときは順番に処理する
            foreach (var f in files)
            {
                if (ct.IsCancellationRequested) { cancelled = true; break; }
                One(f);
            }
        }
        else
        {
            // 中身の置換だけなら、ファイル間に依存関係がないので並列で処理する
            var po = new ParallelOptions { CancellationToken = ct, MaxDegreeOfParallelism = Math.Max(1, Environment.ProcessorCount - 1) };
            try { Parallel.ForEach(files, po, One); }
            catch (OperationCanceledException) { cancelled = true; }
        }
        // ディレクトリ名は、ファイルの処理が終わってから深い階層順に変更する
        int dirsRenamed = 0;
        if (o.ReplaceDir && !cancelled)
        {
            progress.Report(new Prog(0, 0, "--- ディレクトリ名（深い階層から順に処理。表示は変更前のパス）---"));
            foreach (var r in Engine.RenameDirectories(dir, rp, o, sub, excludes, planned, ct))
            {
                string rel = Path.GetRelativePath(dir, r.Path);
                switch (r.Status)
                {
                    case Status.Changed:
                        dirsRenamed++; changed++;
                        progress.Report(new Prog(0, 0, $"[フォルダ名置換{verb}] {rel}  → {Path.GetFileName(r.RenamedTo)}"));
                        break;
                    case Status.Skipped:
                        skipped++;
                        progress.Report(new Prog(0, 0, $"[スキップ] {rel}  {r.Message}"));
                        break;
                    case Status.Error:
                        errors++;
                        progress.Report(new Prog(0, 0, $"[エラー] {rel}  {r.Message}"));
                        break;
                }
            }
            if (ct.IsCancellationRequested) cancelled = true;
        }
        return new Summary(files.Count, changed, repl, renamed, skipped, errors, cancelled, dirsRenamed);
    }
}
