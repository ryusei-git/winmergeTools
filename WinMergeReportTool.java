import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.charset.Charset;
import java.nio.charset.IllegalCharsetNameException;
import java.nio.charset.StandardCharsets;
import java.nio.charset.UnsupportedCharsetException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * WinMerge を使ったテストケース比較レポート自動生成ツール（単一ファイル版）。
 *
 * <p>処理の流れ:
 * <ol>
 *   <li>カレントディレクトリ（--root で変更可）配下のディレクトリ階層を report フォルダへ複製する</li>
 *   <li>テストケースディレクトリ（INPUT/OUTPUT/LOG/LOG_COMPARE を持つディレクトリ）を検出する</li>
 *   <li>INPUT ⇔ OUTPUT、LOG ⇔ LOG_COMPARE をフォルダ比較（/r）し HTML レポートを出力する</li>
 *   <li>同名ファイル同士をファイル比較し、HTML レポートを report 配下の対応ディレクトリへ出力する</li>
 *   <li>テストケース名をシート名にした Excel ブックを作成し、HTML レポートを各シートへ貼り付ける</li>
 * </ol>
 *
 * <p>実行例（Windows）:
 * <pre>
 *   java WinMergeReportTool.java
 *   java WinMergeReportTool.java --winmerge "C:\Program Files\WinMerge\WinMergeU.exe"
 *   java WinMergeReportTool.java --root D:\tests --report D:\tests\report --excel D:\tests\report\result.xlsx
 *   java WinMergeReportTool.java --pair INPUT:OUTPUT --pair LOG:LOG_COMPARE
 * </pre>
 *
 * <p>Excel 出力は 2 系統ある。
 * <ul>
 *   <li>com  : Excel の COM 自動化（VBScript 経由）で HTML レポートを実際にシートへ貼り付ける（Windows + Excel 必須）</li>
 *   <li>xlsx : 外部ライブラリなしで .xlsx を直接生成し、HTML の表を解析してセルに展開＋元 HTML へのリンクを張る</li>
 * </ul>
 * 既定は auto（Windows + cscript があれば com、失敗したら xlsx へフォールバック）。
 */
public final class WinMergeReportTool {

    /**
     * 外部プロセス（WinMerge / cscript）の出力を読むための文字コード。
     * JDK 18 以降は file.encoding が UTF-8 固定になるため、OS 本来の文字コード
     * （日本語 Windows なら windows-31j）を native.encoding から取得する。
     */
    private static final Charset NATIVE_CHARSET = nativeCharset();

    private static Charset nativeCharset() {
        String name = System.getProperty("native.encoding");
        if (name != null && !name.isEmpty()) {
            try {
                return Charset.forName(name);
            } catch (IllegalCharsetNameException | UnsupportedCharsetException ignore) {
                // 既定へフォールバックする
            }
        }
        return Charset.defaultCharset();
    }

    private static final DateTimeFormatter TS_HUMAN = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
    private static final DateTimeFormatter TS_FILE = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    // ======================================================================
    // 設定
    // ======================================================================

    /** 比較する 2 つのサブディレクトリの組。 */
    static final class ComparePair {
        final String left;
        final String right;
        final String label;

        ComparePair(String left, String right) {
            this.left = left;
            this.right = right;
            this.label = left + "_vs_" + right;
        }
    }

    static final class Config {
        Path root = Paths.get("").toAbsolutePath().normalize();
        Path reportRoot;                 // 既定: root/report
        Path excelPath;                  // 既定: reportRoot/comparison_report_<timestamp>.xlsx
        String winMerge;                 // WinMergeU.exe のパス
        final List<ComparePair> pairs = new ArrayList<>();
        final List<String> extraArgs = new ArrayList<>();
        long timeoutSec = 180;
        String excelMode = "auto";       // auto | com | xlsx | none
        int maxTableRows = 500;          // xlsx へ展開する 1 レポートあたりの最大行数
        boolean cleanReport = false;     // 実行前に report フォルダを削除するか
        boolean enableExitCode = true;   // /enableexitcode を使うか（WinMerge 2.16 以降）
        int maxPathLength = 240;         // Windows の MAX_PATH(260) を考慮した上限
    }

    // ======================================================================
    // 比較結果モデル
    // ======================================================================

    enum Status {
        SAME("一致"),
        DIFF("差分あり"),
        LEFT_ONLY("左のみ"),
        RIGHT_ONLY("右のみ"),
        ERROR("エラー"),
        SKIPPED("スキップ");

        final String label;

        Status(String label) {
            this.label = label;
        }
    }

    static final class CompareResult {
        String kind;        // "フォルダ比較" / "ファイル比較"
        String pairLabel;   // INPUT_vs_OUTPUT など
        String name;        // 比較対象名（ファイルの相対パス、フォルダ比較は "(フォルダ全体)"）
        Path left;
        Path right;
        Path report;        // 生成された HTML レポート（無い場合 null）
        Status status = Status.SKIPPED;
        String message = "";
        int exitCode = -1;
    }

    static final class TestCase {
        String name;                 // シート名の元になるテストケース名
        Path dir;                    // 実ディレクトリ
        Path rel;                    // root からの相対パス
        final List<CompareResult> results = new ArrayList<>();
    }

    /** コマンドライン引数の誤りを表す例外（使い方を表示して終了する）。 */
    static final class UsageException extends RuntimeException {
        UsageException(String message) {
            super(message);
        }
    }

    private final Config cfg;
    private final List<TestCase> testCases = new ArrayList<>();

    private WinMergeReportTool(Config cfg) {
        this.cfg = cfg;
    }

    // ======================================================================
    // エントリポイント
    // ======================================================================

    public static void main(String[] args) {
        try {
            Config cfg = parseArgs(args);
            if (cfg == null) {
                return; // --help
            }
            int code = new WinMergeReportTool(cfg).run();
            System.exit(code);
        } catch (UsageException e) {
            System.err.println("[ERROR] " + e.getMessage());
            System.err.println();
            printUsage();
            System.exit(2);
        } catch (Exception e) {
            System.err.println("[FATAL] 予期しないエラーが発生しました: " + e);
            e.printStackTrace();
            System.exit(1);
        }
    }

    private static void printUsage() {
        System.out.println(String.join(System.lineSeparator(),
                "使い方: java WinMergeReportTool.java [オプション]",
                "",
                "  --root <dir>        比較対象のルート（既定: カレントディレクトリ）",
                "  --report <dir>      レポート出力先（既定: <root>/report）",
                "  --excel <file>      Excel 出力先（既定: <report>/comparison_report_<日時>.xlsx）",
                "  --winmerge <exe>    WinMergeU.exe のパス（既定: 環境変数 WINMERGE_PATH → 既知の場所を自動探索）",
                "  --pair <L>:<R>      比較するサブディレクトリの組（複数指定可）",
                "                      既定: --pair INPUT:OUTPUT --pair LOG:LOG_COMPARE",
                "  --excel-mode <mode> auto | com | xlsx | none（既定: auto）",
                "  --timeout <sec>     WinMerge 1 回あたりのタイムアウト秒（既定: 180）",
                "  --max-rows <n>      xlsx へ展開する 1 レポートあたりの最大行数（既定: 500）",
                "  --clean             実行前に report フォルダを削除する",
                "  --no-exitcode       /enableexitcode を使わず、ファイル内容の一致判定を Java 側で行う",
                "                      （WinMerge 2.14 以前で /enableexitcode が使えない場合）",
                "  --max-path <n>      レポートパスの上限文字数（既定: 240、超える場合は短縮名にする）",
                "  --winmerge-arg <a>  WinMerge へ渡す追加引数（複数指定可）",
                "  --help              このヘルプ"));
    }

    private static Config parseArgs(String[] args) {
        Config cfg = new Config();
        List<ComparePair> pairs = new ArrayList<>();
        for (int i = 0; i < args.length; i++) {
            String a = args[i];
            switch (a) {
                case "--help":
                case "-h":
                    printUsage();
                    return null;
                case "--root":
                    cfg.root = Paths.get(next(args, ++i, a)).toAbsolutePath().normalize();
                    break;
                case "--report":
                    cfg.reportRoot = Paths.get(next(args, ++i, a)).toAbsolutePath().normalize();
                    break;
                case "--excel":
                    cfg.excelPath = Paths.get(next(args, ++i, a)).toAbsolutePath().normalize();
                    break;
                case "--winmerge":
                    cfg.winMerge = next(args, ++i, a);
                    break;
                case "--pair": {
                    String v = next(args, ++i, a);
                    int sep = v.indexOf(':');
                    if (sep <= 0 || sep == v.length() - 1) {
                        throw new UsageException("--pair は LEFT:RIGHT の形式で指定してください: " + v);
                    }
                    pairs.add(new ComparePair(v.substring(0, sep), v.substring(sep + 1)));
                    break;
                }
                case "--excel-mode":
                    cfg.excelMode = next(args, ++i, a).toLowerCase(Locale.ROOT);
                    if (!Arrays.asList("auto", "com", "xlsx", "none").contains(cfg.excelMode)) {
                        throw new UsageException("--excel-mode は auto|com|xlsx|none のいずれか: " + cfg.excelMode);
                    }
                    break;
                case "--timeout":
                    cfg.timeoutSec = parseLong(next(args, ++i, a), a);
                    break;
                case "--max-rows":
                    cfg.maxTableRows = (int) parseLong(next(args, ++i, a), a);
                    break;
                case "--clean":
                    cfg.cleanReport = true;
                    break;
                case "--no-exitcode":
                    cfg.enableExitCode = false;
                    break;
                case "--max-path":
                    cfg.maxPathLength = (int) parseLong(next(args, ++i, a), a);
                    break;
                case "--winmerge-arg":
                    cfg.extraArgs.add(next(args, ++i, a));
                    break;
                default:
                    throw new UsageException("不明なオプション: " + a);
            }
        }
        if (pairs.isEmpty()) {
            pairs.add(new ComparePair("INPUT", "OUTPUT"));
            pairs.add(new ComparePair("LOG", "LOG_COMPARE"));
        }
        cfg.pairs.addAll(pairs);
        if (cfg.reportRoot == null) {
            cfg.reportRoot = cfg.root.resolve("report");
        }
        if (cfg.excelPath == null) {
            cfg.excelPath = cfg.reportRoot.resolve(
                    "comparison_report_" + LocalDateTime.now().format(TS_FILE) + ".xlsx");
        }
        return cfg;
    }

    private static String next(String[] args, int i, String opt) {
        if (i >= args.length) {
            throw new UsageException(opt + " には値が必要です");
        }
        return args[i];
    }

    private static long parseLong(String value, String opt) {
        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException e) {
            throw new UsageException(opt + " には数値を指定してください: " + value);
        }
    }

    // ======================================================================
    // メイン処理
    // ======================================================================

    private int run() throws IOException, InterruptedException {
        System.out.println("==================================================");
        System.out.println(" WinMerge テストケース比較レポート生成");
        System.out.println("==================================================");
        System.out.println("  ルート        : " + cfg.root);
        System.out.println("  レポート出力先: " + cfg.reportRoot);
        System.out.println("  Excel 出力先  : " + cfg.excelPath);
        System.out.println("  比較対象      : " + cfg.pairs.stream().map(p -> p.label).collect(Collectors.joining(", ")));

        if (!Files.isDirectory(cfg.root)) {
            System.err.println("[ERROR] ルートディレクトリが存在しません: " + cfg.root);
            return 1;
        }

        cfg.winMerge = resolveWinMerge(cfg.winMerge);
        System.out.println("  WinMerge      : " + (cfg.winMerge == null ? "(見つかりません)" : cfg.winMerge));
        System.out.println();
        if (cfg.winMerge == null) {
            System.err.println("[ERROR] WinMergeU.exe が見つかりません。--winmerge で明示指定するか、環境変数 WINMERGE_PATH を設定してください。");
            return 1;
        }

        if (cfg.cleanReport && Files.exists(cfg.reportRoot)) {
            System.out.println("[INFO] 既存の report フォルダを削除します: " + cfg.reportRoot);
            deleteRecursively(cfg.reportRoot);
        }
        Files.createDirectories(cfg.reportRoot);
        if (isWindows() && cfg.reportRoot.toString().length() > 150) {
            System.err.println("[WARN] レポート出力先のパスが長いため、MAX_PATH(260) を超える恐れがあります: "
                    + cfg.reportRoot.toString().length() + " 文字");
            System.err.println("       --report でより浅い場所を指定するか、--max-path で短縮のしきい値を調整してください。");
        }

        // 1) ディレクトリ階層を report 配下へ複製
        int mirrored = mirrorDirectoryTree();
        System.out.println("[INFO] ディレクトリ階層を複製しました（" + mirrored + " 個）");

        // 2) テストケースの検出
        collectTestCases();
        if (testCases.isEmpty()) {
            System.err.println("[WARN] テストケースディレクトリが見つかりませんでした（"
                    + cfg.pairs.stream().map(p -> p.left + "/" + p.right).collect(Collectors.joining(", "))
                    + " を含むディレクトリを探索しました）");
            return 1;
        }
        System.out.println("[INFO] テストケース " + testCases.size() + " 件を検出しました");
        System.out.println();

        // 3) 比較の実行
        for (TestCase tc : testCases) {
            System.out.println("---- テストケース: " + tc.name + " (" + tc.rel + ")");
            processTestCase(tc);
        }

        // 4) 一覧 HTML
        Path index = cfg.reportRoot.resolve("index.html");
        writeIndexHtml(index);
        System.out.println();
        System.out.println("[INFO] 一覧レポート: " + index);

        // 5) Excel 出力
        writeExcel();

        printSummary();
        return hasError() ? 1 : 0;
    }

    private boolean hasError() {
        return testCases.stream().flatMap(t -> t.results.stream()).anyMatch(r -> r.status == Status.ERROR);
    }

    private void printSummary() {
        int same = 0, diff = 0, err = 0, only = 0;
        for (TestCase tc : testCases) {
            for (CompareResult r : tc.results) {
                switch (r.status) {
                    case SAME: same++; break;
                    case DIFF: diff++; break;
                    case ERROR: err++; break;
                    case LEFT_ONLY:
                    case RIGHT_ONLY: only++; break;
                    default: break;
                }
            }
        }
        System.out.println();
        System.out.println("==================== 集計 ====================");
        System.out.println("  テストケース : " + testCases.size());
        System.out.println("  一致         : " + same);
        System.out.println("  差分あり     : " + diff);
        System.out.println("  片側のみ     : " + only);
        System.out.println("  エラー       : " + err);
        System.out.println("==============================================");
    }

    // ----------------------------------------------------------------------
    // WinMerge の場所解決
    // ----------------------------------------------------------------------

    private static String resolveWinMerge(String specified) {
        List<String> candidates = new ArrayList<>();
        if (specified != null && !specified.isEmpty()) {
            candidates.add(specified);
        }
        String env = System.getenv("WINMERGE_PATH");
        if (env != null && !env.isEmpty()) {
            candidates.add(env);
        }
        String prop = System.getProperty("winmerge.path");
        if (prop != null && !prop.isEmpty()) {
            candidates.add(prop);
        }
        candidates.add("C:\\Program Files\\WinMerge\\WinMergeU.exe");
        candidates.add("C:\\Program Files (x86)\\WinMerge\\WinMergeU.exe");
        String localAppData = System.getenv("LOCALAPPDATA");
        if (localAppData != null) {
            candidates.add(localAppData + "\\Programs\\WinMerge\\WinMergeU.exe");
        }
        for (String c : candidates) {
            Path p = Paths.get(c);
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().toString();
            }
        }
        // PATH 上に存在するかどうか（明示指定されたコマンド名の場合）
        for (String c : candidates) {
            if (!c.contains("\\") && !c.contains("/") && isOnPath(c)) {
                return c;
            }
        }
        return null;
    }

    private static boolean isOnPath(String cmd) {
        String path = System.getenv("PATH");
        if (path == null) {
            return false;
        }
        for (String dir : path.split(Pattern.quote(java.io.File.pathSeparator))) {
            if (dir.isEmpty()) {
                continue;
            }
            Path p = Paths.get(dir).resolve(cmd);
            if (Files.isRegularFile(p)) {
                return true;
            }
            if (Files.isRegularFile(Paths.get(dir).resolve(cmd + ".exe"))) {
                return true;
            }
        }
        return false;
    }

    // ----------------------------------------------------------------------
    // 1) ディレクトリ階層の複製
    // ----------------------------------------------------------------------

    private int mirrorDirectoryTree() throws IOException {
        final int[] count = {0};
        Files.walkFileTree(cfg.root, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                if (dir.equals(cfg.root)) {
                    return FileVisitResult.CONTINUE;
                }
                if (isExcludedDir(dir)) {
                    return FileVisitResult.SKIP_SUBTREE;
                }
                Path target = cfg.reportRoot.resolve(cfg.root.relativize(dir).toString());
                if (!Files.exists(target)) {
                    Files.createDirectories(target);
                    count[0]++;
                }
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFileFailed(Path file, IOException exc) {
                System.err.println("[WARN] 走査に失敗しました: " + file + " (" + exc.getMessage() + ")");
                return FileVisitResult.CONTINUE;
            }
        });
        return count[0];
    }

    /** report 自体・VCS 管理ディレクトリ・隠しディレクトリは除外する。 */
    private boolean isExcludedDir(Path dir) {
        try {
            if (Files.isSameFile(dir, cfg.reportRoot)) {
                return true;
            }
        } catch (IOException ignore) {
            // 存在しないなど。名前で判定にフォールバックする。
        }
        if (dir.startsWith(cfg.reportRoot)) {
            return true;
        }
        String name = dir.getFileName() == null ? "" : dir.getFileName().toString();
        return name.equals(".git") || name.equals(".svn") || name.equals(".hg") || name.startsWith("$");
    }

    // ----------------------------------------------------------------------
    // 2) テストケース検出
    // ----------------------------------------------------------------------

    private void collectTestCases() throws IOException {
        List<Path> dirs = new ArrayList<>();
        Files.walkFileTree(cfg.root, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) {
                if (!dir.equals(cfg.root) && isExcludedDir(dir)) {
                    return FileVisitResult.SKIP_SUBTREE;
                }
                if (isTestCaseDir(dir)) {
                    dirs.add(dir);
                    return FileVisitResult.SKIP_SUBTREE; // テストケース配下は再探索しない
                }
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFileFailed(Path file, IOException exc) {
                return FileVisitResult.CONTINUE;
            }
        });
        dirs.sort(Comparator.comparing(Path::toString, String.CASE_INSENSITIVE_ORDER));
        for (Path d : dirs) {
            TestCase tc = new TestCase();
            tc.dir = d;
            tc.rel = cfg.root.relativize(d);
            tc.name = d.getFileName().toString();
            testCases.add(tc);
        }
    }

    /** 設定された比較ペアのいずれかのサブディレクトリを持つならテストケースとみなす。 */
    private boolean isTestCaseDir(Path dir) {
        if (dir.equals(cfg.root)) {
            return false;
        }
        for (ComparePair p : cfg.pairs) {
            if (Files.isDirectory(dir.resolve(p.left)) || Files.isDirectory(dir.resolve(p.right))) {
                return true;
            }
        }
        return false;
    }

    // ----------------------------------------------------------------------
    // 3) 比較の実行
    // ----------------------------------------------------------------------

    private void processTestCase(TestCase tc) throws IOException, InterruptedException {
        for (ComparePair pair : cfg.pairs) {
            Path leftDir = tc.dir.resolve(pair.left);
            Path rightDir = tc.dir.resolve(pair.right);
            Path outDir = cfg.reportRoot.resolve(tc.rel.toString()).resolve(pair.label);

            if (!Files.isDirectory(leftDir) && !Files.isDirectory(rightDir)) {
                continue; // このペアは対象外
            }
            Files.createDirectories(outDir);

            if (!Files.isDirectory(leftDir) || !Files.isDirectory(rightDir)) {
                CompareResult r = new CompareResult();
                r.kind = "フォルダ比較";
                r.pairLabel = pair.label;
                r.name = "(フォルダ全体)";
                r.left = leftDir;
                r.right = rightDir;
                r.status = Status.ERROR;
                r.message = "比較対象のディレクトリが存在しません: "
                        + (Files.isDirectory(leftDir) ? rightDir : leftDir);
                tc.results.add(r);
                System.out.println("     [NG] " + pair.label + " : " + r.message);
                continue;
            }

            // --- フォルダ比較 ---
            tc.results.add(runFolderCompare(pair, leftDir, rightDir, outDir));

            // --- ファイル比較 ---
            Path fileOutDir = outDir.resolve("files");
            for (String rel : listRelativeFiles(leftDir, rightDir)) {
                tc.results.add(runFileCompare(pair, leftDir, rightDir, fileOutDir, rel));
            }
        }
    }

    private CompareResult runFolderCompare(ComparePair pair, Path leftDir, Path rightDir, Path outDir)
            throws IOException, InterruptedException {
        CompareResult r = new CompareResult();
        r.kind = "フォルダ比較";
        r.pairLabel = pair.label;
        r.name = "(フォルダ全体)";
        r.left = leftDir;
        r.right = rightDir;
        r.report = outDir.resolve("_folder_compare.html");

        Files.createDirectories(r.report.getParent());
        List<String> cmd = buildWinMergeCommand(true, leftDir, rightDir, r.report);
        ExecResult ex = exec(cmd);
        applyExitCode(r, ex);
        System.out.println("     [" + statusMark(r.status) + "] " + pair.label + " フォルダ比較 -> "
                + cfg.reportRoot.relativize(r.report));
        return r;
    }

    private CompareResult runFileCompare(ComparePair pair, Path leftDir, Path rightDir, Path fileOutDir, String rel)
            throws IOException, InterruptedException {
        CompareResult r = new CompareResult();
        r.kind = "ファイル比較";
        r.pairLabel = pair.label;
        r.name = rel;
        r.left = leftDir.resolve(rel);
        r.right = rightDir.resolve(rel);

        boolean hasLeft = Files.isRegularFile(r.left);
        boolean hasRight = Files.isRegularFile(r.right);
        if (!hasLeft || !hasRight) {
            r.status = hasLeft ? Status.LEFT_ONLY : Status.RIGHT_ONLY;
            r.message = hasLeft
                    ? pair.right + " 側にファイルがありません"
                    : pair.left + " 側にファイルがありません";
            System.out.println("     [--] " + pair.label + " " + rel + " : " + r.message);
            return r;
        }

        r.report = shortenIfTooLong(fileOutDir, rel);
        if (!r.report.equals(fileOutDir.resolve(rel + ".html"))) {
            r.message = "パスが長いためレポート名を短縮しました";
        }
        Files.createDirectories(r.report.getParent());
        List<String> cmd = buildWinMergeCommand(false, r.left, r.right, r.report);
        ExecResult ex = exec(cmd);
        applyExitCode(r, ex);
        System.out.println("     [" + statusMark(r.status) + "] " + pair.label + " " + rel + " -> "
                + cfg.reportRoot.relativize(r.report));
        return r;
    }

    /**
     * Windows の MAX_PATH(260) 制限を避けるため、レポートの出力パスが長すぎる場合に短縮名へ切り替える。
     * まずサブフォルダを潰した平坦な名前を試し、それでも長い場合はハッシュ付きの短い名前にする。
     */
    private Path shortenIfTooLong(Path fileOutDir, String rel) {
        Path desired = fileOutDir.resolve(rel + ".html");
        if (desired.toString().length() <= cfg.maxPathLength) {
            return desired;
        }
        Path flat = fileOutDir.resolve(rel.replace('/', '_') + ".html");
        if (flat.toString().length() <= cfg.maxPathLength) {
            return flat;
        }
        String hash = String.format("%08x", rel.hashCode());
        String leaf = rel.substring(rel.lastIndexOf('/') + 1);
        int room = Math.max(0, cfg.maxPathLength - fileOutDir.toString().length() - hash.length() - 8);
        if (leaf.length() > room) {
            leaf = leaf.substring(0, room);
        }
        return fileOutDir.resolve(leaf + "_" + hash + ".html");
    }

    private static String statusMark(Status s) {
        switch (s) {
            case SAME: return "OK";
            case DIFF: return "差分";
            case ERROR: return "NG";
            default: return "--";
        }
    }

    private void applyExitCode(CompareResult r, ExecResult ex) {
        r.exitCode = ex.exitCode;
        if (ex.timedOut) {
            r.status = Status.ERROR;
            r.message = "タイムアウト（" + cfg.timeoutSec + "秒）";
            return;
        }
        if (cfg.enableExitCode) {
            // /enableexitcode 指定時の終了コード: 0=同一, 1=差分あり, 2=エラー
            switch (ex.exitCode) {
                case 0:
                    r.status = Status.SAME;
                    break;
                case 1:
                    r.status = Status.DIFF;
                    break;
                default:
                    r.status = Status.ERROR;
                    r.message = "WinMerge 異常終了 (exit=" + ex.exitCode + ") " + ex.output.trim();
                    break;
            }
        } else if (ex.exitCode != 0) {
            r.status = Status.ERROR;
            r.message = "WinMerge 異常終了 (exit=" + ex.exitCode + ") " + ex.output.trim();
        } else {
            // 終了コードが使えない版では、内容比較で一致／差分を判定する
            r.status = contentEquals(r.left, r.right) ? Status.SAME : Status.DIFF;
        }
        if (r.report != null && !Files.isRegularFile(r.report)) {
            if (r.status != Status.ERROR) {
                r.message = "レポートファイルが生成されませんでした";
            }
            r.report = null;
        }
    }

    /**
     * WinMerge のコマンドラインを組み立てる。
     * <ul>
     *   <li>/r  : サブフォルダを含めて比較（フォルダ比較時のみ）</li>
     *   <li>/e  : ESC で終了可能</li>
     *   <li>/u  : 最近使ったファイル一覧に追加しない</li>
     *   <li>/minimize /noninteractive : 最小化で起動しレポート生成後に自動終了</li>
     *   <li>/enableexitcode : 比較結果を終了コードで返す（0=一致, 1=差分, 2=エラー）。
     *       <b>これを付けないと WinMerge は常に 0 を返すため、全件「一致」と判定されてしまう。</b>
     *       WinMerge 2.16 以降で利用可能。古い版では --no-exitcode を指定してファイル内容比較で代替する。</li>
     *   <li>/cfg ReportFiles/ReportType=2 : レポート形式を HTML にする</li>
     *   <li>/or &lt;path&gt; : レポートの出力先</li>
     * </ul>
     * 「選択されたファイルは同一です」のダイアログが出る環境では、一度手動で
     * 「次回から表示しない」にチェックを入れるか、--winmerge-arg で設定を追加すること。
     */
    private List<String> buildWinMergeCommand(boolean folder, Path left, Path right, Path report) {
        List<String> cmd = new ArrayList<>();
        cmd.add(cfg.winMerge);
        if (folder) {
            cmd.add("/r");
        }
        cmd.add("/e");
        cmd.add("/u");
        cmd.add("/minimize");
        cmd.add("/noninteractive");
        if (cfg.enableExitCode) {
            cmd.add("/enableexitcode");
        }
        cmd.add("/cfg");
        cmd.add("ReportFiles/ReportType=2");     // 2 = Simple HTML
        cmd.add("/cfg");
        cmd.add("Settings/DirViewExpandSubdirs=1");
        cmd.addAll(cfg.extraArgs);
        cmd.add("/or");
        cmd.add(report.toString());
        cmd.add(left.toString());
        cmd.add(right.toString());
        return cmd;
    }

    /**
     * ファイル同士／フォルダ同士の内容が一致するかをバイト単位で判定する。
     * WinMerge の /enableexitcode が使えない場合の代替判定に使う
     * （WinMerge のフィルタや改行コード設定は反映されない点に注意）。
     */
    private static boolean contentEquals(Path left, Path right) {
        try {
            if (Files.isDirectory(left) && Files.isDirectory(right)) {
                Set<String> names = listRelativeFiles(left, right);
                for (String rel : names) {
                    if (!contentEquals(left.resolve(rel), right.resolve(rel))) {
                        return false;
                    }
                }
                return true;
            }
            if (!Files.isRegularFile(left) || !Files.isRegularFile(right)) {
                return false;
            }
            if (Files.size(left) != Files.size(right)) {
                return false;
            }
            try (InputStream a = new java.io.BufferedInputStream(Files.newInputStream(left));
                 InputStream b = new java.io.BufferedInputStream(Files.newInputStream(right))) {
                int x;
                while ((x = a.read()) >= 0) {
                    if (x != b.read()) {
                        return false;
                    }
                }
                return b.read() < 0;
            }
        } catch (IOException e) {
            return false;
        }
    }

    /** left / right 配下のファイルを再帰的に列挙し、相対パスの和集合をソートして返す。 */
    private static Set<String> listRelativeFiles(Path leftDir, Path rightDir) throws IOException {
        Set<String> all = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        all.addAll(listRelativeFiles(leftDir));
        all.addAll(listRelativeFiles(rightDir));
        return all;
    }

    private static List<String> listRelativeFiles(Path dir) throws IOException {
        if (!Files.isDirectory(dir)) {
            return List.of();
        }
        try (Stream<Path> s = Files.walk(dir)) {
            return s.filter(Files::isRegularFile)
                    .map(p -> dir.relativize(p).toString().replace('\\', '/'))
                    .collect(Collectors.toList());
        }
    }

    // ----------------------------------------------------------------------
    // 外部プロセス実行
    // ----------------------------------------------------------------------

    static final class ExecResult {
        int exitCode = -1;
        boolean timedOut;
        String output = "";
    }

    private ExecResult exec(List<String> cmd) throws IOException, InterruptedException {
        ExecResult res = new ExecResult();
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectErrorStream(true);
        Process proc = pb.start();
        StringBuilder sb = new StringBuilder();
        Thread reader = new Thread(() -> {
            try (InputStream in = proc.getInputStream()) {
                byte[] buf = new byte[4096];
                int n;
                while ((n = in.read(buf)) > 0) {
                    sb.append(new String(buf, 0, n, NATIVE_CHARSET));
                }
            } catch (IOException ignore) {
                // プロセス終了時の読み取り失敗は無視する
            }
        });
        reader.setDaemon(true);
        reader.start();

        if (!proc.waitFor(cfg.timeoutSec, TimeUnit.SECONDS)) {
            proc.destroyForcibly();
            proc.waitFor();
            res.timedOut = true;
        }
        reader.join(3000);
        res.exitCode = res.timedOut ? -1 : proc.exitValue();
        res.output = sb.toString();
        return res;
    }

    // ======================================================================
    // 4) 一覧 HTML
    // ======================================================================

    private void writeIndexHtml(Path out) throws IOException {
        StringBuilder sb = new StringBuilder();
        sb.append("<!DOCTYPE html>\n<html lang=\"ja\"><head><meta charset=\"UTF-8\">\n")
          .append("<title>比較レポート一覧</title>\n<style>\n")
          .append("body{font-family:'Meiryo UI',sans-serif;font-size:13px;margin:16px;}\n")
          .append("h1{font-size:18px;} h2{font-size:15px;margin-top:24px;border-left:6px solid #4472c4;padding-left:8px;}\n")
          .append("table{border-collapse:collapse;margin-bottom:8px;} th,td{border:1px solid #bfbfbf;padding:3px 8px;}\n")
          .append("th{background:#d9e1f2;} .same{color:#1f7a1f;} .diff{color:#c00000;font-weight:bold;}\n")
          .append(".err{color:#ff6600;font-weight:bold;} .only{color:#7f7f7f;}\n</style></head><body>\n")
          .append("<h1>WinMerge 比較レポート一覧</h1>\n")
          .append("<p>生成日時: ").append(esc(LocalDateTime.now().format(TS_HUMAN))).append("<br>\n")
          .append("対象ルート: ").append(esc(cfg.root.toString())).append("</p>\n");

        for (TestCase tc : testCases) {
            sb.append("<h2>").append(esc(tc.name)).append(" <span style=\"font-weight:normal;font-size:12px;color:#666\">(")
              .append(esc(tc.rel.toString())).append(")</span></h2>\n");
            sb.append("<table><tr><th>比較</th><th>種別</th><th>対象</th><th>結果</th><th>レポート</th><th>備考</th></tr>\n");
            for (CompareResult r : tc.results) {
                sb.append("<tr><td>").append(esc(r.pairLabel)).append("</td><td>").append(esc(r.kind))
                  .append("</td><td>").append(esc(r.name)).append("</td><td class=\"")
                  .append(cssClass(r.status)).append("\">").append(esc(r.status.label)).append("</td><td>");
                if (r.report != null) {
                    String href = encodeForUri(out.getParent().relativize(r.report).toString().replace('\\', '/'));
                    sb.append("<a href=\"").append(href).append("\">HTML</a>");
                } else {
                    sb.append("-");
                }
                sb.append("</td><td>").append(esc(r.message)).append("</td></tr>\n");
            }
            sb.append("</table>\n");
        }
        sb.append("</body></html>\n");
        Files.writeString(out, sb.toString(), StandardCharsets.UTF_8);
    }

    private static String cssClass(Status s) {
        switch (s) {
            case SAME: return "same";
            case DIFF: return "diff";
            case ERROR: return "err";
            default: return "only";
        }
    }

    // ======================================================================
    // 5) Excel 出力
    // ======================================================================

    private void writeExcel() {
        if ("none".equals(cfg.excelMode)) {
            System.out.println("[INFO] Excel 出力はスキップされました（--excel-mode none）");
            return;
        }
        try {
            Files.createDirectories(cfg.excelPath.getParent());
        } catch (IOException e) {
            System.err.println("[ERROR] Excel 出力先を作成できません: " + e.getMessage());
            return;
        }

        boolean comCandidate = "com".equals(cfg.excelMode)
                || ("auto".equals(cfg.excelMode) && isWindows());
        if (comCandidate) {
            try {
                if (writeExcelViaCom()) {
                    System.out.println("[INFO] Excel ブックを作成しました（COM 貼り付け）: " + cfg.excelPath);
                    return;
                }
            } catch (Exception e) {
                System.err.println("[WARN] Excel COM 自動化に失敗しました: " + e.getMessage());
            }
            if ("com".equals(cfg.excelMode)) {
                System.err.println("[ERROR] Excel COM 出力に失敗しました。--excel-mode xlsx を試してください。");
                return;
            }
            System.err.println("[WARN] xlsx 直接生成へフォールバックします。");
        }

        try {
            writeXlsx();
            System.out.println("[INFO] Excel ブックを作成しました（xlsx 直接生成）: " + cfg.excelPath);
        } catch (IOException e) {
            System.err.println("[ERROR] xlsx の生成に失敗しました: " + e.getMessage());
        }
    }

    private static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("win");
    }

    // ----------------------------------------------------------------------
    // 5-a) Excel COM（VBScript）版: HTML レポートを実際にシートへ貼り付ける
    // ----------------------------------------------------------------------

    private boolean writeExcelViaCom() throws IOException, InterruptedException {
        Path vbs = cfg.reportRoot.resolve("_build_excel.vbs");
        Files.writeString(vbs, buildVbScript(), StandardCharsets.UTF_16LE);
        // UTF-16LE の BOM を付与（cscript が Unicode スクリプトとして解釈できるようにする）
        byte[] body = Files.readAllBytes(vbs);
        byte[] withBom = new byte[body.length + 2];
        withBom[0] = (byte) 0xFF;
        withBom[1] = (byte) 0xFE;
        System.arraycopy(body, 0, withBom, 2, body.length);
        Files.write(vbs, withBom);

        ExecResult ex = exec(List.of("cscript", "//nologo", "//B", vbs.toString()));
        if (ex.timedOut || ex.exitCode != 0) {
            System.err.println("[WARN] cscript 実行結果: exit=" + ex.exitCode + " " + ex.output.trim());
            return false;
        }
        return Files.isRegularFile(cfg.excelPath);
    }

    private String buildVbScript() {
        StringBuilder sb = new StringBuilder();
        sb.append("Option Explicit\r\n")
          .append("Dim xl, wb, ws, fso, gRow, gFirst\r\n")
          .append("Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n")
          .append("Set xl = CreateObject(\"Excel.Application\")\r\n")
          .append("xl.Visible = False\r\n")
          .append("xl.DisplayAlerts = False\r\n")
          .append("Set wb = xl.Workbooks.Add\r\n")
          .append("Do While wb.Worksheets.Count > 1\r\n")
          .append("  wb.Worksheets(wb.Worksheets.Count).Delete\r\n")
          .append("Loop\r\n")
          .append("gFirst = True\r\n")
          .append("gRow = 1\r\n");

        // --- サブルーチン定義 ---
        sb.append("\r\n")
          .append("Sub AddSheet(sname)\r\n")
          .append("  If gFirst Then\r\n")
          .append("    Set ws = wb.Worksheets(1)\r\n")
          .append("    gFirst = False\r\n")
          .append("  Else\r\n")
          .append("    wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count)\r\n")
          .append("    Set ws = wb.Worksheets(wb.Worksheets.Count)\r\n")
          .append("  End If\r\n")
          .append("  ws.Name = sname\r\n")
          .append("  gRow = 1\r\n")
          .append("End Sub\r\n\r\n")
          .append("Sub AddLine(txt, isBold)\r\n")
          .append("  ws.Cells(gRow, 1).Value = txt\r\n")
          .append("  ws.Cells(gRow, 1).Font.Bold = isBold\r\n")
          .append("  gRow = gRow + 1\r\n")
          .append("End Sub\r\n\r\n")
          .append("Sub AddRow4(c1, c2, c3, c4)\r\n")
          .append("  ws.Cells(gRow, 1).Value = c1\r\n")
          .append("  ws.Cells(gRow, 2).Value = c2\r\n")
          .append("  ws.Cells(gRow, 3).Value = c3\r\n")
          .append("  ws.Cells(gRow, 4).Value = c4\r\n")
          .append("  gRow = gRow + 1\r\n")
          .append("End Sub\r\n\r\n")
          .append("Sub AddReport(title, htmlPath)\r\n")
          .append("  Dim src, rows\r\n")
          .append("  AddLine title, True\r\n")
          .append("  If Not fso.FileExists(htmlPath) Then\r\n")
          .append("    AddLine \"  レポートがありません: \" & htmlPath, False\r\n")
          .append("    gRow = gRow + 1\r\n")
          .append("    Exit Sub\r\n")
          .append("  End If\r\n")
          .append("  ws.Hyperlinks.Add ws.Cells(gRow, 1), htmlPath, \"\", htmlPath, htmlPath\r\n")
          .append("  gRow = gRow + 1\r\n")
          .append("  On Error Resume Next\r\n")
          .append("  Set src = xl.Workbooks.Open(htmlPath, False, True)\r\n")
          .append("  If Err.Number <> 0 Then\r\n")
          .append("    Err.Clear\r\n")
          .append("    On Error GoTo 0\r\n")
          .append("    AddLine \"  HTML の取り込みに失敗しました\", False\r\n")
          .append("    gRow = gRow + 1\r\n")
          .append("    Exit Sub\r\n")
          .append("  End If\r\n")
          .append("  rows = src.Worksheets(1).UsedRange.Rows.Count\r\n")
          .append("  src.Worksheets(1).UsedRange.Copy\r\n")
          .append("  wb.Activate\r\n")
          .append("  ws.Activate\r\n")
          .append("  ws.Paste ws.Cells(gRow, 1)\r\n")
          .append("  xl.CutCopyMode = False\r\n")
          .append("  src.Close False\r\n")
          .append("  If Err.Number <> 0 Then Err.Clear\r\n")
          .append("  On Error GoTo 0\r\n")
          .append("  gRow = gRow + rows + 2\r\n")
          .append("End Sub\r\n\r\n");

        // --- データ部 ---
        sb.append("On Error Resume Next\r\n");
        Set<String> usedNames = new LinkedHashSet<>();

        // サマリシート
        sb.append(call("AddSheet", uniqueSheetName("サマリ", usedNames)));
        sb.append(call("AddLine", "WinMerge 比較レポート サマリ", "True"));
        sb.append(call("AddLine", "生成日時: " + LocalDateTime.now().format(TS_HUMAN), "False"));
        sb.append(call("AddLine", "対象ルート: " + cfg.root, "False"));
        sb.append("gRow = gRow + 1\r\n");
        sb.append(call("AddRow4", "テストケース", "比較", "対象", "結果"));
        for (TestCase tc : testCases) {
            for (CompareResult r : tc.results) {
                sb.append(call("AddRow4", tc.name, r.pairLabel, r.kind + " " + r.name,
                        r.status.label + (r.message.isEmpty() ? "" : " / " + r.message)));
            }
        }
        sb.append("ws.Columns(\"A:D\").AutoFit\r\n");

        // テストケースごとのシート
        for (TestCase tc : testCases) {
            sb.append(call("AddSheet", uniqueSheetName(tc.name, usedNames)));
            sb.append(call("AddLine", "テストケース: " + tc.name, "True"));
            sb.append(call("AddLine", "パス: " + tc.dir, "False"));
            sb.append("gRow = gRow + 1\r\n");
            for (CompareResult r : tc.results) {
                String title = "[" + r.pairLabel + "] " + r.kind + " : " + r.name
                        + "  => " + r.status.label + (r.message.isEmpty() ? "" : " (" + r.message + ")");
                if (r.report != null) {
                    sb.append(call("AddReport", title, r.report.toString()));
                } else {
                    sb.append(call("AddLine", title, "True"));
                    sb.append("gRow = gRow + 1\r\n");
                }
            }
            sb.append("ws.Columns(\"A:H\").AutoFit\r\n");
        }

        sb.append("\r\n")
          .append("wb.Worksheets(1).Activate\r\n")
          .append("wb.SaveAs ").append(vbsStr(cfg.excelPath.toString())).append(", 51\r\n")
          .append("wb.Close False\r\n")
          .append("xl.Quit\r\n")
          .append("Set ws = Nothing\r\n")
          .append("Set wb = Nothing\r\n")
          .append("Set xl = Nothing\r\n")
          .append("If Err.Number <> 0 Then\r\n")
          .append("  WScript.StdErr.WriteLine \"VBS ERROR: \" & Err.Number & \" \" & Err.Description\r\n")
          .append("  WScript.Quit 1\r\n")
          .append("End If\r\n")
          .append("WScript.Quit 0\r\n");
        return sb.toString();
    }

    /** VBScript のサブルーチン呼び出し行を組み立てる（True/False はそのままリテラルとして渡す）。 */
    private static String call(String sub, String... args) {
        StringBuilder sb = new StringBuilder(sub);
        for (int i = 0; i < args.length; i++) {
            sb.append(i == 0 ? " " : ", ");
            String a = args[i];
            if ("True".equals(a) || "False".equals(a)) {
                sb.append(a);
            } else {
                sb.append(vbsStr(a));
            }
        }
        return sb.append("\r\n").toString();
    }

    private static String vbsStr(String s) {
        String v = s == null ? "" : s.replace("\"", "\"\"");
        v = v.replace("\r", " ").replace("\n", " ");
        return "\"" + v + "\"";
    }

    // ----------------------------------------------------------------------
    // 5-b) xlsx 直接生成（外部ライブラリ不要）
    // ----------------------------------------------------------------------

    // セルスタイル ID（styles.xml の cellXfs の並びと対応）
    private static final int ST_NORMAL = 0;
    private static final int ST_BOLD = 1;
    private static final int ST_TITLE = 2;
    private static final int ST_LINK = 3;
    private static final int ST_HEADER = 4;
    private static final int ST_CELL = 5;

    static final class Cell {
        final String text;
        final int style;
        final String hyperlink; // 外部リンク先（null 可）

        Cell(String text, int style, String hyperlink) {
            this.text = text;
            this.style = style;
            this.hyperlink = hyperlink;
        }
    }

    static final class SheetData {
        final String name;
        final List<List<Cell>> rows = new ArrayList<>();
        double[] colWidths = {40, 18, 18, 18, 18, 18, 18, 18};

        SheetData(String name) {
            this.name = name;
        }

        void add(Cell... cells) {
            rows.add(Arrays.asList(cells));
        }

        void addText(String text, int style) {
            add(new Cell(text, style, null));
        }

        void blank() {
            rows.add(List.of());
        }
    }

    private void writeXlsx() throws IOException {
        List<SheetData> sheets = new ArrayList<>();
        Set<String> usedNames = new LinkedHashSet<>();

        // --- サマリシート ---
        SheetData summary = new SheetData(uniqueSheetName("サマリ", usedNames));
        summary.colWidths = new double[]{24, 22, 14, 44, 12, 40, 40};
        summary.addText("WinMerge 比較レポート サマリ", ST_TITLE);
        summary.addText("生成日時: " + LocalDateTime.now().format(TS_HUMAN), ST_NORMAL);
        summary.addText("対象ルート: " + cfg.root, ST_NORMAL);
        summary.blank();
        summary.add(header("テストケース"), header("比較"), header("種別"), header("対象"),
                header("結果"), header("備考"), header("レポート"));
        for (TestCase tc : testCases) {
            for (CompareResult r : tc.results) {
                summary.add(
                        new Cell(tc.name, ST_CELL, null),
                        new Cell(r.pairLabel, ST_CELL, null),
                        new Cell(r.kind, ST_CELL, null),
                        new Cell(r.name, ST_CELL, null),
                        new Cell(r.status.label, ST_CELL, null),
                        new Cell(r.message, ST_CELL, null),
                        r.report == null
                                ? new Cell("-", ST_CELL, null)
                                : new Cell(cfg.reportRoot.relativize(r.report).toString().replace('\\', '/'),
                                        ST_LINK, hyperlinkTarget(r.report)));
            }
        }
        sheets.add(summary);

        // --- テストケースごとのシート ---
        for (TestCase tc : testCases) {
            SheetData sd = new SheetData(uniqueSheetName(tc.name, usedNames));
            sd.addText("テストケース: " + tc.name, ST_TITLE);
            sd.addText("パス: " + tc.dir, ST_NORMAL);
            sd.blank();
            for (CompareResult r : tc.results) {
                sd.addText("[" + r.pairLabel + "] " + r.kind + " : " + r.name
                        + "  => " + r.status.label + (r.message.isEmpty() ? "" : " (" + r.message + ")"), ST_BOLD);
                if (r.report != null) {
                    sd.add(new Cell("レポートを開く: " + cfg.reportRoot.relativize(r.report).toString().replace('\\', '/'),
                            ST_LINK, hyperlinkTarget(r.report)));
                    List<List<String>> table = extractHtmlRows(r.report, cfg.maxTableRows);
                    boolean first = true;
                    for (List<String> row : table) {
                        List<Cell> cells = new ArrayList<>();
                        for (String v : row) {
                            cells.add(new Cell(v, first ? ST_HEADER : ST_CELL, null));
                        }
                        sd.rows.add(cells);
                        first = false;
                    }
                    if (table.isEmpty()) {
                        sd.addText("（レポート本文を解析できませんでした。上のリンクから HTML を参照してください）", ST_NORMAL);
                    }
                }
                sd.blank();
            }
            sheets.add(sd);
        }

        writeXlsxFile(cfg.excelPath, sheets);
    }

    private static Cell header(String text) {
        return new Cell(text, ST_HEADER, null);
    }

    private String hyperlinkTarget(Path report) {
        Path base = cfg.excelPath.getParent();
        String rel;
        try {
            rel = base.relativize(report).toString().replace('\\', '/');
        } catch (IllegalArgumentException e) {
            return report.toUri().toString();
        }
        return encodeForUri(rel);
    }

    /** xlsx（OOXML）を java.util.zip だけで書き出す。 */
    private static void writeXlsxFile(Path out, List<SheetData> sheets) throws IOException {
        try (ZipOutputStream zip = new ZipOutputStream(Files.newOutputStream(out), StandardCharsets.UTF_8)) {
            // [Content_Types].xml
            StringBuilder ct = new StringBuilder();
            ct.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
              .append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">")
              .append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>")
              .append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>")
              .append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>")
              .append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 1; i <= sheets.size(); i++) {
                ct.append("<Override PartName=\"/xl/worksheets/sheet").append(i)
                  .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }
            ct.append("</Types>");
            put(zip, "[Content_Types].xml", ct.toString());

            put(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                    + "</Relationships>");

            // xl/workbook.xml
            StringBuilder wb = new StringBuilder();
            wb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
              .append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ")
              .append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            for (int i = 0; i < sheets.size(); i++) {
                wb.append("<sheet name=\"").append(esc(sheets.get(i).name)).append("\" sheetId=\"")
                  .append(i + 1).append("\" r:id=\"rId").append(i + 1).append("\"/>");
            }
            wb.append("</sheets></workbook>");
            put(zip, "xl/workbook.xml", wb.toString());

            // xl/_rels/workbook.xml.rels
            StringBuilder wr = new StringBuilder();
            wr.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
              .append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < sheets.size(); i++) {
                wr.append("<Relationship Id=\"rId").append(i + 1)
                  .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
                  .append(i + 1).append(".xml\"/>");
            }
            wr.append("<Relationship Id=\"rId").append(sheets.size() + 1)
              .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>")
              .append("</Relationships>");
            put(zip, "xl/_rels/workbook.xml.rels", wr.toString());

            put(zip, "xl/styles.xml", stylesXml());

            for (int i = 0; i < sheets.size(); i++) {
                SheetData sd = sheets.get(i);
                List<String> links = new ArrayList<>();
                put(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", sheetXml(sd, links));
                if (!links.isEmpty()) {
                    StringBuilder sr = new StringBuilder();
                    sr.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
                      .append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    for (int k = 0; k < links.size(); k++) {
                        sr.append("<Relationship Id=\"rId").append(k + 1)
                          .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"")
                          .append(esc(links.get(k))).append("\" TargetMode=\"External\"/>");
                    }
                    sr.append("</Relationships>");
                    put(zip, "xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels", sr.toString());
                }
            }
        }
    }

    private static void put(ZipOutputStream zip, String name, String content) throws IOException {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(content.getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
    }

    private static String stylesXml() {
        String font = "<sz val=\"11\"/><name val=\"Meiryo UI\"/>";
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<fonts count=\"4\">"
                + "<font>" + font + "</font>"
                + "<font><b/>" + font + "</font>"
                + "<font><b/><sz val=\"14\"/><name val=\"Meiryo UI\"/></font>"
                + "<font><u/><color rgb=\"FF0563C1\"/>" + font + "</font>"
                + "</fonts>"
                + "<fills count=\"3\">"
                + "<fill><patternFill patternType=\"none\"/></fill>"
                + "<fill><patternFill patternType=\"gray125\"/></fill>"
                + "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFD9E1F2\"/><bgColor indexed=\"64\"/></patternFill></fill>"
                + "</fills>"
                + "<borders count=\"2\">"
                + "<border><left/><right/><top/><bottom/><diagonal/></border>"
                + "<border><left style=\"thin\"><color rgb=\"FFBFBFBF\"/></left><right style=\"thin\"><color rgb=\"FFBFBFBF\"/></right>"
                + "<top style=\"thin\"><color rgb=\"FFBFBFBF\"/></top><bottom style=\"thin\"><color rgb=\"FFBFBFBF\"/></bottom><diagonal/></border>"
                + "</borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"6\">"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                + "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\"/>"
                + "</cellXfs>"
                + "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
                + "</styleSheet>";
    }

    private static String sheetXml(SheetData sd, List<String> outLinks) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
          .append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ")
          .append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

        sb.append("<cols>");
        for (int c = 0; c < sd.colWidths.length; c++) {
            sb.append("<col min=\"").append(c + 1).append("\" max=\"").append(c + 1)
              .append("\" width=\"").append(sd.colWidths[c]).append("\" customWidth=\"1\"/>");
        }
        sb.append("</cols>");

        StringBuilder hyper = new StringBuilder();
        sb.append("<sheetData>");
        for (int rIdx = 0; rIdx < sd.rows.size(); rIdx++) {
            List<Cell> row = sd.rows.get(rIdx);
            int rowNum = rIdx + 1;
            if (row.isEmpty()) {
                continue;
            }
            sb.append("<row r=\"").append(rowNum).append("\">");
            for (int cIdx = 0; cIdx < row.size(); cIdx++) {
                Cell cell = row.get(cIdx);
                if (cell == null) {
                    continue;
                }
                String ref = colName(cIdx) + rowNum;
                String text = sanitizeCell(cell.text);
                sb.append("<c r=\"").append(ref).append("\" s=\"").append(cell.style).append("\"");
                if (text.isEmpty()) {
                    sb.append("/>");
                } else {
                    sb.append(" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
                      .append(esc(text)).append("</t></is></c>");
                }
                if (cell.hyperlink != null && !cell.hyperlink.isEmpty()) {
                    outLinks.add(cell.hyperlink);
                    hyper.append("<hyperlink ref=\"").append(ref).append("\" r:id=\"rId")
                         .append(outLinks.size()).append("\"/>");
                }
            }
            sb.append("</row>");
        }
        sb.append("</sheetData>");
        if (hyper.length() > 0) {
            sb.append("<hyperlinks>").append(hyper).append("</hyperlinks>");
        }
        sb.append("</worksheet>");
        return sb.toString();
    }

    private static String colName(int index) {
        StringBuilder sb = new StringBuilder();
        int i = index;
        while (i >= 0) {
            sb.insert(0, (char) ('A' + (i % 26)));
            i = i / 26 - 1;
        }
        return sb.toString();
    }

    /** Excel のセル制限（32767 文字）と XML で使えない制御文字に対処する。 */
    private static String sanitizeCell(String s) {
        if (s == null) {
            return "";
        }
        StringBuilder sb = new StringBuilder(s.length());
        for (int i = 0; i < s.length() && sb.length() < 32000; i++) {
            char ch = s.charAt(i);
            if (ch == '\t') {
                sb.append(' ');
            } else if (ch == '\n' || ch == '\r') {
                sb.append(' ');
            } else if (ch < 0x20) {
                // 制御文字は捨てる
                continue;
            } else {
                sb.append(ch);
            }
        }
        return sb.toString().trim();
    }

    // ======================================================================
    // HTML レポート解析（xlsx 貼り付け用）
    // ======================================================================

    private static final Pattern P_TR = Pattern.compile("(?is)<tr\\b[^>]*>(.*?)</tr>");
    private static final Pattern P_TD = Pattern.compile("(?is)<t([dh])\\b[^>]*>(.*?)</t\\1>");
    private static final Pattern P_TAG = Pattern.compile("(?s)<[^>]*>");
    private static final Pattern P_META_CHARSET =
            Pattern.compile("(?i)<meta[^>]*charset\\s*=\\s*[\"']?\\s*([A-Za-z0-9_\\-]+)");

    /** HTML レポートから表を抽出して行×列の文字列にする。表が無ければ本文テキストを 1 列で返す。 */
    static List<List<String>> extractHtmlRows(Path html, int maxRows) {
        String content;
        try {
            content = readHtml(html);
        } catch (IOException e) {
            return List.of();
        }
        List<List<String>> rows = new ArrayList<>();
        Matcher mTr = P_TR.matcher(content);
        while (mTr.find() && rows.size() < maxRows) {
            List<String> cells = new ArrayList<>();
            Matcher mTd = P_TD.matcher(mTr.group(1));
            while (mTd.find()) {
                cells.add(htmlToText(mTd.group(2)));
            }
            if (!cells.isEmpty()) {
                rows.add(cells);
            }
        }
        if (!rows.isEmpty()) {
            return rows;
        }
        // 表が見つからない場合は本文テキストを行単位で取り込む
        String body = content.replaceAll("(?is)<(script|style)\\b.*?</\\1>", " ")
                             .replaceAll("(?i)<br\\s*/?>", "\n")
                             .replaceAll("(?i)</(p|div|tr|h[1-6])>", "\n");
        for (String line : P_TAG.matcher(body).replaceAll("").split("\n")) {
            String t = unescapeHtml(line).trim();
            if (!t.isEmpty()) {
                rows.add(List.of(t));
                if (rows.size() >= maxRows) {
                    break;
                }
            }
        }
        return rows;
    }

    private static String readHtml(Path html) throws IOException {
        byte[] bytes = Files.readAllBytes(html);
        String probe = new String(bytes, 0, Math.min(bytes.length, 2048), StandardCharsets.ISO_8859_1);
        Charset cs = StandardCharsets.UTF_8;
        Matcher m = P_META_CHARSET.matcher(probe);
        if (m.find()) {
            try {
                cs = Charset.forName(m.group(1));
            } catch (IllegalCharsetNameException | UnsupportedCharsetException ignore) {
                cs = StandardCharsets.UTF_8;
            }
        }
        return new String(bytes, cs);
    }

    private static String htmlToText(String htmlFragment) {
        String s = htmlFragment.replaceAll("(?i)<br\\s*/?>", " ");
        s = P_TAG.matcher(s).replaceAll("");
        return unescapeHtml(s).replaceAll("\\s+", " ").trim();
    }

    private static final Map<String, String> ENTITIES = Map.of(
            "nbsp", "\u00A0", "lt", "<", "gt", ">", "amp", "&", "quot", "\"",
            "apos", "'", "copy", "\u00A9", "hellip", "\u2026");

    private static String unescapeHtml(String s) {
        if (s.indexOf('&') < 0) {
            return s;
        }
        StringBuilder sb = new StringBuilder(s.length());
        Matcher m = Pattern.compile("&(#x?[0-9A-Fa-f]+|[A-Za-z]+);").matcher(s);
        int last = 0;
        while (m.find()) {
            sb.append(s, last, m.start());
            String e = m.group(1);
            try {
                if (e.startsWith("#x") || e.startsWith("#X")) {
                    sb.appendCodePoint(Integer.parseInt(e.substring(2), 16));
                } else if (e.startsWith("#")) {
                    sb.appendCodePoint(Integer.parseInt(e.substring(1)));
                } else {
                    sb.append(ENTITIES.getOrDefault(e.toLowerCase(Locale.ROOT), m.group(0)));
                }
            } catch (RuntimeException ex) {
                sb.append(m.group(0));
            }
            last = m.end();
        }
        sb.append(s.substring(last));
        return sb.toString();
    }

    // ======================================================================
    // ユーティリティ
    // ======================================================================

    /** Excel のシート名制約（31 文字・禁止文字・重複不可）に合わせて整形する。 */
    static String uniqueSheetName(String raw, Set<String> used) {
        String name = raw == null || raw.isBlank() ? "sheet" : raw;
        name = name.replaceAll("[\\\\/\\*\\?\\[\\]:]", "_");
        if (name.startsWith("'")) {
            name = "_" + name.substring(1);
        }
        if (name.length() > 31) {
            name = name.substring(0, 31);
        }
        String base = name;
        int n = 2;
        while (!used.add(name.toLowerCase(Locale.ROOT))) {
            String suffix = "_" + n++;
            int keep = Math.min(base.length(), 31 - suffix.length());
            name = base.substring(0, keep) + suffix;
        }
        return name;
    }

    private static String esc(String s) {
        if (s == null) {
            return "";
        }
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\"", "&quot;").replace("'", "&#39;");
    }

    /** パス文字列を URI として安全な形にエンコードする（'/' は保持）。 */
    private static String encodeForUri(String path) {
        StringBuilder sb = new StringBuilder();
        for (byte b : path.getBytes(StandardCharsets.UTF_8)) {
            int c = b & 0xFF;
            boolean safe = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')
                    || c == '/' || c == '-' || c == '_' || c == '.' || c == '~';
            if (safe) {
                sb.append((char) c);
            } else {
                sb.append('%').append(String.format("%02X", c));
            }
        }
        return sb.toString();
    }

    private static void deleteRecursively(Path dir) throws IOException {
        try (Stream<Path> s = Files.walk(dir)) {
            s.sorted(Comparator.reverseOrder()).forEach(p -> {
                try {
                    Files.deleteIfExists(p);
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            });
        } catch (UncheckedIOException e) {
            throw e.getCause();
        }
    }

    /** 未使用だが将来の拡張用（プロセス出力のバイト読み取り）。 */
    @SuppressWarnings("unused")
    private static byte[] readAll(InputStream in) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buf = new byte[8192];
        int n;
        while ((n = in.read(buf)) > 0) {
            bos.write(buf, 0, n);
        }
        return bos.toByteArray();
    }
}
