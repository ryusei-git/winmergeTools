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
 * Generates WinMerge comparison reports for a tree of test cases and collects them into an
 * Excel workbook. Everything lives in this one file on purpose: it is meant to be copied to a
 * Windows machine and run with nothing but a JDK, WinMerge and (optionally) Edge or Excel.
 *
 * <p>Pipeline, in the order run() executes it:
 * <ol>
 *   <li>Mirror the directory tree under --root (default: current directory) into the report
 *       folder, so every generated report has a home matching the source layout.</li>
 *   <li>Find test case directories: any directory holding one of the configured pair
 *       directories (INPUT/OUTPUT and LOG/LOG_COMPARE by default). Nested ones count too.</li>
 *   <li>For each pair, run one recursive WinMerge folder comparison (/r) and then one file
 *       comparison per file present on both sides; a file on one side only is recorded
 *       without invoking WinMerge, which needs two existing files.</li>
 *   <li>Write report/index.html listing every comparison.</li>
 *   <li>Build an Excel workbook with one sheet per test case.</li>
 * </ol>
 *
 * <p>Usage (Windows):
 * <pre>
 *   java WinMergeReportTool.java
 *   java WinMergeReportTool.java --winmerge "C:\Program Files\WinMerge\WinMergeU.exe"
 *   java WinMergeReportTool.java --root D:\tests --report D:\tests\report
 *   java WinMergeReportTool.java --excel-mode image
 * </pre>
 *
 * <p>Three ways to get the reports into Excel, chosen with --excel-mode:
 * <ul>
 *   <li>com   - drive Excel over COM from a generated VBScript and paste the HTML as a table,
 *               keeping its formatting. Needs Windows plus an installed Excel.</li>
 *   <li>image - screenshot each report with headless Edge/Chrome and embed the PNGs. Keeps the
 *               exact appearance and needs no Excel, but the text is not searchable.</li>
 *   <li>xlsx  - parse the HTML tables and write the text into cells. Needs nothing at all.</li>
 * </ul>
 * The default, auto, tries com on Windows and falls back to xlsx.
 *
 * <h2>Maintenance notes</h2>
 * <ul>
 *   <li>Java 11 is the language floor; the code compiles clean with --release 11 and 17. Do not
 *       introduce APIs newer than that without also updating README.</li>
 *   <li>Comments are English on purpose. The source is UTF-8 and JDK 17 and older default the
 *       source encoding to the platform charset (MS932 on Japanese Windows), so non-ASCII in
 *       the source only survives when compiled with -encoding UTF-8. The Japanese string
 *       literals are user-facing console text and must stay; comments must not add to them.</li>
 *   <li>No third-party libraries. The .xlsx is assembled by hand with java.util.zip, so any
 *       change to the workbook must keep the OOXML parts, relationship ids and content types
 *       consistent - see writeXlsxFile().</li>
 *   <li>Every file the tool reads or writes names its charset explicitly. Never rely on the
 *       platform default: it differs between the console, the JDK version and the OS.</li>
 * </ul>
 */
public final class WinMergeReportTool {

    /**
     * Charset for reading the output of child processes (WinMerge, cscript, the browser).
     *
     * <p>JDK 18 pinned file.encoding to UTF-8, but a Windows console still emits its own code
     * page (windows-31j on Japanese Windows), so decoding child output as UTF-8 would garble
     * every error message. native.encoding carries the real OS charset; it exists from JDK 17,
     * hence the fallback.
     */
    private static final Charset NATIVE_CHARSET = nativeCharset();

    private static Charset nativeCharset() {
        String name = System.getProperty("native.encoding");
        if (name != null && !name.isEmpty()) {
            try {
                return Charset.forName(name);
            } catch (IllegalCharsetNameException | UnsupportedCharsetException ignore) {
                // Unknown name: fall back to the platform default.
            }
        }
        return Charset.defaultCharset();
    }

    private static final DateTimeFormatter TS_HUMAN = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
    private static final DateTimeFormatter TS_FILE = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    // ======================================================================
    // Configuration
    // ======================================================================

    /** One pair of sibling directories to compare, e.g. INPUT against OUTPUT. */
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
        Path reportRoot;                 // defaults to <root>/report
        Path excelPath;                  // defaults to <reportRoot>/comparison_report_<timestamp>.xlsx
        String winMerge;                 // resolved path of WinMergeU.exe
        final List<ComparePair> pairs = new ArrayList<>();
        final List<String> extraArgs = new ArrayList<>();
        long timeoutSec = 180;
        String excelMode = "auto";       // auto | com | xlsx | image | none
        String browser;                  // Edge/Chrome used for screenshots (image mode only)
        int imageWidth = 1600;           // capture width in px
        int imageMaxHeight = 16000;      // capture height cap in px; PNGs beyond this get clipped
        String imageBorder = "000000";   // RRGGBB outline drawn around each embedded picture; "" for none
        int imageBorderWidthPt = 1;      // outline thickness in points
        int maxTableRows = 500;          // rows expanded into cells per report, to bound sheet size
        boolean cleanReport = false;     // delete the report folder before running
        boolean enableExitCode = true;   // pass /enableexitcode; false for WinMerge before 2.16
        int maxPathLength = 240;         // shorten report paths beyond this (MAX_PATH is 260)
    }

    // ======================================================================
    // Result model
    // ======================================================================

    enum Status {
        SAME("identical"),
        DIFF("different"),
        LEFT_ONLY("left only"),
        RIGHT_ONLY("right only"),
        ERROR("error"),
        SKIPPED("skipped");

        final String label;

        Status(String label) {
            this.label = label;
        }
    }

    static final class CompareResult {
        String kind;        // folder comparison or file comparison, as displayed
        String pairLabel;   // e.g. INPUT_vs_OUTPUT
        String name;        // relative file path, or the whole-folder marker for a folder comparison
        Path left;
        Path right;
        Path report;        // generated HTML report; null when WinMerge produced none
        Status status = Status.SKIPPED;
        String message = "";
        int exitCode = -1;
    }

    static final class TestCase {
        String name;                 // directory name; becomes the sheet name
        Path dir;                    // absolute path of the test case directory
        Path rel;                    // path relative to root, mirrored under the report folder
        final List<CompareResult> results = new ArrayList<>();
    }

    /**
     * Bad command line. Kept distinct from RuntimeException so main() can print the usage text
     * for argument mistakes only, and still report real failures as errors.
     */
    static final class UsageException extends RuntimeException {
        private static final long serialVersionUID = 1L;

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
    // Entry point
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
            System.err.println("[FATAL] Unexpected error: " + e);
            e.printStackTrace();
            System.exit(1);
        }
    }

    private static void printUsage() {
        System.out.println(String.join(System.lineSeparator(),
                "Usage: java WinMergeReportTool.java [options]",
                "",
                "  --root <dir>        root of the tree to compare (default: current directory)",
                "  --report <dir>      report output directory (default: <root>/report)",
                "  --excel <file>      workbook path (default: <report>/comparison_report_<time>.xlsx)",
                "  --winmerge <exe>    path to WinMergeU.exe (default: WINMERGE_PATH, then usual install dirs)",
                "  --pair <L>:<R>      directory pair to compare; repeatable",
                "                      default: --pair INPUT:OUTPUT --pair LOG:LOG_COMPARE",
                "  --excel-mode <mode> auto | com | image | xlsx | none (default: auto)",
                "                      image renders each report to PNG and embeds the pictures",
                "  --browser <exe>     Edge/Chrome used for PNG capture (default: auto-detect)",
                "  --image-width <px>  capture width (default: 1600)",
                "  --image-max-height <px> capture height cap (default: 16000)",
                "  --image-border <rgb> outline colour around each picture as RRGGBB,",
                "                      or none to draw no outline (default: 000000, black)",
                "  --image-border-width <pt> outline thickness in points (default: 1)",
                "  --timeout <sec>     timeout per WinMerge run (default: 180)",
                "  --max-rows <n>      max rows expanded into cells per report (default: 500)",
                "  --clean             delete the report directory before running",
                "  --no-exitcode       do not pass /enableexitcode; compare file contents in Java",
                "                      (for WinMerge 2.14 and older, which lacks that switch)",
                "  --max-path <n>      max report path length (default: 240; longer ones are shortened)",
                "  --winmerge-arg <a>  extra argument passed to WinMerge; repeatable",
                "  --help              show this help"));
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
                        throw new UsageException("--pair must be given as LEFT:RIGHT: " + v);
                    }
                    pairs.add(new ComparePair(v.substring(0, sep), v.substring(sep + 1)));
                    break;
                }
                case "--excel-mode":
                    cfg.excelMode = next(args, ++i, a).toLowerCase(Locale.ROOT);
                    if (!Arrays.asList("auto", "com", "xlsx", "image", "none").contains(cfg.excelMode)) {
                        throw new UsageException("--excel-mode must be auto|com|image|xlsx|none: " + cfg.excelMode);
                    }
                    break;
                case "--browser":
                    cfg.browser = next(args, ++i, a);
                    break;
                case "--image-width":
                    cfg.imageWidth = (int) parseLong(next(args, ++i, a), a);
                    break;
                case "--image-max-height":
                    cfg.imageMaxHeight = (int) parseLong(next(args, ++i, a), a);
                    break;
                case "--image-border": {
                    String v = next(args, ++i, a).trim();
                    if ("none".equalsIgnoreCase(v)) {
                        cfg.imageBorder = "";
                    } else {
                        String hex = v.startsWith("#") ? v.substring(1) : v;
                        if (!hex.matches("(?i)[0-9a-f]{6}")) {
                            throw new UsageException("--image-border must be RRGGBB or none: " + v);
                        }
                        cfg.imageBorder = hex.toUpperCase(Locale.ROOT);
                    }
                    break;
                }
                case "--image-border-width":
                    cfg.imageBorderWidthPt = (int) parseLong(next(args, ++i, a), a);
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
                    throw new UsageException("Unknown option: " + a);
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
            throw new UsageException(opt + " requires a value");
        }
        return args[i];
    }

    private static long parseLong(String value, String opt) {
        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException e) {
            throw new UsageException(opt + " requires a number: " + value);
        }
    }

    // ======================================================================
    // Main flow
    // ======================================================================

    private int run() throws IOException, InterruptedException {
        System.out.println("==================================================");
        System.out.println(" WinMerge test case comparison report");
        System.out.println("==================================================");
        System.out.println("  Root       : " + cfg.root);
        System.out.println("  Report dir : " + cfg.reportRoot);
        System.out.println("  Workbook   : " + cfg.excelPath);
        System.out.println("  Pairs      : " + cfg.pairs.stream().map(p -> p.label).collect(Collectors.joining(", ")));

        if (!Files.isDirectory(cfg.root)) {
            System.err.println("[ERROR] Root directory does not exist: " + cfg.root);
            return 1;
        }

        cfg.winMerge = resolveWinMerge(cfg.winMerge);
        System.out.println("  WinMerge   : " + (cfg.winMerge == null ? "(not found)" : cfg.winMerge));
        System.out.println();
        if (cfg.winMerge == null) {
            System.err.println("[ERROR] WinMergeU.exe not found. Pass --winmerge, or set WINMERGE_PATH.");
            return 1;
        }

        // Guard against --report pointing at (or above) the data being compared: --clean would
        // otherwise delete the test cases themselves.
        if (cfg.root.equals(cfg.reportRoot) || cfg.root.startsWith(cfg.reportRoot)) {
            System.err.println("[ERROR] The report directory contains the tree being compared. Choose another location.");
            System.err.println("        compared: " + cfg.root);
            System.err.println("        report  : " + cfg.reportRoot);
            return 1;
        }
        if (cfg.cleanReport && Files.exists(cfg.reportRoot)) {
            System.out.println("[INFO] Deleting the existing report directory: " + cfg.reportRoot);
            deleteRecursively(cfg.reportRoot);
        }
        Files.createDirectories(cfg.reportRoot);
        if (isWindows() && cfg.reportRoot.toString().length() > 150) {
            System.err.println("[WARN] The report path is long and may exceed MAX_PATH (260): "
                    + cfg.reportRoot.toString().length() + " characters");
            System.err.println("       Use --report to pick a shallower location, or tune --max-path.");
        }

        // Step 1: mirror the source tree so reports can sit beside their originals.
        int mirrored = mirrorDirectoryTree();
        System.out.println("[INFO] Mirrored " + mirrored + " directories");

        // Step 2: locate the test case directories.
        collectTestCases();
        if (testCases.isEmpty()) {
            System.err.println("[WARN] No test case directory found. Looked for directories containing: "
                    + cfg.pairs.stream().map(p -> p.left + "/" + p.right).collect(Collectors.joining(", ")));
            return 1;
        }
        System.out.println("[INFO] Found " + testCases.size() + " test cases");
        System.out.println();

        // Step 3: run the comparisons.
        for (TestCase tc : testCases) {
            System.out.println("---- Test case: " + tc.name + " (" + tc.rel + ")");
            processTestCase(tc);
        }

        // Step 4: index of every comparison.
        Path index = cfg.reportRoot.resolve("index.html");
        writeIndexHtml(index);
        System.out.println();
        System.out.println("[INFO] Index: " + index);

        // Step 5: workbook.
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
        System.out.println("==================== Summary ====================");
        System.out.println("  Test cases : " + testCases.size());
        System.out.println("  Identical  : " + same);
        System.out.println("  Different  : " + diff);
        System.out.println("  One-sided  : " + only);
        System.out.println("  Errors     : " + err);
        System.out.println("==============================================");
    }

    // ----------------------------------------------------------------------
    // Locating WinMerge
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
        // Bare command name (no separators): accept it if PATH resolves it.
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
    // Step 1: mirror the directory tree
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
                System.err.println("[WARN] Failed to walk: " + file + " (" + exc.getMessage() + ")");
                return FileVisitResult.CONTINUE;
            }
        });
        return count[0];
    }

    /**
     * Directories never walked: the report tree itself (mirroring it would recurse), VCS
     * metadata, and Windows system folders such as $RECYCLE.BIN.
     */
    private boolean isExcludedDir(Path dir) {
        try {
            if (Files.isSameFile(dir, cfg.reportRoot)) {
                return true;
            }
        } catch (IOException ignore) {
            // reportRoot may not exist yet; fall through to the path-prefix check.
        }
        if (dir.startsWith(cfg.reportRoot)) {
            return true;
        }
        String name = dir.getFileName() == null ? "" : dir.getFileName().toString();
        return name.equals(".git") || name.equals(".svn") || name.equals(".hg") || name.startsWith("$");
    }

    // ----------------------------------------------------------------------
    // Step 2: discover test cases
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
                    return FileVisitResult.SKIP_SUBTREE; // no nested test cases inside one
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

    /**
     * A directory is a test case if it holds either side of any configured pair. Only one side
     * is required so that a missing OUTPUT or LOG_COMPARE is reported as an error rather than
     * silently skipping the whole test case.
     */
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
    // Step 3: run the comparisons
    // ----------------------------------------------------------------------

    private void processTestCase(TestCase tc) throws IOException, InterruptedException {
        for (ComparePair pair : cfg.pairs) {
            Path leftDir = tc.dir.resolve(pair.left);
            Path rightDir = tc.dir.resolve(pair.right);
            Path outDir = cfg.reportRoot.resolve(tc.rel.toString()).resolve(pair.label);

            if (!Files.isDirectory(leftDir) && !Files.isDirectory(rightDir)) {
                continue; // neither side present: this pair does not apply here
            }
            Files.createDirectories(outDir);

            if (!Files.isDirectory(leftDir) || !Files.isDirectory(rightDir)) {
                CompareResult r = new CompareResult();
                r.kind = "folder compare";
                r.pairLabel = pair.label;
                r.name = "(whole folder)";
                r.left = leftDir;
                r.right = rightDir;
                r.status = Status.ERROR;
                r.message = "Missing directory: "
                        + (Files.isDirectory(leftDir) ? rightDir : leftDir);
                tc.results.add(r);
                System.out.println("     [NG] " + pair.label + " : " + r.message);
                continue;
            }

            // Folder comparison first: it is the only thing that sees files present on one side only.
            tc.results.add(runFolderCompare(pair, leftDir, rightDir, outDir));

            // Then one WinMerge run per file that exists on both sides.
            Path fileOutDir = outDir.resolve("files");
            Set<String> files;
            try {
                files = listRelativeFiles(leftDir, rightDir);
            } catch (IOException | UncheckedIOException e) {
                // An unreadable subdirectory must fail this pair only, not the whole run.
                CompareResult r = new CompareResult();
                r.kind = "file compare";
                r.pairLabel = pair.label;
                r.name = "(file listing)";
                r.left = leftDir;
                r.right = rightDir;
                r.status = Status.ERROR;
                r.message = "Cannot list files: " + e.getMessage();
                tc.results.add(r);
                System.out.println("     [NG] " + pair.label + " : " + r.message);
                continue;
            }
            for (String rel : files) {
                tc.results.add(runFileCompare(pair, leftDir, rightDir, fileOutDir, rel));
            }
        }
    }

    private CompareResult runFolderCompare(ComparePair pair, Path leftDir, Path rightDir, Path outDir)
            throws IOException, InterruptedException {
        CompareResult r = new CompareResult();
        r.kind = "folder compare";
        r.pairLabel = pair.label;
        r.name = "(whole folder)";
        r.left = leftDir;
        r.right = rightDir;
        r.report = outDir.resolve("_folder_compare.html");

        Files.createDirectories(r.report.getParent());
        List<String> cmd = buildWinMergeCommand(true, leftDir, rightDir, r.report);
        ExecResult ex = exec(cmd);
        applyExitCode(r, ex);
        System.out.println("     [" + statusMark(r.status) + "] " + pair.label + " folder compare -> "
                + reportLabel(r));
        return r;
    }

    private CompareResult runFileCompare(ComparePair pair, Path leftDir, Path rightDir, Path fileOutDir, String rel)
            throws IOException, InterruptedException {
        CompareResult r = new CompareResult();
        r.kind = "file compare";
        r.pairLabel = pair.label;
        r.name = rel;
        r.left = leftDir.resolve(rel);
        r.right = rightDir.resolve(rel);

        boolean hasLeft = Files.isRegularFile(r.left);
        boolean hasRight = Files.isRegularFile(r.right);
        if (!hasLeft || !hasRight) {
            r.status = hasLeft ? Status.LEFT_ONLY : Status.RIGHT_ONLY;
            r.message = hasLeft
                    ? "not present in " + pair.right
                    : "not present in " + pair.left;
            System.out.println("     [--] " + pair.label + " " + rel + " : " + r.message);
            return r;
        }

        r.report = shortenIfTooLong(fileOutDir, rel);
        if (!r.report.equals(fileOutDir.resolve(rel + ".html"))) {
            r.message = "report name shortened because the path was too long";
        }
        Files.createDirectories(r.report.getParent());
        List<String> cmd = buildWinMergeCommand(false, r.left, r.right, r.report);
        ExecResult ex = exec(cmd);
        applyExitCode(r, ex);
        System.out.println("     [" + statusMark(r.status) + "] " + pair.label + " " + rel + " -> "
                + reportLabel(r));
        return r;
    }

    /**
     * Keeps a report path under the Windows MAX_PATH limit of 260 characters. The report path is
     * always longer than the source path it mirrors, so deep inputs would otherwise fail to be
     * written at all - by WinMerge and by Java alike.
     *
     * <p>Two fallbacks, tried in order: flatten the subdirectories into one name, then truncate
     * the leaf and append a hash of the relative path so distinct files keep distinct names.
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

    /**
     * Progress label for one comparison. applyExitCode() clears CompareResult.report when
     * WinMerge produced no HTML, so this must tolerate null instead of relativizing it.
     */
    private String reportLabel(CompareResult r) {
        if (r.report == null) {
            return r.message.isEmpty() ? "(no report)" : "(no report: " + r.message + ")";
        }
        return cfg.reportRoot.relativize(r.report).toString();
    }

    private static String statusMark(Status s) {
        switch (s) {
            case SAME: return "OK";
            case DIFF: return "DIFF";
            case ERROR: return "NG";
            default: return "--";
        }
    }

    private void applyExitCode(CompareResult r, ExecResult ex) {
        r.exitCode = ex.exitCode;
        if (ex.timedOut) {
            r.status = Status.ERROR;
            r.message = "timed out after " + cfg.timeoutSec + "s";
            return;
        }
        if (cfg.enableExitCode) {
            // Exit codes with /enableexitcode: 0 identical, 1 different, 2 or more an error.
            switch (ex.exitCode) {
                case 0:
                    r.status = Status.SAME;
                    break;
                case 1:
                    r.status = Status.DIFF;
                    break;
                default:
                    r.status = Status.ERROR;
                    r.message = "WinMerge failed (exit=" + ex.exitCode + ") " + ex.output.trim();
                    break;
            }
        } else if (ex.exitCode != 0) {
            r.status = Status.ERROR;
            r.message = "WinMerge failed (exit=" + ex.exitCode + ") " + ex.output.trim();
        } else {
            // WinMerge older than 2.16 has no /enableexitcode and always exits 0, so the
            // verdict has to come from comparing the bytes ourselves.
            r.status = contentEquals(r.left, r.right) ? Status.SAME : Status.DIFF;
        }
        if (r.report != null && !Files.isRegularFile(r.report)) {
            if (r.status != Status.ERROR) {
                r.message = "no report file was produced";
            }
            r.report = null;
        }
    }

    /**
     * Builds the WinMerge command line. Every switch here is load-bearing:
     * <ul>
     *   <li>/r - recurse into subfolders (folder comparison only).</li>
     *   <li>/e - allow closing with Esc; /u - do not pollute the recent-files list.</li>
     *   <li>/minimize /noninteractive - start minimised and exit once the report is written.</li>
     *   <li>/enableexitcode - <b>required.</b> Without it WinMerge always exits 0 and every
     *       comparison would be recorded as identical. Added in WinMerge 2.16; for older
     *       versions the user passes --no-exitcode and contentEquals() decides instead.</li>
     *   <li>/cfg ReportFiles/ReportType=2 - emit the report as simple HTML.</li>
     *   <li>/or &lt;path&gt; - where to write it. The parent directory must already exist.</li>
     * </ul>
     *
     * <p>If a "the files are identical" dialog blocks the run, tick its do-not-show-again box
     * once, or push extra settings through --winmerge-arg. The exec() timeout is the backstop.
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
     * Byte-for-byte equality of two files, or recursively of two directories. Used only as the
     * stand-in verdict when --no-exitcode is set.
     *
     * <p>Deliberately not equivalent to WinMerge: comparison filters, whitespace and
     * line-ending options configured in WinMerge are not applied here, so a run with
     * --no-exitcode can report a difference that WinMerge itself would ignore.
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
        } catch (IOException | UncheckedIOException e) {
            // Files.walk reports an unreadable subdirectory as UncheckedIOException; treat any
            // traversal failure as "not equal" rather than aborting the whole run.
            return false;
        }
    }

    /**
     * Union of the relative file paths under both sides, sorted case-insensitively so the
     * ordering matches how Windows treats names and both sides of a pair line up.
     */
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
    // Child process execution
    // ----------------------------------------------------------------------

    static final class ExecResult {
        int exitCode = -1;
        boolean timedOut;
        String output = "";
    }

    private ExecResult exec(List<String> cmd) throws IOException, InterruptedException {
        return exec(cmd, cfg.timeoutSec);
    }

    private ExecResult exec(List<String> cmd, long timeoutSec) throws IOException, InterruptedException {
        ExecResult res = new ExecResult();
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectErrorStream(true);
        Process proc = pb.start();
        // StringBuffer, not StringBuilder: the reader thread keeps appending if join() below
        // times out (a grandchild process can hold the inherited pipe open), while this thread
        // reads it. Unsynchronized access would corrupt or truncate the diagnostics.
        StringBuffer sb = new StringBuffer();
        Thread reader = new Thread(() -> {
            try (InputStream in = proc.getInputStream()) {
                byte[] buf = new byte[4096];
                int n;
                while ((n = in.read(buf)) > 0) {
                    sb.append(new String(buf, 0, n, NATIVE_CHARSET));
                }
            } catch (IOException ignore) {
                // The pipe closing as the process exits is normal; nothing to report.
            }
        });
        reader.setDaemon(true);
        reader.start();

        if (!proc.waitFor(timeoutSec, TimeUnit.SECONDS)) {
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
    // Step 4: the index page
    // ======================================================================

    private void writeIndexHtml(Path out) throws IOException {
        StringBuilder sb = new StringBuilder();
        sb.append("<!DOCTYPE html>\n<html lang=\"ja\"><head><meta charset=\"UTF-8\">\n")
          .append("<title>Comparison reports</title>\n<style>\n")
          .append("body{font-family:'Meiryo UI',sans-serif;font-size:13px;margin:16px;}\n")
          .append("h1{font-size:18px;} h2{font-size:15px;margin-top:24px;border-left:6px solid #4472c4;padding-left:8px;}\n")
          .append("table{border-collapse:collapse;margin-bottom:8px;} th,td{border:1px solid #bfbfbf;padding:3px 8px;}\n")
          .append("th{background:#d9e1f2;} .same{color:#1f7a1f;} .diff{color:#c00000;font-weight:bold;}\n")
          .append(".err{color:#ff6600;font-weight:bold;} .only{color:#7f7f7f;}\n</style></head><body>\n")
          .append("<h1>WinMerge comparison reports</h1>\n")
          .append("<p>Generated: ").append(esc(LocalDateTime.now().format(TS_HUMAN))).append("<br>\n")
          .append("Root: ").append(esc(cfg.root.toString())).append("</p>\n");

        for (TestCase tc : testCases) {
            sb.append("<h2>").append(esc(tc.name)).append(" <span style=\"font-weight:normal;font-size:12px;color:#666\">(")
              .append(esc(tc.rel.toString())).append(")</span></h2>\n");
            sb.append("<table><tr><th>Pair</th><th>Kind</th><th>Target</th><th>Result</th><th>Report</th><th>Note</th></tr>\n");
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
    // Step 5: Excel output (com / image / xlsx)
    // ======================================================================

    private void writeExcel() {
        if ("none".equals(cfg.excelMode)) {
            System.out.println("[INFO] Excel output skipped (--excel-mode none)");
            return;
        }
        try {
            Files.createDirectories(cfg.excelPath.getParent());
        } catch (IOException e) {
            System.err.println("[ERROR] Cannot create the workbook directory: " + e.getMessage());
            return;
        }

        if ("image".equals(cfg.excelMode)) {
            try {
                boolean pasted = writeXlsxWithImages();
                System.out.println("[INFO] Workbook written ("
                        + (pasted ? "embedded images" : "text expansion") + "): " + cfg.excelPath);
            } catch (IOException | InterruptedException e) {
                System.err.println("[ERROR] Failed to build the image workbook: " + e.getMessage());
            }
            return;
        }

        boolean comCandidate = "com".equals(cfg.excelMode)
                || ("auto".equals(cfg.excelMode) && isWindows());
        if (comCandidate) {
            try {
                if (writeExcelViaCom()) {
                    System.out.println("[INFO] Workbook written (Excel COM paste): " + cfg.excelPath);
                    return;
                }
            } catch (Exception e) {
                System.err.println("[WARN] Could not run the Excel COM automation: " + e);
            }
            if ("com".equals(cfg.excelMode)) {
                System.err.println("[ERROR] Excel COM output failed. Try --excel-mode xlsx.");
                return;
            }
            System.err.println("[WARN] Falling back to direct xlsx generation.");
        }

        try {
            writeXlsx();
            System.out.println("[INFO] Workbook written (direct xlsx): " + cfg.excelPath);
        } catch (IOException e) {
            System.err.println("[ERROR] Failed to write the xlsx: " + e.getMessage());
        }
    }

    private static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("win");
    }

    // ----------------------------------------------------------------------
    // Rendering HTML to PNG with headless Edge / Chrome
    // ----------------------------------------------------------------------

    /**
     * Finds a Chromium-based browser for screenshots. Edge ships with Windows, so image mode
     * normally needs no installation; Chrome is accepted as a second choice.
     */
    private String resolveBrowser() {
        List<String> candidates = new ArrayList<>();
        if (cfg.browser != null && !cfg.browser.isEmpty()) {
            candidates.add(cfg.browser);
        }
        String env = System.getenv("BROWSER_PATH");
        if (env != null && !env.isEmpty()) {
            candidates.add(env);
        }
        String pf = System.getenv("ProgramFiles");
        String pf86 = System.getenv("ProgramFiles(x86)");
        String local = System.getenv("LOCALAPPDATA");
        for (String root : new String[]{pf86, pf, local}) {
            if (root == null) {
                continue;
            }
            candidates.add(root + "\\Microsoft\\Edge\\Application\\msedge.exe");
            candidates.add(root + "\\Google\\Chrome\\Application\\chrome.exe");
        }
        // Non-Windows paths, used when testing this tool outside Windows.
        candidates.add("/opt/pw-browsers/chromium/chrome-linux/chrome");
        candidates.add("/usr/bin/chromium");
        candidates.add("/usr/bin/google-chrome");
        for (String c : candidates) {
            if (Files.isRegularFile(Paths.get(c))) {
                return Paths.get(c).toAbsolutePath().toString();
            }
        }
        return null;
    }

    /**
     * Screenshots one HTML report to PNG.
     *
     * <p>A headless capture only covers the window, not the whole document, so the window
     * height is estimated from the number of table rows in the report. Over-estimating just
     * adds white space; under-estimating clips the bottom. The top is never lost, and the
     * sheet always links to the original HTML as the fallback for a clipped capture.
     */
    private boolean renderHtmlToPng(String browser, Path html, Path png) throws IOException, InterruptedException {
        int rows = extractHtmlRows(html, cfg.imageMaxHeight / 18).size();
        // 24px per row, deliberately generous, to absorb rows that wrap onto two lines.
        int height = Math.max(200, Math.min(cfg.imageMaxHeight, 200 + rows * 24));
        // Keep the throwaway browser profile out of the report tree.
        Path profile = Paths.get(System.getProperty("java.io.tmpdir"), "winmerge_report_browser");
        Files.createDirectories(png.getParent());
        // Drop any PNG from an earlier run: without --clean a failed capture would otherwise
        // look like a success and the stale image would be embedded again.
        Files.deleteIfExists(png);

        List<String> cmd = new ArrayList<>(List.of(
                browser,
                "--headless=new",
                "--disable-gpu",
                "--hide-scrollbars",
                "--no-first-run",
                "--no-default-browser-check",
                "--default-background-color=FFFFFFFF",
                "--virtual-time-budget=3000",
                "--user-data-dir=" + profile,
                "--window-size=" + cfg.imageWidth + "," + height,
                "--screenshot=" + png,
                html.toUri().toString()));
        if (!isWindows()) {
            cmd.add(3, "--no-sandbox");
        }
        ExecResult ex = exec(cmd, cfg.timeoutSec);
        if (!Files.isRegularFile(png) || Files.size(png) == 0 || pngSize(png) == null) {
            System.err.println("[WARN] Screenshot failed: " + html.getFileName()
                    + " (exit=" + ex.exitCode + ") " + ex.output.trim());
            return false;
        }
        return true;
    }

    /**
     * Reads width and height from the PNG IHDR chunk, at fixed offsets 16 and 20. Returns null
     * for anything that is not a PNG, which also serves as the "did the browser really write an
     * image" check in renderHtmlToPng().
     */
    private static int[] pngSize(Path png) {
        try {
            byte[] b = Files.readAllBytes(png);
            if (b.length < 24 || (b[0] & 0xFF) != 0x89 || b[1] != 'P' || b[2] != 'N' || b[3] != 'G') {
                return null;
            }
            int w = ((b[16] & 0xFF) << 24) | ((b[17] & 0xFF) << 16) | ((b[18] & 0xFF) << 8) | (b[19] & 0xFF);
            int h = ((b[20] & 0xFF) << 24) | ((b[21] & 0xFF) << 16) | ((b[22] & 0xFF) << 8) | (b[23] & 0xFF);
            return (w > 0 && h > 0) ? new int[]{w, h} : null;
        } catch (IOException e) {
            return null;
        }
    }

    // ----------------------------------------------------------------------
    // Excel COM path: paste the HTML into sheets by driving Excel from VBScript
    // ----------------------------------------------------------------------

    private boolean writeExcelViaCom() throws IOException, InterruptedException {
        Path vbs = cfg.reportRoot.resolve("_build_excel.vbs");
        Path log = cfg.reportRoot.resolve("_build_excel.log");
        Files.deleteIfExists(log);

        Files.writeString(vbs, buildVbScript(log), StandardCharsets.UTF_16LE);
        // cscript only treats a .vbs as Unicode when it starts with a UTF-16LE BOM. Without
        // this the Japanese sheet names and paths in the script would be misread.
        byte[] body = Files.readAllBytes(vbs);
        byte[] withBom = new byte[body.length + 2];
        withBom[0] = (byte) 0xFF;
        withBom[1] = (byte) 0xFE;
        System.arraycopy(body, 0, withBom, 2, body.length);
        Files.write(vbs, withBom);

        // One cscript run builds the whole workbook, so the per-comparison timeout is far too
        // short. Killing it midway also orphans an invisible EXCEL.EXE (xl.Quit never runs),
        // so allow generous time and scale it with the number of reports to paste.
        long comTimeout = Math.max(cfg.timeoutSec, 60L + 20L * countReports());
        ExecResult ex = exec(List.of("cscript", "//nologo", "//B", vbs.toString()), comTimeout);
        boolean created = Files.isRegularFile(cfg.excelPath);
        if (created && !ex.timedOut && ex.exitCode == 0) {
            if (!ex.output.isBlank()) {
                System.err.println("[WARN] Excel COM warning: " + ex.output.trim());
            }
            return true;
        }

        // Always say why it failed. This used to return false silently, which hid the real
        // cause (Excel not installed) behind a bare fallback message. The script collects its
        // own diagnostics because VBScript clears Err when a Sub returns - see buildVbScript.
        System.err.println("[WARN] Excel COM automation failed"
                + (ex.timedOut ? " (timed out)" : " (cscript exit=" + ex.exitCode + ")")
                + (created ? "" : " / no workbook produced"));
        if (!ex.output.isBlank()) {
            for (String line : ex.output.trim().split("\\R")) {
                System.err.println("       " + line);
            }
        }
        System.err.println("       script: " + vbs);
        if (Files.isRegularFile(log)) {
            System.err.println("       log   : " + log);
        }
        return false;
    }

    /**
     * Emits the VBScript that drives Excel.
     *
     * <p>Error handling is the subtle part. VBScript resets the Err object when a Sub returns,
     * so testing Err.Number at the end of the script cannot see a failure that happened inside
     * AddSheet or AddReport - the script would exit 0 having produced nothing. Instead every
     * operation calls Report to copy the error into an accumulator, and success is decided by
     * whether the workbook file actually exists. Diagnostics go to stderr and to a log file.
     *
     * <p>When editing the emitted script, keep to that discipline: no bare operation without a
     * following Report, and never reintroduce a final Err.Number check as the success test.
     */
    private int countReports() {
        int n = 0;
        for (TestCase tc : testCases) {
            for (CompareResult r : tc.results) {
                if (r.report != null) {
                    n++;
                }
            }
        }
        return n;
    }

    private String buildVbScript(Path logFile) {
        StringBuilder sb = new StringBuilder();
        sb.append("Option Explicit\r\n")
          .append("Dim xl, wb, ws, fso, gRow, gFirst, gErrors, gOut, gLog\r\n")
          .append("gErrors = \"\"\r\n")
          .append("gFirst = True\r\n")
          .append("gRow = 1\r\n")
          .append("gOut = ").append(vbsStr(cfg.excelPath.toString())).append("\r\n")
          .append("gLog = ").append(vbsStr(logFile.toString())).append("\r\n")
          .append("Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n")
          .append("\r\n")
          .append("On Error Resume Next\r\n")
          .append("Set xl = CreateObject(\"Excel.Application\")\r\n")
          .append("If Err.Number <> 0 Then\r\n")
          .append("  WScript.StdErr.WriteLine \"Cannot start Excel (probably not installed): \" ")
          .append("& Err.Number & \" \" & Err.Description\r\n")
          .append("  WScript.Quit 2\r\n")
          .append("End If\r\n")
          .append("On Error GoTo 0\r\n")
          .append("\r\n")
          .append("xl.Visible = False\r\n")
          .append("xl.DisplayAlerts = False\r\n")
          .append("Set wb = xl.Workbooks.Add\r\n")
          .append("On Error Resume Next\r\n")
          .append("Do While wb.Worksheets.Count > 1\r\n")
          .append("  wb.Worksheets(wb.Worksheets.Count).Delete\r\n")
          .append("Loop\r\n")
          .append("Report \"init\"\r\n");

        // --- Subroutine definitions (VBScript hoists these, so call sites may precede them) ---
        sb.append("\r\n")
          .append("Sub Report(where)\r\n")
          .append("  If Err.Number <> 0 Then\r\n")
          .append("    gErrors = gErrors & where & \" : \" & Err.Number & \" \" & Err.Description & vbCrLf\r\n")
          .append("    Err.Clear\r\n")
          .append("  End If\r\n")
          .append("End Sub\r\n\r\n")
          .append("Sub AddSheet(sname)\r\n")
          .append("  On Error Resume Next\r\n")
          .append("  If gFirst Then\r\n")
          .append("    Set ws = wb.Worksheets(1)\r\n")
          .append("    gFirst = False\r\n")
          .append("  Else\r\n")
          .append("    wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count)\r\n")
          .append("    Set ws = wb.Worksheets(wb.Worksheets.Count)\r\n")
          .append("  End If\r\n")
          .append("  Report \"add sheet \" & sname\r\n")
          .append("  ws.Name = sname\r\n")
          .append("  Report \"name sheet \" & sname\r\n")
          .append("  gRow = 1\r\n")
          .append("End Sub\r\n\r\n")
          .append("Sub AddLine(txt, isBold)\r\n")
          .append("  On Error Resume Next\r\n")
          .append("  ws.Cells(gRow, 1).Value = txt\r\n")
          .append("  ws.Cells(gRow, 1).Font.Bold = isBold\r\n")
          .append("  Report \"write row\"\r\n")
          .append("  gRow = gRow + 1\r\n")
          .append("End Sub\r\n\r\n")
          .append("Sub AddRow4(c1, c2, c3, c4)\r\n")
          .append("  On Error Resume Next\r\n")
          .append("  ws.Cells(gRow, 1).Value = c1\r\n")
          .append("  ws.Cells(gRow, 2).Value = c2\r\n")
          .append("  ws.Cells(gRow, 3).Value = c3\r\n")
          .append("  ws.Cells(gRow, 4).Value = c4\r\n")
          .append("  Report \"write row\"\r\n")
          .append("  gRow = gRow + 1\r\n")
          .append("End Sub\r\n\r\n")
          .append("Sub AddReport(title, htmlPath)\r\n")
          .append("  Dim src, used\r\n")
          .append("  On Error Resume Next\r\n")
          .append("  AddLine title, True\r\n")
          .append("  If Not fso.FileExists(htmlPath) Then\r\n")
          .append("    AddLine \"  report not found: \" & htmlPath, False\r\n")
          .append("    gRow = gRow + 1\r\n")
          .append("    Exit Sub\r\n")
          .append("  End If\r\n")
          .append("  ws.Hyperlinks.Add ws.Cells(gRow, 1), htmlPath, , , htmlPath\r\n")
          .append("  Report \"add link \" & htmlPath\r\n")
          .append("  gRow = gRow + 1\r\n")
          .append("  Set src = xl.Workbooks.Open(htmlPath, False, True)\r\n")
          .append("  If Err.Number <> 0 Then\r\n")
          .append("    Report \"cannot open HTML \" & htmlPath\r\n")
          .append("    AddLine \"  failed to import the HTML\", False\r\n")
          .append("    gRow = gRow + 1\r\n")
          .append("    Exit Sub\r\n")
          .append("  End If\r\n")
          .append("  Set used = src.Worksheets(1).UsedRange\r\n")
          // Range.Copy(Destination) instead of the clipboard: Worksheet.Paste requires the
          // target sheet to be active and leaves CutCopyMode set, which is fragile in an
          // invisible Excel instance.
          .append("  used.Copy ws.Cells(gRow, 1)\r\n")
          .append("  Report \"paste \" & htmlPath\r\n")
          .append("  gRow = gRow + used.Rows.Count + 2\r\n")
          .append("  src.Close False\r\n")
          .append("  Report \"close HTML \" & htmlPath\r\n")
          .append("End Sub\r\n\r\n");

        // --- Data section: one call per sheet and per report ---
        Set<String> usedNames = new LinkedHashSet<>();

        // Summary sheet, always first.
        sb.append(call("AddSheet", uniqueSheetName("Summary", usedNames)));
        sb.append(call("AddLine", "WinMerge comparison report - summary", "True"));
        sb.append(call("AddLine", "Generated: " + LocalDateTime.now().format(TS_HUMAN), "False"));
        sb.append(call("AddLine", "Root: " + cfg.root, "False"));
        sb.append("gRow = gRow + 1\r\n");
        sb.append(call("AddRow4", "Test case", "Pair", "Target", "Result"));
        for (TestCase tc : testCases) {
            for (CompareResult r : tc.results) {
                sb.append(call("AddRow4", tc.name, r.pairLabel, r.kind + " " + r.name,
                        r.status.label + (r.message.isEmpty() ? "" : " / " + r.message)));
            }
        }
        sb.append("ws.Columns(\"A:D\").AutoFit\r\n")
          .append("Report \"autofit summary\"\r\n");

        // One sheet per test case.
        for (TestCase tc : testCases) {
            sb.append(call("AddSheet", uniqueSheetName(tc.name, usedNames)));
            sb.append(call("AddLine", "Test case: " + tc.name, "True"));
            sb.append(call("AddLine", "Path: " + tc.dir, "False"));
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
        }

        // --- Save, then decide success ---
        sb.append("\r\n")
          .append("wb.Worksheets(1).Activate\r\n")
          .append("Report \"select first sheet\"\r\n")
          .append("wb.SaveAs gOut, 51\r\n")
          .append("Report \"save workbook\"\r\n")
          .append("wb.Close False\r\n")
          .append("Report \"close workbook\"\r\n")
          .append("xl.Quit\r\n")
          .append("Report \"quit Excel\"\r\n")
          .append("\r\n")
          .append("Dim logStream\r\n")
          .append("Set logStream = fso.CreateTextFile(gLog, True, True)\r\n")
          .append("If Err.Number = 0 Then\r\n")
          .append("  logStream.WriteLine \"target: \" & gOut\r\n")
          .append("  logStream.WriteLine \"saved: \" & CStr(fso.FileExists(gOut))\r\n")
          .append("  logStream.WriteLine gErrors\r\n")
          .append("  logStream.Close\r\n")
          .append("End If\r\n")
          .append("Err.Clear\r\n")
          .append("\r\n")
          // Err is unreliable here (cleared on Sub exit), so trust the file system.
          .append("If fso.FileExists(gOut) Then\r\n")
          .append("  If Len(gErrors) > 0 Then WScript.StdErr.WriteLine gErrors\r\n")
          .append("  WScript.Quit 0\r\n")
          .append("Else\r\n")
          .append("  WScript.StdErr.WriteLine \"the workbook was not saved\" & vbCrLf & gErrors\r\n")
          .append("  WScript.Quit 1\r\n")
          .append("End If\r\n");
        return sb.toString();
    }

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
    // Hand-built .xlsx (OOXML via java.util.zip, no third-party library)
    // ----------------------------------------------------------------------

    // Cell style ids. These are indexes into the cellXfs list in stylesXml(); adding a style
    // means appending an <xf> there and a constant here, in the same order.
    private static final int ST_NORMAL = 0;
    private static final int ST_BOLD = 1;
    private static final int ST_TITLE = 2;
    private static final int ST_LINK = 3;
    private static final int ST_HEADER = 4;
    private static final int ST_CELL = 5;

    static final class Cell {
        final String text;
        final int style;
        final String hyperlink; // external link target, or null

        Cell(String text, int style, String hyperlink) {
            this.text = text;
            this.style = style;
            this.hyperlink = hyperlink;
        }
    }

    /** One embedded picture, anchored to a single cell and drawn at its natural pixel size. */
    static final class ImageRef {
        final Path png;
        final int anchorRow;   // zero-based row the picture is anchored to
        final int widthPx;
        final int heightPx;
        final String borderRgb;   // RRGGBB outline, or empty for no outline
        final int borderWidthPt;

        ImageRef(Path png, int anchorRow, int widthPx, int heightPx, String borderRgb, int borderWidthPt) {
            this.png = png;
            this.anchorRow = anchorRow;
            this.widthPx = widthPx;
            this.heightPx = heightPx;
            this.borderRgb = borderRgb;
            this.borderWidthPt = borderWidthPt;
        }
    }

    static final class SheetData {
        final String name;
        final List<List<Cell>> rows = new ArrayList<>();
        final List<ImageRef> images = new ArrayList<>();
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

    /**
     * Builds the workbook with each report embedded as a picture. Needs no Excel: the drawing,
     * media and relationship parts are written straight into the package.
     *
     * @return true if pictures were embedded, false if it fell back to the text-only workbook
     *         because no browser was found. The caller reports which one actually happened.
     */
    private boolean writeXlsxWithImages() throws IOException, InterruptedException {
        String browser = resolveBrowser();
        if (browser == null) {
            System.err.println("[WARN] No browser (Edge / Chrome) found for PNG capture.");
            System.err.println("       Pass --browser with the path to msedge.exe. Falling back to text expansion.");
            writeXlsx();
            return false;
        }
        System.out.println("[INFO] Browser used for capture: " + browser);

        Path imageDir = cfg.reportRoot.resolve("_images");
        List<SheetData> sheets = new ArrayList<>();
        Set<String> usedNames = new LinkedHashSet<>();
        sheets.add(buildSummarySheet(usedNames));

        int done = 0;
        int failed = 0;
        for (TestCase tc : testCases) {
            SheetData sd = new SheetData(uniqueSheetName(tc.name, usedNames));
            sd.colWidths = new double[]{60, 18, 18, 18, 18, 18, 18, 18};
            sd.addText("Test case: " + tc.name, ST_TITLE);
            sd.addText("Path: " + tc.dir, ST_NORMAL);
            sd.blank();

            for (CompareResult r : tc.results) {
                sd.addText("[" + r.pairLabel + "] " + r.kind + " : " + r.name
                        + "  => " + r.status.label + (r.message.isEmpty() ? "" : " (" + r.message + ")"), ST_BOLD);
                if (r.report == null) {
                    sd.blank();
                    continue;
                }
                String rel = cfg.reportRoot.relativize(r.report).toString().replace('\\', '/');
                sd.add(new Cell("Open report: " + rel, ST_LINK, hyperlinkTarget(r.report)));

                Path png = imageDir.resolve(rel + ".png");
                if (renderHtmlToPng(browser, r.report, png)) {
                    int[] size = pngSize(png);
                    int w = size == null ? cfg.imageWidth : size[0];
                    int h = size == null ? 600 : size[1];
                    sd.images.add(new ImageRef(png, sd.rows.size(), w, h,
                            cfg.imageBorder, cfg.imageBorderWidthPt));
                    // Advance past the picture. A one-cell anchor does not reserve space, so
                    // without these blank rows the next heading would sit under the image.
                    int rowsNeeded = (int) Math.ceil(h / 20.0);
                    for (int i = 0; i < rowsNeeded; i++) {
                        sd.blank();
                    }
                    done++;
                } else {
                    sd.addText("(screenshot failed; open the HTML through the link above)", ST_NORMAL);
                    failed++;
                }
                sd.blank();
            }
            sheets.add(sd);
            System.out.println("     captured: " + tc.name);
        }

        writeXlsxFile(cfg.excelPath, sheets);
        System.out.println("[INFO] Embedded " + done + " images"
                + (failed > 0 ? " (" + failed + " failed)" : "") + " / PNGs kept in: " + imageDir);
        return true;
    }

    /** The summary sheet is identical in image and text mode. */
    private SheetData buildSummarySheet(Set<String> usedNames) {
        SheetData summary = new SheetData(uniqueSheetName("Summary", usedNames));
        summary.colWidths = new double[]{24, 22, 14, 44, 12, 40, 40};
        summary.addText("WinMerge comparison report - summary", ST_TITLE);
        summary.addText("Generated: " + LocalDateTime.now().format(TS_HUMAN), ST_NORMAL);
        summary.addText("Root: " + cfg.root, ST_NORMAL);
        summary.blank();
        summary.add(header("Test case"), header("Pair"), header("Kind"), header("Target"),
                header("Result"), header("Note"), header("Report"));
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
        return summary;
    }

    private void writeXlsx() throws IOException {
        List<SheetData> sheets = new ArrayList<>();
        Set<String> usedNames = new LinkedHashSet<>();

        // --- Summary sheet ---
        SheetData summary = new SheetData(uniqueSheetName("Summary", usedNames));
        summary.colWidths = new double[]{24, 22, 14, 44, 12, 40, 40};
        summary.addText("WinMerge comparison report - summary", ST_TITLE);
        summary.addText("Generated: " + LocalDateTime.now().format(TS_HUMAN), ST_NORMAL);
        summary.addText("Root: " + cfg.root, ST_NORMAL);
        summary.blank();
        summary.add(header("Test case"), header("Pair"), header("Kind"), header("Target"),
                header("Result"), header("Note"), header("Report"));
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

        // --- One sheet per test case ---
        for (TestCase tc : testCases) {
            SheetData sd = new SheetData(uniqueSheetName(tc.name, usedNames));
            sd.addText("Test case: " + tc.name, ST_TITLE);
            sd.addText("Path: " + tc.dir, ST_NORMAL);
            sd.blank();
            for (CompareResult r : tc.results) {
                sd.addText("[" + r.pairLabel + "] " + r.kind + " : " + r.name
                        + "  => " + r.status.label + (r.message.isEmpty() ? "" : " (" + r.message + ")"), ST_BOLD);
                if (r.report != null) {
                    sd.add(new Cell("Open report: " + cfg.reportRoot.relativize(r.report).toString().replace('\\', '/'),
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
                        sd.addText("(could not parse the report; open the HTML through the link above)", ST_NORMAL);
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

    /**
     * Writes the .xlsx package.
     *
     * <p>The parts must stay mutually consistent: every sheet needs an Override in
     * [Content_Types].xml, a relationship from workbook.xml.rels, and - when it has hyperlinks
     * or pictures - its own sheetN.xml.rels. Relationship ids are positional: hyperlinks take
     * rId1..rIdN in the order sheetXml() emits them, and the drawing takes rId(N+1). Changing
     * that numbering in one place without the other produces a file Excel refuses to open.
     */
    private static void writeXlsxFile(Path out, List<SheetData> sheets) throws IOException {
        try (ZipOutputStream zip = new ZipOutputStream(Files.newOutputStream(out), StandardCharsets.UTF_8)) {
            // [Content_Types].xml
            StringBuilder ct = new StringBuilder();
            ct.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
              .append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">")
              .append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>")
              .append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>")
              .append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>")
              .append("<Default Extension=\"png\" ContentType=\"image/png\"/>")
              .append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 1; i <= sheets.size(); i++) {
                ct.append("<Override PartName=\"/xl/worksheets/sheet").append(i)
                  .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                if (!sheets.get(i - 1).images.isEmpty()) {
                    ct.append("<Override PartName=\"/xl/drawings/drawing").append(i)
                      .append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
                }
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

            int mediaSeq = 0;
            for (int i = 0; i < sheets.size(); i++) {
                SheetData sd = sheets.get(i);
                int sheetNo = i + 1;
                List<String> links = new ArrayList<>();
                put(zip, "xl/worksheets/sheet" + sheetNo + ".xml", sheetXml(sd, links));

                if (!links.isEmpty() || !sd.images.isEmpty()) {
                    StringBuilder sr = new StringBuilder();
                    sr.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
                      .append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    for (int k = 0; k < links.size(); k++) {
                        sr.append("<Relationship Id=\"rId").append(k + 1)
                          .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"")
                          .append(esc(links.get(k))).append("\" TargetMode=\"External\"/>");
                    }
                    if (!sd.images.isEmpty()) {
                        // Drawing id follows the hyperlink ids; sheetXml() computes the same
                        // number independently, so the two must be kept in step.
                        sr.append("<Relationship Id=\"rId").append(links.size() + 1)
                          .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing")
                          .append(sheetNo).append(".xml\"/>");
                    }
                    sr.append("</Relationships>");
                    put(zip, "xl/worksheets/_rels/sheet" + sheetNo + ".xml.rels", sr.toString());
                }

                if (!sd.images.isEmpty()) {
                    StringBuilder dr = new StringBuilder();
                    dr.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
                      .append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    List<String> mediaNames = new ArrayList<>();
                    for (int k = 0; k < sd.images.size(); k++) {
                        String media = "image" + (++mediaSeq) + ".png";
                        mediaNames.add(media);
                        putBinary(zip, "xl/media/" + media, Files.readAllBytes(sd.images.get(k).png));
                        dr.append("<Relationship Id=\"rId").append(k + 1)
                          .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/")
                          .append(media).append("\"/>");
                    }
                    dr.append("</Relationships>");
                    put(zip, "xl/drawings/_rels/drawing" + sheetNo + ".xml.rels", dr.toString());
                    put(zip, "xl/drawings/drawing" + sheetNo + ".xml", drawingXml(sd.images));
                }
            }
        }
    }

    /** OOXML measures drawings in EMU (English Metric Units); 1 pixel is 9525 EMU. */
    private static final int EMU_PER_PX = 9525;

    private static String drawingXml(List<ImageRef> images) {
        StringBuilder sb = new StringBuilder();
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
          .append("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ")
          .append("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
        for (int k = 0; k < images.size(); k++) {
            ImageRef img = images.get(k);
            long cx = (long) img.widthPx * EMU_PER_PX;
            long cy = (long) img.heightPx * EMU_PER_PX;
            sb.append("<xdr:oneCellAnchor>")
              .append("<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>")
              .append(img.anchorRow).append("</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>")
              .append("<xdr:ext cx=\"").append(cx).append("\" cy=\"").append(cy).append("\"/>")
              .append("<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"").append(k + 1)
              .append("\" name=\"Report ").append(k + 1).append("\"/><xdr:cNvPicPr/></xdr:nvPicPr>")
              .append("<xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId")
              .append(k + 1).append("\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>")
              .append("<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"").append(cx)
              .append("\" cy=\"").append(cy).append("\"/></a:xfrm>")
              .append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            if (img.borderRgb != null && !img.borderRgb.isEmpty()) {
                // Outline. In CT_ShapeProperties <a:ln> must follow the geometry, and its width
                // is in EMU: 1pt is 12700 EMU.
                sb.append("<a:ln w=\"").append(Math.max(1, img.borderWidthPt) * 12700)
                  .append("\"><a:solidFill><a:srgbClr val=\"").append(img.borderRgb)
                  .append("\"/></a:solidFill></a:ln>");
            }
            sb.append("</xdr:spPr>")
              .append("</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>");
        }
        sb.append("</xdr:wsDr>");
        return sb.toString();
    }

    private static void putBinary(ZipOutputStream zip, String name, byte[] content) throws IOException {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(content);
        zip.closeEntry();
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
        if (!sd.images.isEmpty()) {
            // Must match the id writeXlsxFile() assigns to the drawing relationship.
            sb.append("<drawing r:id=\"rId").append(outLinks.size() + 1).append("\"/>");
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

    /**
     * Makes a string safe for a cell: Excel caps a cell at 32767 characters, and control
     * characters below 0x20 are not representable in XML at all.
     */
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
                // Control characters cannot appear in XML; drop them.
                continue;
            } else {
                sb.append(ch);
            }
        }
        return sb.toString().trim();
    }

    // ======================================================================
    // Parsing the WinMerge HTML report (for the text and image modes)
    // ======================================================================

    private static final Pattern P_TR = Pattern.compile("(?is)<tr\\b[^>]*>(.*?)</tr>");
    private static final Pattern P_TD = Pattern.compile("(?is)<t([dh])\\b[^>]*>(.*?)</t\\1>");
    private static final Pattern P_TAG = Pattern.compile("(?s)<[^>]*>");
    private static final Pattern P_META_CHARSET =
            Pattern.compile("(?i)<meta[^>]*charset\\s*=\\s*[\"']?\\s*([A-Za-z0-9_\\-]+)");

    /**
     * Extracts the report as rows of cell text. Regex rather than a parser because WinMerge
     * emits a small, predictable table and this file takes no dependencies. If no table is
     * found, falls back to the body text one line per row, so the caller always gets something.
     * Also used by image mode purely to count rows for the capture height.
     */
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
        // No table: strip the markup and keep the visible lines.
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
    // Utilities
    // ======================================================================

    /**
     * Coerces a test case name into a legal Excel sheet name: at most 31 characters, none of
     * \ / * ? [ ] :, not starting with an apostrophe, and unique within the workbook. The
     * `used` set is compared case-insensitively because Excel treats sheet names that way.
     */
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

    /**
     * Percent-encodes a relative path for use as a hyperlink target, keeping '/' as the
     * separator. Japanese file names become UTF-8 percent escapes, which both Excel and
     * browsers resolve correctly.
     */
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
}
