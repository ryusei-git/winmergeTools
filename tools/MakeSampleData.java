import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Comparator;
import java.util.stream.Stream;

/**
 * Generates the sample tree used to verify WinMergeReportTool on a fresh machine.
 *
 * <p>Each test case exercises one behaviour of the tool, so the console summary can be checked
 * against a known expectation (11 identical / 11 different / 3 one-sided / 1 error):
 * <ul>
 *   <li>TC001 - both pairs identical; everything should report as identical.</li>
 *   <li>TC002 - content differences, including a Shift_JIS file, a UTF-8 file with a Japanese
 *       name, and a file in a subdirectory, to cover encoding and recursion.</li>
 *   <li>TC003 - a file on one side only, plus an empty LOG_COMPARE directory.</li>
 *   <li>TC004 - LOG_COMPARE missing entirely; that pair must be recorded as an error.</li>
 *   <li>nested/TC005 - a test case one level down, to prove nested discovery works.</li>
 * </ul>
 *
 * <p>Usage:
 * <pre>
 *   javac -encoding UTF-8 -d %TEMP%\wmrt tools\MakeSampleData.java
 *   java -cp %TEMP%\wmrt MakeSampleData --dest D:\tests\sample --clean
 * </pre>
 *
 * <p>Comments are English because the source is UTF-8 and JDK 17 and older default the source
 * encoding to the platform charset; see WinMergeReportTool for the full rationale.
 */
public final class MakeSampleData {

    /** Sample data is written with CRLF, matching what Windows tools normally produce. */
    private static final String NL = "\r\n";

    private static final Charset SJIS = charset("windows-31j");

    public static void main(String[] args) throws IOException {
        Path dest = Paths.get("sample").toAbsolutePath().normalize();
        boolean clean = false;
        for (int i = 0; i < args.length; i++) {
            switch (args[i]) {
                case "--dest":
                    if (++i >= args.length) {
                        System.err.println("[ERROR] --dest には値が必要です");
                        System.exit(2);
                    }
                    dest = Paths.get(args[i]).toAbsolutePath().normalize();
                    break;
                case "--clean":
                    clean = true;
                    break;
                case "--help":
                case "-h":
                    System.out.println("使い方: java MakeSampleData [--dest <dir>] [--clean]");
                    return;
                default:
                    System.err.println("[ERROR] 不明なオプション: " + args[i]);
                    System.exit(2);
            }
        }

        if (clean && Files.exists(dest)) {
            System.out.println("[INFO] 既存のサンプルを削除します: " + dest);
            deleteRecursively(dest);
        }
        Files.createDirectories(dest);

        makeIdentical(dest.resolve("TC001_一致"));
        makeDifferent(dest.resolve("TC002_差分あり"));
        makeOneSided(dest.resolve("TC003_片側のみ"));
        makeMissingPairDir(dest.resolve("TC004_LOG_COMPAREなし"));
        makeIdentical(dest.resolve("nested").resolve("TC005_入れ子"));
        // Introduce one difference in the nested case so its detection is visible in the report.
        write(dest.resolve("nested/TC005_入れ子/OUTPUT/data.csv"),
                csv("3,gamma"), StandardCharsets.UTF_8);

        System.out.println("[INFO] サンプルデータを作成しました: " + dest);
        System.out.println();
        System.out.println("  TC001_一致             全て一致になるはず");
        System.out.println("  TC002_差分あり         data.csv / 日本語ファイル名.txt / sjis_log.txt / app.log が差分");
        System.out.println("  TC003_片側のみ         only_in_input.txt=左のみ, only_in_output.txt=右のみ, LOG は左のみ");
        System.out.println("  TC004_LOG_COMPAREなし  LOG ペアがエラーとして記録されるはず");
        System.out.println("  nested/TC005_入れ子    入れ子でも検出され、data.csv が差分になるはず");
    }

    /** Test case where both pairs match exactly. */
    private static void makeIdentical(Path tc) throws IOException {
        for (String dir : new String[]{"INPUT", "OUTPUT"}) {
            write(tc.resolve(dir).resolve("data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
            write(tc.resolve(dir).resolve("sub/note.txt"), "サブフォルダ内のファイルです。" + NL, StandardCharsets.UTF_8);
        }
        for (String dir : new String[]{"LOG", "LOG_COMPARE"}) {
            write(tc.resolve(dir).resolve("app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        }
    }

    /** Test case with content differences, mixing encodings and file name shapes. */
    private static void makeDifferent(Path tc) throws IOException {
        write(tc.resolve("INPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/data.csv"), csv("2,BETA"), StandardCharsets.UTF_8);

        write(tc.resolve("INPUT/日本語ファイル名.txt"),
                "1行目：変更なし" + NL + "2行目：変更前" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/日本語ファイル名.txt"),
                "1行目：変更なし" + NL + "2行目：変更後" + NL, StandardCharsets.UTF_8);

        // Shift_JIS file: checks that WinMerge detects the encoding rather than us.
        write(tc.resolve("INPUT/sjis_log.txt"), log("シフトJISのログです"), SJIS);
        write(tc.resolve("OUTPUT/sjis_log.txt"), log("シフトJISのログです（変更後）"), SJIS);

        write(tc.resolve("INPUT/sub/note.txt"), "サブフォルダ内のファイルです。" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/sub/note.txt"), "サブフォルダ内のファイルです（変更後）。" + NL, StandardCharsets.UTF_8);

        write(tc.resolve("LOG/app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        write(tc.resolve("LOG_COMPARE/app.log"), log("処理を終了しました（想定値）"), StandardCharsets.UTF_8);
    }

    /** Test case with one-sided files and an empty LOG_COMPARE directory. */
    private static void makeOneSided(Path tc) throws IOException {
        write(tc.resolve("INPUT/common.txt"), "両方にあるファイル" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/common.txt"), "両方にあるファイル" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("INPUT/only_in_input.txt"), "INPUT にしかないファイル" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/only_in_output.txt"), "OUTPUT にしかないファイル" + NL, StandardCharsets.UTF_8);

        write(tc.resolve("LOG/app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        Files.createDirectories(tc.resolve("LOG_COMPARE")); // intentionally left empty
    }

    /** Test case missing the LOG_COMPARE directory, which must surface as an error. */
    private static void makeMissingPairDir(Path tc) throws IOException {
        write(tc.resolve("INPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("LOG/app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        // LOG_COMPARE is deliberately not created.
    }

    private static String csv(String lastRow) {
        return "id,name" + NL + "1,alpha" + NL + lastRow + NL;
    }

    private static String log(String lastLine) {
        return "2026-08-31 10:00:00 INFO  処理を開始しました" + NL
                + "2026-08-31 10:00:01 DEBUG レコード件数: 2" + NL
                + "2026-08-31 10:00:02 INFO  " + lastLine + NL;
    }

    private static void write(Path file, String content, Charset cs) throws IOException {
        Files.createDirectories(file.getParent());
        Files.write(file, content.getBytes(cs));
    }

    private static Charset charset(String name) {
        try {
            return Charset.forName(name);
        } catch (RuntimeException e) {
            return StandardCharsets.UTF_8;
        }
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
