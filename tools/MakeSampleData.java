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
 * <p>This file is pure ASCII, so it compiles identically whatever -encoding is in effect.
 * The Japanese file name and Shift_JIS contents in TC002 are the point of that test case -
 * the tool has to handle both - so they are kept as \\uXXXX escapes rather than removed.
 * Decode them with any Unicode table if you need to read them; do not paste raw Japanese back
 * in. See WinMergeReportTool for the full encoding rationale.
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
                        System.err.println("[ERROR] --dest requires a value");
                        System.exit(2);
                    }
                    dest = Paths.get(args[i]).toAbsolutePath().normalize();
                    break;
                case "--clean":
                    clean = true;
                    break;
                case "--help":
                case "-h":
                    System.out.println("Usage: java MakeSampleData [--dest <dir>] [--clean]");
                    return;
                default:
                    System.err.println("[ERROR] Unknown option: " + args[i]);
                    System.exit(2);
            }
        }

        if (clean && Files.exists(dest)) {
            System.out.println("[INFO] Deleting the existing sample: " + dest);
            deleteRecursively(dest);
        }
        Files.createDirectories(dest);

        makeIdentical(dest.resolve("TC001_identical"));
        makeDifferent(dest.resolve("TC002_different"));
        makeOneSided(dest.resolve("TC003_one_sided"));
        makeMissingPairDir(dest.resolve("TC004_no_log_compare"));
        makeIdentical(dest.resolve("nested").resolve("TC005_nested"));
        // Introduce one difference in the nested case so its detection is visible in the report.
        write(dest.resolve("nested/TC005_nested/OUTPUT/data.csv"),
                csv("3,gamma"), StandardCharsets.UTF_8);

        System.out.println("[INFO] Sample data created: " + dest);
        System.out.println();
        System.out.println("  TC001_identical       everything identical");
        System.out.println("  TC002_different       data.csv, the Japanese-named file, sjis_log.txt and app.log differ");
        System.out.println("  TC003_one_sided       only_in_input.txt left only, only_in_output.txt right only");
        System.out.println("  TC004_no_log_compare  the LOG pair is recorded as an error");
        System.out.println("  nested/TC005_nested   found despite nesting; data.csv differs");
    }

    /** Test case where both pairs match exactly. */
    private static void makeIdentical(Path tc) throws IOException {
        for (String dir : new String[]{"INPUT", "OUTPUT"}) {
            write(tc.resolve(dir).resolve("data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
            write(tc.resolve(dir).resolve("sub/note.txt"), "\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u5185\u306e\u30d5\u30a1\u30a4\u30eb\u3067\u3059\u3002" + NL, StandardCharsets.UTF_8);
        }
        for (String dir : new String[]{"LOG", "LOG_COMPARE"}) {
            write(tc.resolve(dir).resolve("app.log"), log("\u51e6\u7406\u3092\u7d42\u4e86\u3057\u307e\u3057\u305f"), StandardCharsets.UTF_8);
        }
    }

    /** Test case with content differences, mixing encodings and file name shapes. */
    private static void makeDifferent(Path tc) throws IOException {
        write(tc.resolve("INPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/data.csv"), csv("2,BETA"), StandardCharsets.UTF_8);

        write(tc.resolve("INPUT/\u65e5\u672c\u8a9e\u30d5\u30a1\u30a4\u30eb\u540d.txt"),
                "1\u884c\u76ee\uff1a\u5909\u66f4\u306a\u3057" + NL + "2\u884c\u76ee\uff1a\u5909\u66f4\u524d" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/\u65e5\u672c\u8a9e\u30d5\u30a1\u30a4\u30eb\u540d.txt"),
                "1\u884c\u76ee\uff1a\u5909\u66f4\u306a\u3057" + NL + "2\u884c\u76ee\uff1a\u5909\u66f4\u5f8c" + NL, StandardCharsets.UTF_8);

        // Shift_JIS file: checks that WinMerge detects the encoding rather than us.
        write(tc.resolve("INPUT/sjis_log.txt"), log("\u30b7\u30d5\u30c8JIS\u306e\u30ed\u30b0\u3067\u3059"), SJIS);
        write(tc.resolve("OUTPUT/sjis_log.txt"), log("\u30b7\u30d5\u30c8JIS\u306e\u30ed\u30b0\u3067\u3059\uff08\u5909\u66f4\u5f8c\uff09"), SJIS);

        write(tc.resolve("INPUT/sub/note.txt"), "\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u5185\u306e\u30d5\u30a1\u30a4\u30eb\u3067\u3059\u3002" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/sub/note.txt"), "\u30b5\u30d6\u30d5\u30a9\u30eb\u30c0\u5185\u306e\u30d5\u30a1\u30a4\u30eb\u3067\u3059\uff08\u5909\u66f4\u5f8c\uff09\u3002" + NL, StandardCharsets.UTF_8);

        write(tc.resolve("LOG/app.log"), log("\u51e6\u7406\u3092\u7d42\u4e86\u3057\u307e\u3057\u305f"), StandardCharsets.UTF_8);
        write(tc.resolve("LOG_COMPARE/app.log"), log("\u51e6\u7406\u3092\u7d42\u4e86\u3057\u307e\u3057\u305f\uff08\u60f3\u5b9a\u5024\uff09"), StandardCharsets.UTF_8);
    }

    /** Test case with one-sided files and an empty LOG_COMPARE directory. */
    private static void makeOneSided(Path tc) throws IOException {
        write(tc.resolve("INPUT/common.txt"), "\u4e21\u65b9\u306b\u3042\u308b\u30d5\u30a1\u30a4\u30eb" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/common.txt"), "\u4e21\u65b9\u306b\u3042\u308b\u30d5\u30a1\u30a4\u30eb" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("INPUT/only_in_input.txt"), "INPUT \u306b\u3057\u304b\u306a\u3044\u30d5\u30a1\u30a4\u30eb" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/only_in_output.txt"), "OUTPUT \u306b\u3057\u304b\u306a\u3044\u30d5\u30a1\u30a4\u30eb" + NL, StandardCharsets.UTF_8);

        write(tc.resolve("LOG/app.log"), log("\u51e6\u7406\u3092\u7d42\u4e86\u3057\u307e\u3057\u305f"), StandardCharsets.UTF_8);
        Files.createDirectories(tc.resolve("LOG_COMPARE")); // intentionally left empty
    }

    /** Test case missing the LOG_COMPARE directory, which must surface as an error. */
    private static void makeMissingPairDir(Path tc) throws IOException {
        write(tc.resolve("INPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("LOG/app.log"), log("\u51e6\u7406\u3092\u7d42\u4e86\u3057\u307e\u3057\u305f"), StandardCharsets.UTF_8);
        // LOG_COMPARE is deliberately not created.
    }

    private static String csv(String lastRow) {
        return "id,name" + NL + "1,alpha" + NL + lastRow + NL;
    }

    private static String log(String lastLine) {
        return "2026-08-31 10:00:00 INFO  \u51e6\u7406\u3092\u958b\u59cb\u3057\u307e\u3057\u305f" + NL
                + "2026-08-31 10:00:01 DEBUG \u30ec\u30b3\u30fc\u30c9\u4ef6\u6570: 2" + NL
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
