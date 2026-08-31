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
 * WinMergeReportTool の動作確認用サンプルデータを生成する。
 *
 * <p>期待する挙動を一通り踏むように、次のテストケースを作る。
 * <ul>
 *   <li>TC001_一致          : INPUT/OUTPUT も LOG/LOG_COMPARE も完全一致</li>
 *   <li>TC002_差分あり      : 内容差分あり（UTF-8 / Shift_JIS / 日本語ファイル名を含む）</li>
 *   <li>TC003_片側のみ      : 片方にしか無いファイル、空の LOG_COMPARE</li>
 *   <li>TC004_LOG_COMPAREなし : LOG_COMPARE 自体が無く、エラーとして記録される</li>
 *   <li>nested/TC005_入れ子 : 入れ子のテストケースが検出されるか</li>
 * </ul>
 *
 * <p>実行例:
 * <pre>
 *   javac -encoding UTF-8 -d %TEMP%\wmrt tools\MakeSampleData.java
 *   java -cp %TEMP%\wmrt MakeSampleData --dest D:\tests\sample --clean
 * </pre>
 */
public final class MakeSampleData {

    /** テキストファイルは Windows のテストデータらしく CRLF で書き出す。 */
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
        // 入れ子のケースだけ 1 箇所差分を入れて、検出されたことが分かるようにする
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

    /** INPUT/OUTPUT・LOG/LOG_COMPARE がすべて一致するテストケース。 */
    private static void makeIdentical(Path tc) throws IOException {
        for (String dir : new String[]{"INPUT", "OUTPUT"}) {
            write(tc.resolve(dir).resolve("data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
            write(tc.resolve(dir).resolve("sub/note.txt"), "サブフォルダ内のファイルです。" + NL, StandardCharsets.UTF_8);
        }
        for (String dir : new String[]{"LOG", "LOG_COMPARE"}) {
            write(tc.resolve(dir).resolve("app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        }
    }

    /** 内容に差分があるテストケース。文字コードとファイル名のパターンも混ぜる。 */
    private static void makeDifferent(Path tc) throws IOException {
        write(tc.resolve("INPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/data.csv"), csv("2,BETA"), StandardCharsets.UTF_8);

        write(tc.resolve("INPUT/日本語ファイル名.txt"),
                "1行目：変更なし" + NL + "2行目：変更前" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/日本語ファイル名.txt"),
                "1行目：変更なし" + NL + "2行目：変更後" + NL, StandardCharsets.UTF_8);

        // Shift_JIS のファイル。WinMerge の文字コード自動判別の確認用
        write(tc.resolve("INPUT/sjis_log.txt"), log("シフトJISのログです"), SJIS);
        write(tc.resolve("OUTPUT/sjis_log.txt"), log("シフトJISのログです（変更後）"), SJIS);

        write(tc.resolve("INPUT/sub/note.txt"), "サブフォルダ内のファイルです。" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/sub/note.txt"), "サブフォルダ内のファイルです（変更後）。" + NL, StandardCharsets.UTF_8);

        write(tc.resolve("LOG/app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        write(tc.resolve("LOG_COMPARE/app.log"), log("処理を終了しました（想定値）"), StandardCharsets.UTF_8);
    }

    /** 片方にしか存在しないファイルと、空の LOG_COMPARE を持つテストケース。 */
    private static void makeOneSided(Path tc) throws IOException {
        write(tc.resolve("INPUT/common.txt"), "両方にあるファイル" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/common.txt"), "両方にあるファイル" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("INPUT/only_in_input.txt"), "INPUT にしかないファイル" + NL, StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/only_in_output.txt"), "OUTPUT にしかないファイル" + NL, StandardCharsets.UTF_8);

        write(tc.resolve("LOG/app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        Files.createDirectories(tc.resolve("LOG_COMPARE")); // 空のまま
    }

    /** LOG_COMPARE ディレクトリ自体が無く、エラーとして記録されるテストケース。 */
    private static void makeMissingPairDir(Path tc) throws IOException {
        write(tc.resolve("INPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("OUTPUT/data.csv"), csv("2,beta"), StandardCharsets.UTF_8);
        write(tc.resolve("LOG/app.log"), log("処理を終了しました"), StandardCharsets.UTF_8);
        // LOG_COMPARE は作らない
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
