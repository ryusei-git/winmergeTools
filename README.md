# winmergeTools

WinMerge でテストケースごとのフォルダ／ファイル比較を自動実行し、HTML レポートと Excel ブックにまとめるツールです。

**実装は `WinMergeReportTool.java` の 1 ファイルだけ。** 外部ライブラリは不要で、このファイルを配布先にコピーすれば動きます。

---

## 1. すぐ試す

```bat
cd D:\tests
javac -encoding UTF-8 -d %TEMP%\wmrt C:\tools\winmergeTools\WinMergeReportTool.java
java -cp %TEMP%\wmrt WinMergeReportTool
```

`D:\tests` 配下のテストケースを比較し、`D:\tests\report` に結果を出力します。

JDK 18 以降なら、コンパイルせず 1 コマンドでも実行できます。

```bat
cd D:\tests
java C:\tools\winmergeTools\WinMergeReportTool.java
```

> ソースは全て ASCII なので、JDK 17 以前でもこの 1 コマンド実行は文字化けしません。
> ただし単一ファイル実行自体が JDK 11 以降の機能です。

### 必要なもの

| | 用途 |
| --- | --- |
| **JDK 11 以降** | 必須。`java` だけの JRE ではなく `javac` を含むもの |
| **WinMerge 2.16 以降** | 必須。2.14 以前は `--no-exitcode` が必要（→ 6 章） |
| **Edge か Chrome** | `--excel-mode image` のときだけ。Windows 標準の Edge で可 |
| **Excel** | `--excel-mode com` のときだけ |

---

## 2. 入力と出力

### 入力（比較対象）

テストケースのディレクトリに `INPUT` / `OUTPUT` / `LOG` / `LOG_COMPARE` を置きます。

```
D:\tests\
├─ TC001\
│   ├─ INPUT\        ┐ 比較される
│   ├─ OUTPUT\       ┘
│   ├─ LOG\          ┐ 比較される
│   └─ LOG_COMPARE\  ┘
└─ suite\
    └─ TC002\        ← 入れ子でも検出される
        ├─ INPUT\ ...
```

### 出力

```
D:\tests\report\
├─ index.html                          ← 全比較結果の一覧。まずここを開く
├─ comparison_report_<日時>.xlsx       ← Excel ブック
└─ TC001\
    ├─ INPUT_vs_OUTPUT\
    │   ├─ _folder_compare.html        ← フォルダ比較の結果
    │   └─ files\<相対パス>.html       ← ファイルごとの比較結果
    └─ LOG_vs_LOG_COMPARE\
        └─ （同じ構成）
```

Excel ブックは `サマリ`（Summary）シート＋テストケース名のシートで構成されます。

---

## 3. 処理の流れ

1. `--root`（既定はカレントディレクトリ）配下のディレクトリ階層を `report` に複製する
2. `INPUT` / `OUTPUT` / `LOG` / `LOG_COMPARE` を持つディレクトリをテストケースとして検出する
3. ペアごとに **フォルダ比較**（`/r` で再帰）を 1 回実行する
4. 両側にある同名ファイルごとに **ファイル比較** を実行する
   （片側にしか無いファイルは WinMerge を呼ばず「left only / right only」として記録）
5. `index.html` と Excel ブックを生成する

判定は WinMerge の終了コード（`0`=一致 / `1`=差分 / `2`以上=エラー）で行います。

---

## 4. Excel 出力の 3 モード

`--excel-mode` で選びます。既定は `auto`（Windows なら `com` を試し、失敗したら `xlsx`）。

| モード | 見た目 | 文字の検索 | 必要なもの |
| --- | --- | --- | --- |
| `image` | ◎ 完全に再現 | ✕ 画像なので不可 | Edge か Chrome |
| `com` | ○ 表として貼り付け | ◎ | **Excel** |
| `xlsx` | △ 書式なし | ◎ | なし |

```bat
rem 見た目重視。Excel が無くても動く
java -cp %TEMP%\wmrt WinMergeReportTool --excel-mode image

rem Excel で開いて中身を検索・コピーしたい
java -cp %TEMP%\wmrt WinMergeReportTool --excel-mode com

rem 何も追加要件なしで確実に出したい
java -cp %TEMP%\wmrt WinMergeReportTool --excel-mode xlsx
```

- `image`: HTML を Edge のヘッドレスモードで PNG 化して埋め込みます。PNG は `report\_images\` にも残ります。高さはレポートの行数から見積もるため、極端に長い差分は下端が切れることがあります（そのときはシート内のリンクから元 HTML へ）。
- `com`: 生成した VBScript を `cscript` で実行して Excel を操作します。失敗した場合は `report\_build_excel.vbs` と `report\_build_excel.log` に原因が残ります。
- `xlsx`: HTML の表を解析してセルに展開し、元 HTML へのハイパーリンクを張ります。

---

## 5. コマンドラインオプション

| オプション | 既定値 | 説明 |
| --- | --- | --- |
| `--root <dir>` | カレントディレクトリ | 比較対象のルート |
| `--report <dir>` | `<root>\report` | レポート出力先 |
| `--excel <file>` | `<report>\comparison_report_<日時>.xlsx` | Excel の出力先 |
| `--winmerge <exe>` | 自動探索 | `WinMergeU.exe` のパス |
| `--pair <L>:<R>` | `INPUT:OUTPUT` と `LOG:LOG_COMPARE` | 比較する組。複数指定可 |
| `--excel-mode <mode>` | `auto` | `auto` / `com` / `image` / `xlsx` / `none` |
| `--browser <exe>` | 自動探索 | PNG 化に使う Edge/Chrome |
| `--image-width <px>` | `1600` | PNG の横幅 |
| `--image-max-height <px>` | `16000` | PNG の高さ上限 |
| `--timeout <sec>` | `180` | WinMerge 1 回あたりのタイムアウト |
| `--max-rows <n>` | `500` | xlsx へ展開する 1 レポートの最大行数 |
| `--max-path <n>` | `240` | レポートパスの上限文字数 |
| `--clean` | — | 実行前に `report` を削除する |
| `--no-exitcode` | — | 一致判定を Java 側のバイト比較で行う |
| `--winmerge-arg <a>` | — | WinMerge へ渡す追加引数。複数指定可 |
| `--help` | — | ヘルプ表示 |

### よく使う組み合わせ

```bat
rem WinMerge の場所を明示する
java -cp %TEMP%\wmrt WinMergeReportTool --winmerge "C:\Program Files\WinMerge\WinMergeU.exe"

rem 比較対象と出力先を別々に指定する
java -cp %TEMP%\wmrt WinMergeReportTool --root D:\tests --report E:\out\report

rem 比較するフォルダの組を変える
java -cp %TEMP%\wmrt WinMergeReportTool --pair EXPECTED:ACTUAL --pair LOG:LOG_EXPECTED

rem 毎回まっさらな report を作り直す
java -cp %TEMP%\wmrt WinMergeReportTool --clean

rem WinMerge に追加設定を渡す（行末の空白を無視する例）
java -cp %TEMP%\wmrt WinMergeReportTool --winmerge-arg /cfg --winmerge-arg Settings/IgnoreSpace=1
```

`WinMergeU.exe` は `--winmerge` → 環境変数 `WINMERGE_PATH` → `C:\Program Files\WinMerge\` などの順に探します。

### 終了コード

| コード | 意味 |
| --- | --- |
| `0` | 正常終了（差分の有無は問わない） |
| `1` | 比較エラーあり / WinMerge 未検出 / 実行時エラー |
| `2` | コマンドライン引数の誤り |

---

## 6. つまずきやすいところ

### WinMerge が 2.14 以前だと全件「一致」になる

判定に使う `/enableexitcode` が **WinMerge 2.16 以降**の機能です。これが無いと WinMerge は常に 0 を返すため、差分があっても「一致」と記録されます。

古い WinMerge を使う場合は `--no-exitcode` を付けてください。レポート生成は WinMerge が行い、一致判定だけを Java 側のバイト比較で代替します（WinMerge の比較フィルタや改行コード無視の設定は反映されません）。

判定が正しいかは `tools\setup_test_env.bat` で確認できます（→ 7 章）。

### パスが長いとレポートが作れない

`report\<TC>\<比較名>\files\<相対パス>.html` は元のパスより必ず長くなり、Windows の 260 文字制限に当たります。`--max-path`（既定 240）を超える場合は自動的に短縮名へ切り替えますが、`--report` で浅い場所を指定するのが確実です。

### WinMerge を起動したまま実行しない

既存インスタンスに処理が渡り、レポートが作られないことがあります。

### 「選択されたファイルは同一です」のダイアログ

出る環境では、一度手動で「次回から表示しない」にチェックを入れてください。止まった場合も `--timeout`（既定 180 秒）で強制終了して次へ進みます。

### Excel の `com` モード実行中は Excel を触らない

裏で Excel が自動操作されます。他のブックは閉じておいてください。

---

## 7. 動作確認用のツール（任意）

配布先の環境が正しいかを確認するスクリプトを `tools\` に用意しています。**本番運用には不要です。**

```bat
cd /d D:\tests
C:\tools\winmergeTools\tools\setup_test_env.bat
```

1. JDK を探してコンパイル
2. サンプルデータを `D:\tests\sample` に生成
3. `WinMergeU.exe` の場所を特定
4. **WinMerge の疎通確認** — 差分あり／一致のフォルダ比較を実際に実行

最後の判定行が要点です。

| 出力 | 意味 |
| --- | --- |
| `OK : /enableexitcode works.` | そのまま使える |
| `NG : the folders differ but WinMerge returned 0.` | `--no-exitcode` を付ける |

続けて本番と同じ実行をします。

```bat
cd /d D:\tests\sample
C:\tools\winmergeTools\run.bat
```

集計が次のようになれば正常です。

```
  Test cases : 5
  Identical  : 11
  Different  : 11
  One-sided  : 3
  Errors     : 1
```

生成されるサンプルの内訳:

| テストケース | 期待する結果 |
| --- | --- |
| `TC001_identical` | すべて「identical」 |
| `TC002_different` | すべて「different」（UTF-8 / Shift_JIS / 日本語ファイル名 / サブフォルダを含む） |
| `TC003_one_sided` | 「left only」「right only」 |
| `TC004_no_log_compare` | LOG ペアが「error」 |
| `nested/TC005_nested` | 入れ子でも検出され `data.csv` が「different」 |

### JDK を入れたくない場合

zip 版の JDK を展開するだけで動きます（インストーラ・レジストリ・システム PATH を触りません）。

1. [Adoptium](https://adoptium.net/temurin/releases/) で **Windows / x64 / JDK / .zip** を取得
2. `C:\tools` に展開（`C:\tools\jdk-21.0.x+y` のようなフォルダ名のままでよい）
3. `run.bat` と `setup_test_env.bat` が自動的に見つけます

別の場所に置いた場合は `set "JAVA_HOME=E:\portable\jdk-21"` を先に実行してください。

---

## 8. 保守について

このソースは **Microsoft Copilot による保守**を前提にしています。

- **ソースは全て ASCII。** UTF-8 の日本語をソースに書き戻さないでください。JDK 17 以前は既定のソース文字コードが MS932 になり、`javac -encoding UTF-8` を付けない限り壊れます。コメント・メッセージともに英語で書いてください
- **外部ライブラリは追加しない。** `.xlsx` は `java.util.zip` で手組みしているため、OOXML のパーツ構成・リレーション ID・コンテンツタイプの整合性を壊さないこと
- **Java 11 が言語レベルの下限。** `javac --release 11` が通ることを確認してください
- **ファイル入出力では文字コードを必ず明示する。** プラットフォーム既定に依存しないこと
- 各メソッドのコメントには「なぜそうしているか」を書いてあります。特に `/enableexitcode`、VBScript の `Err` の扱い、OOXML のリレーション採番は、理由を読まずに変更すると壊れます

| ファイル | 役割 |
| --- | --- |
| `WinMergeReportTool.java` | 本体。これ 1 つで動く |
| `run.bat` | Windows 用の起動バッチ（任意） |
| `tools\MakeSampleData.java` | 動作確認用サンプルの生成（任意） |
| `tools\setup_test_env.bat` | 環境確認と WinMerge の疎通確認（任意） |
| `tools\find_jdk.bat` | JDK の探索（上記バッチから呼ばれる） |

`.bat` は **ASCII のみ・CRLF 改行**にしてください。cmd.exe はバッチをバイト位置で読むため、`chcp` とマルチバイト文字が混在すると行の切れ目がずれて壊れます（`.gitattributes` で CRLF を固定しています）。
