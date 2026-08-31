# winmergeTools

WinMerge を使って、テストケースごとのフォルダ/ファイル比較レポート（HTML）を自動生成し、
Excel ブックにまとめる単一ファイルの Java ツールです。

- 実装は `WinMergeReportTool.java` の 1 ファイルのみ（外部ライブラリ不要 / JDK 11 以降）
- `run.bat` … Windows 用の起動バッチ（任意）
- `tools/` … 動作確認用のサンプルデータ生成と疎通確認スクリプト（任意）

## 想定するディレクトリ構成

```
<カレントディレクトリ>
├─ TC001/                  ← テストケースディレクトリ
│   ├─ INPUT/
│   ├─ OUTPUT/
│   ├─ LOG/
│   └─ LOG_COMPARE/
├─ サブフォルダ/
│   └─ TC002/              ← 入れ子になっていても検出する
│       ├─ INPUT/ ...
└─ report/                 ← 出力先（自動生成）
```

## 処理内容

1. カレントディレクトリ配下のディレクトリ階層を `report` フォルダの中に複製する
2. `INPUT` / `OUTPUT` / `LOG` / `LOG_COMPARE` を持つディレクトリをテストケースとして検出する
3. `INPUT` ⇔ `OUTPUT`、`LOG` ⇔ `LOG_COMPARE` を **フォルダ比較**（`/r` で再帰）し HTML レポートを出力
4. 両側に存在する同名ファイル同士を **ファイル比較** し HTML レポートを出力
   （片側にしか無いファイルは「左のみ / 右のみ」として記録）
5. 比較結果を `report` 配下の対応するディレクトリへ配置
6. テストケース名をシート名にした Excel ブックを作成し、HTML レポートを各シートへ貼り付け
7. 全結果の一覧 `report/index.html` を生成

### 出力レイアウト

```
report/
├─ index.html                                   ← 一覧
├─ comparison_report_<日時>.xlsx                ← Excel
├─ _build_excel.vbs                             ← COM 貼り付け用スクリプト（com モード時）
├─ _build_excel.log                             ← COM 実行時の診断ログ（com モード時）
├─ _images/                                     ← PNG 化したレポート（image モード時）
└─ TC001/
    ├─ INPUT_vs_OUTPUT/
    │   ├─ _folder_compare.html                 ← フォルダ比較結果
    │   └─ files/<相対パス>.html                ← ファイル比較結果
    └─ LOG_vs_LOG_COMPARE/
        ├─ _folder_compare.html
        └─ files/<相対パス>.html
```

## 実行方法（Windows）

同梱の `run.bat` を使うのが確実です（`javac -encoding UTF-8` でコンパイルします）。
カレントディレクトリは呼び出し元のままなので、比較したい場所へ `cd` してから実行してください。

```bat
cd D:\tests
C:\tools\winmergeTools\run.bat
```

`java` コマンドだけで実行する場合:

```bat
cd D:\tests
javac -encoding UTF-8 -d %TEMP%\wmrt C:\tools\winmergeTools\WinMergeReportTool.java
java -Dfile.encoding=UTF-8 -cp %TEMP%\wmrt WinMergeReportTool
```

> **注意**: `java WinMergeReportTool.java`（単一ファイル実行）は **JDK 18 以降** でのみ安全です。
> JDK 17 以前は日本語 Windows のソース既定文字コードが MS932 になり、UTF-8 のソースに含まれる
> 日本語リテラルが壊れます。その場合は上記のように `javac -encoding UTF-8` でコンパイルしてください。

### オプション

| オプション | 説明 |
| --- | --- |
| `--root <dir>` | 比較対象のルート（既定: カレントディレクトリ） |
| `--report <dir>` | レポート出力先（既定: `<root>/report`） |
| `--excel <file>` | Excel 出力先（既定: `<report>/comparison_report_<日時>.xlsx`） |
| `--winmerge <exe>` | `WinMergeU.exe` のパス |
| `--pair <L>:<R>` | 比較するサブディレクトリの組（複数指定可） |
| `--excel-mode <mode>` | `auto` / `com` / `image` / `xlsx` / `none`（既定: `auto`） |
| `--browser <exe>` | PNG 化に使う Edge/Chrome のパス（既定: 自動探索） |
| `--image-width <px>` | PNG の横幅（既定: 1600） |
| `--image-max-height <px>` | PNG の高さ上限（既定: 16000） |
| `--timeout <sec>` | WinMerge 1 回あたりのタイムアウト秒（既定: 180） |
| `--max-rows <n>` | xlsx へ展開する 1 レポートあたりの最大行数（既定: 500） |
| `--clean` | 実行前に `report` フォルダを削除する |
| `--no-exitcode` | `/enableexitcode` を使わず、一致判定を Java 側の内容比較で行う |
| `--max-path <n>` | レポートパスの上限文字数（既定: 240。超える場合は短縮名にする） |
| `--winmerge-arg <a>` | WinMerge へ渡す追加引数（複数指定可） |
| `--help` | ヘルプ表示 |

```bat
java WinMergeReportTool.java ^
     --winmerge "C:\Program Files\WinMerge\WinMergeU.exe" ^
     --pair INPUT:OUTPUT --pair LOG:LOG_COMPARE ^
     --excel D:\tests\report\result.xlsx
```

`WinMergeU.exe` は `--winmerge` → 環境変数 `WINMERGE_PATH` → システムプロパティ `winmerge.path`
→ 既定のインストール先（`C:\Program Files\WinMerge\` など）の順に探索します。

## JDK を入れたくない場合（ポータブル JDK）

インストーラを使わず、zip を展開するだけで動かせます。レジストリ・システム環境変数・
システムの PATH は一切変更されず、撤去はフォルダを削除するだけです。

1. [Adoptium](https://adoptium.net/temurin/releases/) で
   **OS=Windows / Architecture=x64 / Package Type=JDK / 拡張子 .zip** を選んでダウンロード
   （`.msi` はインストーラなので選ばない）
2. `C:\tools\jdk` に展開する（`C:\tools\jdk\bin\javac.exe` になるように）
3. あとは通常どおり `run.bat` / `tools\setup_test_env.bat` を実行するだけ

`tools/find_jdk.bat` が次の順に JDK を探し、見つけた場所をそのバッチの中でだけ使います。

1. PATH 上の `javac`
2. `%JAVA_HOME%\bin\javac.exe`
3. リポジトリ直下の `jdk\` フォルダ（`C:\tools\winmergeTools\jdk\`）
4. `C:\tools\jdk` / `C:\jdk` / `D:\tools\jdk`、および `C:\tools\jdk*`（展開直後の
   `jdk-21.0.5+11` のようなフォルダ名のままでも拾う）

別の場所に展開した場合は、実行するコマンドプロンプトで指定してください。

```bat
set "JAVA_HOME=E:\portable\jdk-21"
C:\tools\winmergeTools\tools\setup_test_env.bat
```

## テスト環境の用意（Windows）

WinMerge をインストールした直後の動作確認用に、サンプルデータの生成と疎通確認を行う
スクリプトを `tools/` に用意しています。

```bat
cd /d D:\tests
C:\tools\winmergeTools\tools\setup_test_env.bat
```

このバッチは次の 4 つを順に実行します。

1. JDK の確認と、`WinMergeReportTool.java` / `MakeSampleData.java` のコンパイル
2. サンプルデータの生成（`D:\tests\sample`）
3. `WinMergeU.exe` の場所の特定（`WINMERGE_PATH` → 既定のインストール先）
4. **WinMerge の疎通確認** — 差分あり／一致のフォルダ比較を実際に実行し、
   終了コードと HTML レポートの生成を確認する

疎通確認の判定はここが要点です。

| 結果 | 意味 | 対応 |
| --- | --- | --- |
| diff=1, same=0 | `/enableexitcode` が有効 | そのまま実行できる |
| diff=0 | 差分があるのに 0 が返った | WinMerge が `/enableexitcode` 非対応。`--no-exitcode` を付けて実行する |
| それ以外 | 想定外 | WinMerge のバージョンと `%TEMP%\wm_smoke` のレポートを確認する |

### 生成されるサンプルデータ

`tools/MakeSampleData.java` が、想定する挙動を一通り踏むテストケースを作ります。

| テストケース | 内容 | 期待する結果 |
| --- | --- | --- |
| `TC001_一致` | INPUT/OUTPUT・LOG/LOG_COMPARE が完全一致 | すべて「一致」 |
| `TC002_差分あり` | 内容差分（UTF-8 / Shift_JIS / 日本語ファイル名 / サブフォルダ） | すべて「差分あり」 |
| `TC003_片側のみ` | 片方にしか無いファイル、空の LOG_COMPARE | 「左のみ」「右のみ」 |
| `TC004_LOG_COMPAREなし` | LOG_COMPARE ディレクトリ自体が無い | LOG ペアが「エラー」 |
| `nested/TC005_入れ子` | 入れ子のテストケース | 検出され、`data.csv` が「差分あり」 |

サンプルだけ作り直したい場合:

```bat
java -cp %TEMP%\wmrt MakeSampleData --dest D:\tests\sample --clean
```

生成後、比較を実行します。

```bat
cd /d D:\tests\sample
C:\tools\winmergeTools\run.bat
```

`sample\report\index.html` を開き、上の表の「期待する結果」と一致していれば、
WinMerge との連携は正しく動いています。

## Excel 出力の 3 モード

| モード | 内容 | 前提 |
| --- | --- | --- |
| `com` | VBScript 経由で Excel を操作し、HTML レポートを**そのままシートへ貼り付け**（書式・色を保持） | Windows + Excel + `cscript` |
| `image` | HTML を**ヘッドレスブラウザで PNG 化して画像として貼り付け**。見た目が完全に再現される | Edge か Chrome（Windows 標準の Edge で可） |
| `xlsx` | 外部ライブラリなしで `.xlsx` を直接生成。HTML の表を解析してセルに展開し、元 HTML へのハイパーリンクを付与 | なし |

`image` は Excel が無くても動きます（画像パーツを xlsx へ直接書き込むため）。

```bat
C:\tools\winmergeTools\run.bat --excel-mode image
```

- PNG は `report\_images\<レポートと同じ階層>.png` にも残るので、単体でも参照できます
- 画像の横幅は `--image-width`（既定 1600）、高さ上限は `--image-max-height`（既定 16000）
- 高さはレポートの行数から見積もります。極端に長いレポートでは下端が切れることがあるので、
  その場合は各シートのリンクから元の HTML を参照してください
- 画像なのでセル内の文字は検索・コピーできません。検索性が要るなら `xlsx` モードを使ってください
- ブラウザは Edge → Chrome の順に自動探索します。見つからない場合は `--browser` で指定してください

既定の `auto` は Windows なら `com` を試し、失敗した場合に `xlsx` へフォールバックします。
フォールバックした場合は理由を `[WARN]` として表示し、`report\_build_excel.vbs`（実行した
スクリプト）と `report\_build_excel.log`（各操作の成否）を残すので、そこから原因を追えます。
スクリプトは `cscript //nologo report\_build_excel.vbs` で手動実行して確かめることもできます。

各ブックの構成:

- `サマリ` シート … 全テストケースの比較結果一覧
- テストケース名のシート … そのテストケースの各比較レポート（フォルダ比較 → ファイル比較の順）

シート名は Excel の制約（31 文字・`\ / * ? [ ] :` 不可・重複不可）に合わせて自動整形されます。

## 終了コード

| コード | 意味 |
| --- | --- |
| 0 | 正常終了（差分の有無は問わない） |
| 1 | 比較エラーあり / WinMerge 未検出 / 実行時エラー |
| 2 | コマンドライン引数の誤り |

WinMerge 自体の終了コードは `0=一致` / `1=差分あり` / `2 以上=エラー` として解釈しています。

## Windows で実行するときの注意

### 1. `/enableexitcode`（重要）

一致／差分の判定には WinMerge の終了コードを使うため、`/enableexitcode` を必ず付けています。
**このオプションが無いと WinMerge は常に 0 を返し、全件「一致」と判定されてしまいます。**
WinMerge 2.16 以降で利用できます。それ以前の版を使う場合は `--no-exitcode` を指定してください
（レポート生成は WinMerge、一致判定は Java 側のバイト比較で行います。ただし WinMerge の
比較フィルタや改行コード無視の設定は反映されません）。

### 2. パス長（MAX_PATH 260）

`report/<TC>/<比較名>/files/<相対パス>.html` は元のパスより長くなります。260 文字を超えると
WinMerge も Java もファイルを作れないため、`--max-path`（既定 240）を超える場合は
サブフォルダを潰した平坦な名前 → それでも長ければハッシュ付きの短縮名へ自動的に切り替えます。
出力先自体が深い場合は起動時に警告を出すので、`--report` で浅い場所を指定してください。

### 3. 文字コード

ソースは UTF-8 です。JDK 17 以前で `java WinMergeReportTool.java` を直接実行すると日本語が
壊れるため、`run.bat` か `javac -encoding UTF-8` を使ってください。

一方、**バッチファイル（`.bat`）は ASCII のみ**で書いています。cmd.exe はバッチをバイト位置で
読み進めるため、`chcp 65001` とマルチバイト文字が混在すると行の切れ目がずれ、途中から
壊れたコマンドとして実行されてしまうためです。日本語のメッセージは Java 側が出力します。
`.bat` は CRLF 改行である必要もあるため、`.gitattributes` で `eol=crlf` を強制しています。

実行時に文字コードを指定しないのも意図的です。Java はコンソール自身のコードページで出力するため、
`-Dfile.encoding=UTF-8` を付けない方が日本語が正しく表示されます。ファイル入出力は
コード内で文字コードを明示しているので影響を受けません（WinMerge の標準出力は
`native.encoding`、HTML/xlsx は UTF-8、VBScript は UTF-16LE）。

### 4. その他

- WinMerge を起動したまま実行すると、既存インスタンスへ処理が渡ってレポートが作られない
  ことがあります。実行前に WinMerge は終了しておいてください。
- 「選択されたファイルは同一です」のダイアログが出る環境では、一度手動で
  「次回から表示しない」にチェックを入れてください（`/noninteractive` でも抑止されない場合があります）。
  ダイアログで停止した場合も `--timeout`（既定 180 秒）でプロセスを強制終了し、次の比較へ進みます。
- レポート形式は `/cfg ReportFiles/ReportType=2`（Simple HTML）を指定しています。
  変更したい場合は `--winmerge-arg` で追加設定を渡せます。
- Excel の `com` モードは、実行中に Excel が自動操作されます。他の Excel 作業と同時に実行しないでください。
