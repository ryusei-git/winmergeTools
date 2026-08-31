# winmergeTools

WinMerge を使って、テストケースごとのフォルダ/ファイル比較レポート（HTML）を自動生成し、
Excel ブックにまとめる単一ファイルの Java ツールです。

- 実装は `WinMergeReportTool.java` の 1 ファイルのみ（外部ライブラリ不要 / JDK 11 以降）

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
└─ TC001/
    ├─ INPUT_vs_OUTPUT/
    │   ├─ _folder_compare.html                 ← フォルダ比較結果
    │   └─ files/<相対パス>.html                ← ファイル比較結果
    └─ LOG_vs_LOG_COMPARE/
        ├─ _folder_compare.html
        └─ files/<相対パス>.html
```

## 実行方法

JDK 11 以降であれば、コンパイル不要で直接実行できます。

```bat
cd D:\tests
java WinMergeReportTool.java
```

コンパイルして使う場合:

```bat
javac -encoding UTF-8 WinMergeReportTool.java
java WinMergeReportTool
```

### オプション

| オプション | 説明 |
| --- | --- |
| `--root <dir>` | 比較対象のルート（既定: カレントディレクトリ） |
| `--report <dir>` | レポート出力先（既定: `<root>/report`） |
| `--excel <file>` | Excel 出力先（既定: `<report>/comparison_report_<日時>.xlsx`） |
| `--winmerge <exe>` | `WinMergeU.exe` のパス |
| `--pair <L>:<R>` | 比較するサブディレクトリの組（複数指定可） |
| `--excel-mode <mode>` | `auto` / `com` / `xlsx` / `none`（既定: `auto`） |
| `--timeout <sec>` | WinMerge 1 回あたりのタイムアウト秒（既定: 180） |
| `--max-rows <n>` | xlsx へ展開する 1 レポートあたりの最大行数（既定: 500） |
| `--clean` | 実行前に `report` フォルダを削除する |
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

## Excel 出力の 2 モード

| モード | 内容 | 前提 |
| --- | --- | --- |
| `com` | VBScript 経由で Excel を操作し、HTML レポートを**そのままシートへ貼り付け**（書式・色を保持） | Windows + Excel + `cscript` |
| `xlsx` | 外部ライブラリなしで `.xlsx` を直接生成。HTML の表を解析してセルに展開し、元 HTML へのハイパーリンクを付与 | なし |

既定の `auto` は Windows なら `com` を試し、失敗した場合に `xlsx` へフォールバックします。

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

## 補足

- 「選択されたファイルは同一です」のダイアログが出る環境では、一度手動で
  「次回から表示しない」にチェックを入れてください（`/noninteractive` でも抑止されない場合があります）。
- レポート形式は `/cfg ReportFiles/ReportType=2`（Simple HTML）を指定しています。
  変更したい場合は `--winmerge-arg` で追加設定を渡せます。
