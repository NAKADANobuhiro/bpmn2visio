# bpmn2visio

BPMN 2.0 ファイル（`.bpmn`）を Microsoft Visio の `.vsdx` に自動変換するツールです。
Visio COM オートメーション（pywin32）とレーン数別テンプレート方式を採用し、
正規の BPMN ステンシルマスター（タスク・ゲートウェイ・イベント等）と GlueTo コネクタを使用します。

> **注意：** スイムレーンを動的に生成することはできません。生成後に移動・分離・サイズ変更は Visio の画面上で実施してください。

**English:** [README.md](README.md)

## 特徴

- 正規の Visio BPMN マスター図形を使用（タスク・開始/終了イベント・ゲートウェイ等）
- GlueTo による動的コネクタ（図形移動後も接続維持）
- Visio GUARD() 問題を回避するテンプレート方式（A3 横固定）
- 2〜5 レーンに対応（`templates/lane2.vstx` ～ `lane5.vstx`）
- BPMN レーン座標に基づく図形の自動配置

## 動作環境

- Windows 10/11
- Microsoft Visio Plan2 2019 以降（デスクトップ版）
- Python 3.8 以降
- pywin32：`pip install pywin32`

## ファイル構成

```
bpmn2visio/
├── bpmn2visio.py          # メイン変換スクリプト（BPMN パーサー含む）
├── sample.bpmn             # サンプル BPMN（3 レーン・受注処理プロセス）
├── templates/              # レーン数別 Visio テンプレート（手動作成が必要）
│   ├── lane2.vstx
│   ├── lane3.vstx
│   ├── lane4.vstx
│   └── lane5.vstx
├── bpmn2visio.bat         # エクスプローラー用ドラッグ＆ドロップランチャー
├── README.md
└── README.ja.md
```

> **注意：** `templates/` フォルダ内の `.vstx` ファイルはバイナリのため Git 管理外です。
> 作成手順は「テンプレートの準備」セクションを参照してください。

## クイックスタート

### 1. pywin32 のインストール

```
pip install pywin32
```

### 2. テンプレートの準備

Visio テンプレート（`.vstx`）はレーン数ごとに Visio UI で手動作成します。
詳細は「[テンプレートの準備](#テンプレートの準備)」セクションを参照してください。

### 3. 変換の実行

```
python bpmn2visio.py sample.bpmn
```

出力：`sample_com.vsdx`（入力ファイルと同じフォルダ）

### 4. エクスプローラーからドロップして変換

Windows エクスプローラーで `.bpmn` ファイルを **`bpmn2visio.bat`** にドラッグ＆ドロップします。
変換結果の `.vsdx` は `.bpmn` ファイルと同じフォルダに生成されます。
変換後もコンソールウィンドウが開いたままになるため、ログを確認できます。

### 5. 複数ファイルの一括変換

```
for %f in (*.bpmn) do python bpmn2visio.py %f
```

## 対応 BPMN 要素

| BPMN 要素 | Visio マスター |
|-----------|--------------|
| startEvent | 開始イベント |
| endEvent | 終了イベント |
| userTask / task | タスク |
| exclusiveGateway | ゲートウェイ（XOR） |
| parallelGateway | ゲートウェイ（AND） |
| inclusiveGateway | ゲートウェイ（OR） |
| subProcess | 展開されたサブプロセス |
| intermediateCatchEvent | 中間イベント |
| sequenceFlow | シーケンス フロー |

## テンプレートの準備

テンプレートは Visio の正規 BPMN ファイルを `.vstx` 形式で保存したものです。使用するレーン数ごとに 1 回だけ作成します。

1. Visio を起動 → **ファイル > 新規 > BPMN 図**
2. ページサイズを A3 横に変更：**デザイン > ページ設定 > 用紙サイズ: A3 / 向き: 横**
3. 左パネルの BPMN ステンシルから **「プール/レーン」** マスターをキャンバスにドロップ
4. Swimlane List を右クリック → **「レーンを追加」** を必要な数だけ繰り返す
5. プール全体をページいっぱいに広げる（GUARD() のため COM では変更不可）
6. **ファイル > 名前を付けて保存 > Visio テンプレート (*.vstx)** で保存 → `templates/lane3.vstx`
7. 必要なレーン数ぶん繰り返す（`lane2.vstx`, `lane4.vstx`, `lane5.vstx` …）

対応するテンプレートがない場合、スクリプトは最も近いレーン数のテンプレートで自動的に代替します。

## トラブルシューティング

**「テンプレートディレクトリが見つかりません」エラー**
→ `bpmn2visio.py` と同じフォルダに `templates/` を作成し、必要な `.vstx` ファイルを追加してください。

**「BPMN ステンシルが見つかりません」エラー**
→ スクリプト実行前に Visio で BPMN テンプレートを開いてください（ファイル > 新規 > BPMN 図）。スクリプトは起動中の Visio インスタンスに接続してステンシルを検索します。

**「マスターが見つかりません」警告**
→ Visio の言語バージョンによってマスター名が異なります。利用可能なマスター名を確認するには Visio で以下のマクロを実行してください（Alt+F11 → モジュール挿入 → 実行）：
```vba
Sub ListMasters()
    Dim msg As String, oD As Visio.Document, oM As Visio.Master
    For Each oD In Application.Documents
        If oD.Type = visTypeStencil Then
            For Each oM In oD.Masters : msg = msg & oM.Name & Chr(13) : Next
        End If
    Next
    MsgBox msg
End Sub
```
確認したマスター名を `bpmn2visio.py` の `MASTER_CANDIDATES` に追加してください。

**「BPMNレーン数 ≠ テンプレートレーン数」警告**
→ BPMN ファイルのレーン数に対応するテンプレート（`templates/lane{N}.vstx`）を作成してください。

## 仕組み

bpmn2visio は、COM から BPMN Pool/Lane 図形のサイズを変更できない Visio の GUARD() 制約を回避するため、正しいプール/レーン構造をあらかじめ持つ `.vstx` テンプレートを開く方式を採用しています。スクリプトはテンプレートからレーンの座標を読み取り、BPMN 要素の位置をレーン相対スケーリングで Visio 座標に変換し、`GlueTo` を使って動的コネクタで図形を接続します。

## 関連プロジェクト

- [markdown2bpmn](https://github.com/NAKADANobuhiro/markdown2bpmn) : Markdown 形式で BPMN を記述し、`.bpmn` ファイルを生成するツール

## ライセンス

MIT
