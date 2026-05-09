# bpmn2visio

BPMN 2.0 ファイル（`.bpmn`）を Microsoft Visio の `.vsdx` に自動変換するツールです。
Visio COM オートメーション（pywin32）とレーン数別テンプレート方式を採用し、
正規の BPMN ステンシルマスター（タスク・ゲートウェイ・イベント等）と GlueTo コネクタを使用します。

## 特徴

- 正規の Visio BPMN マスター図形を使用（タスク・開始/終了イベント・ゲートウェイ等）
- GlueTo による動的コネクタ（図形移動後も接続維持）
- Visio GUARD() 問題を回避するテンプレート方式（A3 横固定）
- 2〜5 レーンに対応（`templates/lane2.vstx` ～ `lane5.vstx`）
- BPMN レーン座標に基づく図形の自動配置

## 動作環境

- Windows 10/11
- Microsoft Visio 2019 以降（デスクトップ版）
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
├── README.md
├── SPEC.md
├── DESIGN.md
└── Runbook.md
```

> **注意**：`templates/` フォルダ内の `.vstx` ファイルはバイナリのため Git 管理外です。
> 初回セットアップ時は Runbook.md の手順に従って手動作成してください。

## クイックスタート

### 1. pywin32 のインストール

```
pip install pywin32
```

### 2. テンプレートの準備

`templates/` フォルダに `lane3.vstx` 等を作成します（詳細は [Runbook.md](Runbook.md)）。

### 3. 変換の実行

```
python bpmn2visio.py sample.bpmn
```

出力：`sample_com.vsdx`（入力ファイルと同じフォルダ）

### 4. 複数ファイルの一括変換

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

詳細は [SPEC.md](SPEC.md) を参照してください。

## トラブルシューティング

**「テンプレートが見つかりません」エラー**
→ `templates/lane{N}.vstx` が未作成です。Runbook.md を参照して作成してください。

**「BPMN ステンシルが見つかりません」エラー**
→ Visio に BPMN テンプレートを先に開いてください（ファイル > 新規 > BPMN 図）。

**マスターが見つからない警告**
→ 日本語版 Visio のマスター名が異なる場合は、`bpmn2visio.py` の `MASTER_CANDIDATES` を更新してください。

詳細は [Runbook.md](Runbook.md) を参照してください。
