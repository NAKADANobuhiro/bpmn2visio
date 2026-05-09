# bpmn2visio

A tool to automatically convert BPMN 2.0 files (`.bpmn`) to Microsoft Visio diagrams (`.vsdx`).
It uses Visio COM automation (pywin32) with a lane-count-based template approach, placing shapes using genuine BPMN stencil masters (Task, Gateway, Event, etc.) and GlueTo connectors.

> **Note:** Swim lanes cannot be generated dynamically. After conversion, please manually resize, move, or split lanes in Visio as needed.

**日本語版:** [README.ja.md](README.ja.md)

## Features

- Uses genuine Visio BPMN master shapes (Task, Start/End Event, Gateway, etc.)
- Dynamic connectors via GlueTo (connectors follow shapes when moved)
- Template-based approach to bypass the Visio GUARD() lock issue (fixed A3 landscape)
- Supports 2–5 lanes (`templates/lane2.vstx` through `lane5.vstx`)
- Automatic shape placement based on BPMN lane coordinates

## Requirements

- Windows 10/11
- Microsoft Visio 2019 or later (desktop version)
- Python 3.8 or later
- pywin32: `pip install pywin32`

## File Structure

```
bpmn2visio/
├── bpmn2visio.py          # Main conversion script (includes BPMN parser)
├── sample.bpmn             # Sample BPMN (3-lane order processing process)
├── templates/              # Per-lane-count Visio templates (manual creation required)
│   ├── lane2.vstx
│   ├── lane3.vstx
│   ├── lane4.vstx
│   └── lane5.vstx
├── README.md
└── README.ja.md
```

> **Note:** `.vstx` files in `templates/` are binary and should not be committed to Git.
> See "Preparing Templates" below for manual creation steps.

## Quick Start

### 1. Install pywin32

```
pip install pywin32
```

### 2. Prepare templates

Visio template files (`.vstx`) must be created manually in the Visio UI, once per lane count. See [Preparing Templates](#preparing-templates) below.

### 3. Run the conversion

```
python bpmn2visio.py sample.bpmn
```

Output: `sample_com.vsdx` in the same folder as the input file.

### 4. Batch convert multiple files

```
for %f in (*.bpmn) do python bpmn2visio.py %f
```

## Supported BPMN Elements

| BPMN Element | Visio Master |
|---|---|
| startEvent | 開始イベント / Start Event |
| endEvent | 終了イベント / End Event |
| userTask / task | タスク / Task |
| exclusiveGateway | ゲートウェイ / Gateway (XOR) |
| parallelGateway | ゲートウェイ / Gateway (AND) |
| inclusiveGateway | ゲートウェイ / Gateway (OR) |
| subProcess | 展開されたサブプロセス / Sub-Process |
| intermediateCatchEvent | 中間イベント / Intermediate Event |
| sequenceFlow | シーケンス フロー / Sequence Flow |

## Preparing Templates

Each template is a standard Visio BPMN file saved as `.vstx`. Create one file per lane count you need.

1. Open Visio → **File > New > BPMN Diagram**
2. Set page size to A3 landscape: **Design > Page Setup > Paper size: A3 / Orientation: Landscape**
3. Drop a **Pool/Lane** master from the BPMN stencil panel onto the canvas
4. Right-click the Swimlane List and select **Add Lane** until you have the desired number of lanes
5. Expand the pool to fill the page (this cannot be done via COM due to the GUARD() constraint)
6. Save as **File > Save As > Visio Template (*.vstx)** → `templates/lane3.vstx`
7. Repeat for each lane count you need (`lane2.vstx`, `lane4.vstx`, `lane5.vstx`, …)

If no exact match is found, the script will automatically use the closest available template.

## Troubleshooting

**"Template directory not found" error**
→ Create the `templates/` folder next to `bpmn2visio.py` and add the required `.vstx` files.

**"BPMN stencil not found" error**
→ Open a BPMN diagram template in Visio first (File > New > BPMN Diagram), then run the script. The script connects to the running Visio instance and searches for the stencil there.

**"Master not found" warning**
→ Master shape names differ between Visio language versions. To check available names, run this macro in Visio (Alt+F11 → Insert Module → Run):
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
Then update `MASTER_CANDIDATES` in `bpmn2visio.py` to match.

**Lane count mismatch warning**
→ Create a template matching the exact lane count of your BPMN file (`templates/lane{N}.vstx`).

## How It Works

bpmn2visio avoids the Visio GUARD() constraint — which prevents COM automation from resizing BPMN Pool/Lane shapes — by opening a pre-built `.vstx` template that already contains the correct pool/lane structure. The script reads lane geometry from the template, maps BPMN element positions to Visio coordinates using lane-relative scaling, and connects shapes using `GlueTo` for live connector routing.

## License

MIT
