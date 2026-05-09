#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bpmn2visio — BPMN → Visio COM 自動変換スクリプト（テンプレート方式）

事前に Visio UI で作成したレーン数別テンプレート（lane3.vstx 等）を
ベースに開き、BPMN 要素を配置します。
テンプレートは正規の BPMN ステンシルマスターで A3 横に作成されているため、
CFF Container の幅ロック問題を完全に回避できます。

テンプレート構造（inspect_template.py で確認済み）:
  page-level shapes:
    CFF Container  ... プール外枠
    Swimlane List  ... レーン管理（PinX = lane_left = 0.894）
    Pool / Lane    ... Lane1（PinY=9.429, H=3.740）
    Pool / Lane.13 ... Lane2（PinY=5.650, H=3.819）
    Pool / Lane.17 ... Lane3（PinY=2.067, H=3.346）
    Phase List     ... コンテンツ領域管理（PinX = content_left = 1.405）
    Separator      ... 不使用

  各 Pool/Lane 配下:
    Sheet.X  ... 背景矩形
    Sheet.X  ... コネクタポイント
    Sheet.X  ... レーンラベル帯（text = lane_name）

使い方:
  python bpmn2visio.py sample.bpmn

前提条件:
  pip install pywin32
  templates/lane3.vstx（以降 lane4.vstx, lane5.vstx...）が存在すること
  → templates/ フォルダの作成手順は Runbook.md を参照
"""

import sys
import os
import time
import xml.etree.ElementTree as ET

try:
    import win32com.client
    import pywintypes
except ImportError:
    print("エラー: pywin32 が見つかりません。pip install pywin32 を実行してください。")
    sys.exit(1)

# ── BPMN 名前空間 ─────────────────────────────────────────────────────────────
_NS_BPMN   = 'http://www.omg.org/spec/BPMN/20100524/MODEL'
_NS_BPMNDI = 'http://www.omg.org/spec/BPMN/20100524/DI'
_NS_DC     = 'http://www.omg.org/spec/DD/20100524/DC'
_NS_DI     = 'http://www.omg.org/spec/DD/20100524/DI'

PPI = 96.0  # BPMN 座標系: 1インチ = 96ピクセル

# ── テンプレート由来の定数（inspect_template.py で確認済み） ─────────────────
LANE_LABEL_W = 0.511   # レーンラベル帯の幅（インチ）

# ── BPMN タイプ → Visio BPMN ステンシルマスター候補名 ─────────────────────────
# 実機確認済み: 日本語 Visio の BPMN_M.vssx マスター名 (2026-05-02)
MASTER_CANDIDATES = {
    'startEvent':             ['開始イベント',          'Start Event',    'Start'],
    'endEvent':               ['終了イベント',          'End Event',      'End'],
    'userTask':               ['タスク',                'Task',           'User Task'],
    'task':                   ['タスク',                'Task'],
    'serviceTask':            ['タスク',                'Task',           'Service Task'],
    'scriptTask':             ['タスク',                'Task',           'Script Task'],
    'sendTask':               ['タスク',                'Task',           'Send Task'],
    'receiveTask':            ['タスク',                'Task',           'Receive Task'],
    'manualTask':             ['タスク',                'Task',           'Manual Task'],
    'businessRuleTask':       ['タスク',                'Task'],
    'subProcess':             ['展開されたサブプロセス', '折りたたまれたサブプロセス', 'Sub-Process'],
    'callActivity':           ['タスク',                'Task',           'Call Activity'],
    'exclusiveGateway':       ['ゲートウェイ',          'Exclusive Gateway', 'Gateway'],
    'parallelGateway':        ['ゲートウェイ',          'Parallel Gateway',  'Gateway'],
    'inclusiveGateway':       ['ゲートウェイ',          'Inclusive Gateway', 'Gateway'],
    'eventBasedGateway':      ['ゲートウェイ',          'Event Gateway',     'Gateway'],
    'intermediateCatchEvent': ['中間イベント',          'Intermediate Event'],
    'intermediateThrowEvent': ['中間イベント',          'Intermediate Event'],
    'textAnnotation':         ['テキスト注釈',          'Text Annotation'],
}

# コネクタマスター候補（実機確認済み）
CONNECTOR_CANDIDATES = ['シーケンス フロー', 'Sequence Flow', 'Sequence flow']

# ゲートウェイタイプ別 Prop.GatewayType 値
GATEWAY_TYPE_PROP = {
    'exclusiveGateway':  1,   # XOR
    'inclusiveGateway':  2,   # OR
    'parallelGateway':   3,   # AND
    'eventBasedGateway': 4,
}


# ── BPMN パーサー ─────────────────────────────────────────────────────────────

def parse_bpmn(bpmn_path):
    """BPMN XML を解析してデータ構造を返す。"""
    tree = ET.parse(bpmn_path)
    root = tree.getroot()

    def tag(local):
        return f'{{{_NS_BPMN}}}{local}'

    elements  = {}   # id → {type, name}
    flows     = {}   # id → {name, sourceRef, targetRef}
    lanes     = {}   # id → {name, refs: [...]}
    pool_id   = None
    pool_name = ''

    for process in root.iter(tag('process')):
        # レーン
        for laneSet in process.iter(tag('laneSet')):
            for lane in laneSet.iter(tag('lane')):
                lid = lane.get('id', '')
                lanes[lid] = {
                    'name': lane.get('name', lid),
                    'refs': [fn.text for fn in lane.findall(tag('flowNodeRef')) if fn.text],
                }

        # 図形要素
        shape_tags = {
            'startEvent', 'endEvent',
            'task', 'userTask', 'serviceTask', 'scriptTask',
            'sendTask', 'receiveTask', 'manualTask', 'businessRuleTask',
            'subProcess', 'callActivity',
            'exclusiveGateway', 'parallelGateway', 'inclusiveGateway', 'eventBasedGateway',
            'intermediateCatchEvent', 'intermediateThrowEvent',
        }
        for st in shape_tags:
            for el in process.findall(tag(st)):
                eid = el.get('id', '')
                elements[eid] = {
                    'type': st,
                    'name': el.get('name', '').replace('&#10;', '\n'),
                }

        # シーケンスフロー
        for sf in process.findall(tag('sequenceFlow')):
            fid = sf.get('id', '')
            flows[fid] = {
                'name':      sf.get('name', ''),
                'sourceRef': sf.get('sourceRef', ''),
                'targetRef': sf.get('targetRef', ''),
            }

    # プール名
    for collab in root.iter(tag('collaboration')):
        for participant in collab.iter(tag('participant')):
            pool_id   = participant.get('id', '')
            pool_name = participant.get('name', 'Pool')

    # BPMNDI レイアウト情報
    shapes_di = {}   # bpmnElement → {x, y, w, h}
    edges_di  = {}   # bpmnElement → {waypoints: [(x,y),...]}

    for diagram in root.iter(f'{{{_NS_BPMNDI}}}BPMNDiagram'):
        for plane in diagram.iter(f'{{{_NS_BPMNDI}}}BPMNPlane'):
            for shape in plane.findall(f'{{{_NS_BPMNDI}}}BPMNShape'):
                elem_id = shape.get('bpmnElement', '')
                bounds  = shape.find(f'{{{_NS_DC}}}Bounds')
                if bounds is not None:
                    shapes_di[elem_id] = {
                        'x': float(bounds.get('x', 0)),
                        'y': float(bounds.get('y', 0)),
                        'w': float(bounds.get('width', 100)),
                        'h': float(bounds.get('height', 60)),
                    }
            for edge in plane.findall(f'{{{_NS_BPMNDI}}}BPMNEdge'):
                elem_id = edge.get('bpmnElement', '')
                wps = [
                    (float(wp.get('x', 0)), float(wp.get('y', 0)))
                    for wp in edge.findall(f'{{{_NS_DI}}}waypoint')
                ]
                edges_di[elem_id] = {'waypoints': wps}

    return {
        'elements':  elements,
        'flows':     flows,
        'lanes':     lanes,
        'pool_id':   pool_id,
        'pool_name': pool_name,
        'shapes_di': shapes_di,
        'edges_di':  edges_di,
    }


# ── Visio ヘルパー ────────────────────────────────────────────────────────────

def find_master(stencil, candidates):
    for name in candidates:
        try:
            m = stencil.Masters(name)
            if m is not None:
                return m
        except Exception:
            pass
    return None


def find_master_in_docs(visio_app, candidates):
    for doc in visio_app.Documents:
        for name in candidates:
            try:
                m = doc.Masters(name)
                if m is not None:
                    return m, doc
            except Exception:
                pass
    return None, None


def find_lane_shapes(page):
    """ページ直下の Pool/Lane 図形を PinY 降順（上→下）で返す。"""
    lanes = []
    try:
        cnt = page.Shapes.Count
    except Exception:
        return lanes
    for i in range(1, cnt + 1):
        try:
            shp = page.Shapes(i)
            nu  = shp.NameU
            if nu == 'Pool / Lane' or nu.startswith('Pool / Lane.'):
                py = shp.CellsU("PinY").ResultIU
                lanes.append((py, shp))
        except Exception:
            pass
    lanes.sort(key=lambda t: t[0], reverse=True)
    return [shp for _, shp in lanes]


def set_lane_text(lane_shp, text):
    """Pool/Lane 図形本体とラベル帯サブ図形にテキストを設定する。"""
    try:
        lane_shp.Text = text
    except Exception as e:
        print(f"  警告: lane_shp.Text 設定失敗: {e}")
    try:
        for i in range(1, lane_shp.Shapes.Count + 1):
            try:
                sub = lane_shp.Shapes(i)
                if sub.Text and sub.Text.strip():
                    sub.Text = text
            except Exception:
                pass
    except Exception:
        pass


def build_visio_lane_rects(lane_shapes):
    """Pool/Lane 図形のリストから座標辞書リストを構築する（上→下順）。"""
    rects = []
    for shp in lane_shapes:
        px = shp.CellsU("PinX").ResultIU
        py = shp.CellsU("PinY").ResultIU
        w  = shp.CellsU("Width").ResultIU
        h  = shp.CellsU("Height").ResultIU
        left   = px - w / 2
        right  = px + w / 2
        top    = py + h / 2
        bottom = py - h / 2
        rects.append({
            'top':          top,
            'bottom':       bottom,
            'left':         left,
            'right':        right,
            'content_left': left + LANE_LABEL_W,
            'shp':          shp,
        })
    return rects


def select_template(script_dir, num_lanes):
    """レーン数に対応するテンプレートを返す。完全一致なければ最近傍で代替。"""
    tpl_dir = os.path.join(script_dir, "templates")
    if not os.path.isdir(tpl_dir):
        return None, f"テンプレートディレクトリが見つかりません: {tpl_dir}"

    exact = os.path.join(tpl_dir, f"lane{num_lanes}.vstx")
    if os.path.exists(exact):
        return exact, None

    available = []
    for fname in os.listdir(tpl_dir):
        if fname.startswith("lane") and fname.endswith(".vstx"):
            try:
                available.append(int(fname[4:-5]))
            except ValueError:
                pass

    if not available:
        return None, "templates/ にテンプレートが見つかりません"

    closest = min(available, key=lambda n: abs(n - num_lanes))
    path = os.path.join(tpl_dir, f"lane{closest}.vstx")
    msg  = (f"警告: lane{num_lanes}.vstx が未作成。"
            f"lane{closest}.vstx を代替使用します（レーン数が異なる場合があります）。")
    return path, msg


# ── メイン変換処理 ────────────────────────────────────────────────────────────

def create_bpmn_diagram(bpmn_path):
    data      = parse_bpmn(bpmn_path)
    elements  = data['elements']
    flows     = data['flows']
    lanes     = data['lanes']
    shapes_di = data['shapes_di']
    pool_id   = data['pool_id']
    pool_name = data['pool_name'] or 'Pool'

    base        = os.path.splitext(os.path.basename(bpmn_path))[0]
    script_dir  = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(os.path.dirname(os.path.abspath(bpmn_path)),
                               base + '_com.vsdx')

    # ── BPMN プール境界 ────────────────────────────────────────────────────────
    if pool_id and pool_id in shapes_di:
        psd = shapes_di[pool_id]
        pool_bx, pool_by = psd['x'], psd['y']
        pool_bw, pool_bh = psd['w'], psd['h']
    else:
        xs = [v['x']           for v in shapes_di.values()]
        ys = [v['y']           for v in shapes_di.values()]
        xr = [v['x'] + v['w'] for v in shapes_di.values()]
        yb = [v['y'] + v['h'] for v in shapes_di.values()]
        pool_bx = min(xs) - 20
        pool_by = min(ys) - 20
        pool_bw = max(xr) - pool_bx
        pool_bh = max(yb) - pool_by

    # ── 要素 → レーン ID マッピング ───────────────────────────────────────────
    elem_to_lane_id = {}
    for lane_id, linfo in lanes.items():
        for ref in linfo.get('refs', []):
            elem_to_lane_id[ref] = lane_id

    # ── BPMN レーン境界 ───────────────────────────────────────────────────────
    bpmn_lane_bounds = {}
    for lane_id, linfo in lanes.items():
        if lane_id in shapes_di:
            lsd = shapes_di[lane_id]
            bpmn_lane_bounds[lane_id] = {
                'name': linfo['name'],
                'bx': lsd['x'], 'by': lsd['y'],
                'bw': lsd['w'], 'bh': lsd['h'],
            }

    # ── レーンを上→下の順（BPMN Y 座標昇順）にソート ─────────────────────
    def lane_sort_key(lid):
        if lid in shapes_di:
            sd = shapes_di[lid]
            return sd['y'] + sd['h'] / 2
        return float('inf')

    lane_order = sorted(lanes.keys(), key=lane_sort_key)
    lane_names = [lanes[lid]['name'] for lid in lane_order]
    num_lanes  = len(lane_order)

    print(f"BPMN レーン数: {num_lanes}")
    for i, lid in enumerate(lane_order):
        print(f"  [{i}] {lid}: '{lane_names[i]}'")

    # ── テンプレートを選択 ────────────────────────────────────────────────────
    tpl_path, warn = select_template(script_dir, num_lanes)
    if warn:
        print(f"  {warn}")
    if tpl_path is None:
        print(f"エラー: テンプレートが見つかりません。{warn}")
        sys.exit(1)
    print(f"テンプレート: {tpl_path}")

    # ── Visio 起動 / 接続 ─────────────────────────────────────────────────────
    print("Visio に接続中...")
    try:
        visio = win32com.client.GetActiveObject("Visio.Application")
        print("  既存の Visio インスタンスに接続しました。")
    except Exception:
        visio = win32com.client.Dispatch("Visio.Application")
        print("  Visio を起動しました。")

    visio.Visible = True
    visio.AlertResponse = 7

    # ── テンプレートを開く ────────────────────────────────────────────────────
    print("テンプレートを開いています...")
    doc  = visio.Documents.Add(tpl_path)
    page = doc.Pages(1)
    page.Name = base

    try:
        rpw = page.PageSheet.CellsU("PageWidth").ResultIU
        rph = page.PageSheet.CellsU("PageHeight").ResultIU
        print(f"  ページサイズ: {rpw:.3f} x {rph:.3f} in "
              f"({'横' if rpw > rph else '縦'})")
    except Exception:
        pass

    # ── BPMN ステンシルを探す ──────────────────────────────────────────────────
    print("BPMN ステンシルを検索中...")
    stencil = None
    for d in visio.Documents:
        if d.Type == 2 and 'bpmn' in d.Name.lower():
            stencil = d
            print(f"  ステンシル発見: {d.Name}")
            break

    if stencil is None:
        import glob
        search_patterns = [
            r"C:\Program Files\Microsoft Office\root\Office16\Visio Content\*\BPMN_M.vssx",
            r"C:\Program Files\Microsoft Office\root\Office16\Visio Content\BPMN_M.vssx",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\Visio Content\*\BPMN_M.vssx",
            r"C:\Program Files\Microsoft Office\root\Office16\1041\BPMN_M.vssx",
        ]
        stencil_candidates = []
        for pat in search_patterns:
            stencil_candidates.extend(glob.glob(pat))
        if stencil_candidates:
            stencil_path = stencil_candidates[0]
            print(f"  ステンシル発見: {stencil_path}")
            try:
                visio.Documents.OpenEx(stencil_path, 68)
                time.sleep(0.5)
                for d in visio.Documents:
                    if d.Type == 2 and 'bpmn' in d.Name.lower():
                        stencil = d
                        print(f"  ステンシルロード完了: {d.Name}")
                        break
            except Exception as e:
                print(f"  ステンシルのオープン失敗: {e}")

    visio.AlertResponse = 0

    if stencil is None:
        print("エラー: BPMN ステンシルが見つかりません。")
        doc.Close()
        sys.exit(1)

    # ── CFF Container のプール名を更新 ────────────────────────────────────────
    print(f"プールタイトルを '{pool_name}' に設定中...")
    for i in range(1, page.Shapes.Count + 1):
        try:
            shp = page.Shapes(i)
            if shp.NameU == 'CFF Container':
                try:
                    shp.Text = pool_name
                except Exception:
                    pass
                try:
                    for j in range(1, shp.Shapes.Count + 1):
                        try:
                            sub = shp.Shapes(j)
                            if sub.Text and sub.Text.strip():
                                sub.Text = pool_name
                        except Exception:
                            pass
                except Exception:
                    pass
                print("  CFF Container テキスト更新完了")
                break
        except Exception:
            pass

    # ── テンプレートからレーン図形を取得 ─────────────────────────────────────
    print("テンプレートのレーン図形を取得中...")
    tpl_lane_shapes = find_lane_shapes(page)
    print(f"  テンプレートレーン数: {len(tpl_lane_shapes)}")

    if len(tpl_lane_shapes) == 0:
        print("エラー: テンプレートにレーン図形が見つかりません。")
        doc.Close()
        sys.exit(1)

    if len(tpl_lane_shapes) != num_lanes:
        print(f"  警告: BPMNレーン数({num_lanes}) ≠ テンプレートレーン数({len(tpl_lane_shapes)})")
        print("  レーン名の割り当てがずれる可能性があります。")

    # ── レーン名を設定 ────────────────────────────────────────────────────────
    print("レーン名を設定中...")
    for i, (lane_shp, lname) in enumerate(zip(tpl_lane_shapes, lane_names)):
        set_lane_text(lane_shp, lname)
        try:
            px = lane_shp.CellsU("PinX").ResultIU
            py = lane_shp.CellsU("PinY").ResultIU
            w  = lane_shp.CellsU("Width").ResultIU
            h  = lane_shp.CellsU("Height").ResultIU
            print(f"  [{i}] '{lname}' PinX={px:.3f} PinY={py:.3f} W={w:.3f} H={h:.3f}")
        except Exception:
            print(f"  [{i}] '{lname}' (位置読み取り失敗)")

    time.sleep(0.3)

    # ── Visio レーン座標を構築 ────────────────────────────────────────────────
    visio_lane_rects = build_visio_lane_rects(tpl_lane_shapes)

    lane_id_to_rect = {}
    for i, lid in enumerate(lane_order):
        if i < len(visio_lane_rects):
            lane_id_to_rect[lid] = visio_lane_rects[i]

    if visio_lane_rects:
        pool_vis_top          = visio_lane_rects[0]['top']
        pool_vis_bottom       = visio_lane_rects[-1]['bottom']
        pool_vis_right        = visio_lane_rects[0]['right']
        pool_vis_content_left = visio_lane_rects[0]['content_left']
    else:
        pool_vis_top          = 11.0
        pool_vis_bottom       = 0.5
        pool_vis_right        = 16.1
        pool_vis_content_left = 1.4

    # ── 座標変換関数 ──────────────────────────────────────────────────────────
    def to_pin(eid, bx, by, bw, bh):
        """BPMN ピクセル座標（Y下向き） → Visio PinX/PinY（インチ、Y上向き）"""
        native_w = bw / PPI
        native_h = bh / PPI

        lid = elem_to_lane_id.get(eid)
        if lid and lid in bpmn_lane_bounds and lid in lane_id_to_rect:
            bl = bpmn_lane_bounds[lid]
            vl = lane_id_to_rect[lid]
            content_w = vl['right']  - vl['content_left']
            content_h = vl['top']    - vl['bottom']
            rel_x = (bx + bw / 2 - bl['bx']) / bl['bw']
            rel_y = (by + bh / 2 - bl['by']) / bl['bh']
            pin_x = vl['content_left'] + rel_x * content_w
            pin_y = vl['top']          - rel_y * content_h
            return pin_x, pin_y, native_w, native_h

        # フォールバック: プール全体に対して相対配置
        content_w = pool_vis_right - pool_vis_content_left
        content_h = pool_vis_top   - pool_vis_bottom
        rel_x = (bx + bw / 2 - pool_bx) / pool_bw
        rel_y = (by + bh / 2 - pool_by) / pool_bh
        pin_x = pool_vis_content_left + rel_x * content_w
        pin_y = pool_vis_top          - rel_y * content_h
        return pin_x, pin_y, native_w, native_h

    # ── BPMN 図形の配置 ───────────────────────────────────────────────────────
    print("BPMN 図形を配置中...")
    shape_map = {}

    for eid, einfo in elements.items():
        if eid not in shapes_di:
            continue
        sd    = shapes_di[eid]
        pin_x, pin_y, w, h = to_pin(eid, sd['x'], sd['y'], sd['w'], sd['h'])
        etype = einfo['type']
        ename = einfo['name']

        master = find_master(stencil, MASTER_CANDIDATES.get(etype, ['タスク', 'Task']))
        if master is None:
            print(f"  警告: マスターが見つかりません ({etype}: {eid})")
            continue

        try:
            shp = page.Drop(master, pin_x, pin_y)
            shp.CellsU("Width").FormulaU  = f"{w:.4f} in"
            shp.CellsU("Height").FormulaU = f"{h:.4f} in"
            if ename:
                shp.Text = ename.replace('\n', chr(10))
            if etype in GATEWAY_TYPE_PROP:
                try:
                    shp.CellsU("Prop.GatewayType").FormulaU = str(GATEWAY_TYPE_PROP[etype])
                except Exception:
                    pass
            shape_map[eid] = shp
            lid   = elem_to_lane_id.get(eid, '?')
            lname = lanes[lid]['name'] if lid in lanes else '?'
            print(f"  配置: {eid} ({etype}) レーン='{lname}' @ ({pin_x:.2f}, {pin_y:.2f})")
        except Exception as ex:
            print(f"  警告: 図形配置失敗 ({eid}): {ex}")

    # ── シーケンスフロー ──────────────────────────────────────────────────────
    print("シーケンスフローを接続中...")
    seq_master = find_master(stencil, CONNECTOR_CANDIDATES)
    if seq_master is None:
        seq_master, _ = find_master_in_docs(visio, ["Dynamic connector", "動的コネクタ"])

    for fid, finfo in flows.items():
        src_id = finfo['sourceRef']
        tgt_id = finfo['targetRef']
        fname  = finfo.get('name', '')

        if src_id not in shape_map or tgt_id not in shape_map:
            print(f"  スキップ: {fid} (図形未作成)")
            continue

        src_shp = shape_map[src_id]
        tgt_shp = shape_map[tgt_id]

        try:
            conn = page.Drop(seq_master, 0, 0)
            conn.CellsU("BeginX").GlueTo(src_shp.CellsU("PinX"))
            conn.CellsU("BeginY").GlueTo(src_shp.CellsU("PinY"))
            conn.CellsU("EndX").GlueTo(tgt_shp.CellsU("PinX"))
            conn.CellsU("EndY").GlueTo(tgt_shp.CellsU("PinY"))
            if fname:
                conn.Text = fname
            print(f"  接続: {src_id} -> {tgt_id}" + (f" [{fname}]" if fname else ""))
        except Exception as ex:
            print(f"  警告: 接続失敗 ({fid}): {ex}")

    # ── ページに合わせて表示・保存 ───────────────────────────────────────────
    try:
        visio.ActiveWindow.ViewFit = 1
    except Exception:
        pass

    print(f"\n保存中: {output_path}")
    try:
        doc.SaveAs(output_path)
        print(f"完了: {os.path.basename(output_path)}")
        print(f"  レーン数: {num_lanes}")
        print(f"  図形数: {len(shape_map)}")
        print(f"  フロー数: {len(flows)}")
    except Exception as e:
        print(f"エラー: 保存失敗: {e}")

    return output_path


# ── エントリポイント ──────────────────────────────────────────────────────────
if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(description='BPMN → Visio COM 変換（テンプレート方式）')
    ap.add_argument('bpmn', help='BPMN ファイルのパス')
    args = ap.parse_args()
    create_bpmn_diagram(args.bpmn)
