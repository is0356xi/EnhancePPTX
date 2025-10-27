# EnhancePPTX

以下を参考にしながら、あなた自身の内部ツールのpythonを用いて、pptxファイルを作成すること。
下記の例は、一部なので、あなた自身で適切な表現方法を自作しながら作成して良い。

```py
# utils.py
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def pct_to_emu(pct, total_emu):
    """パーセンテージをEMUに変換する"""
    return int(total_emu * (pct / 100.0))


def parse_geom(pos_pct, prs):
    """
    パーセンテージで指定された位置とサイズをEMUに変換する。
    pos: {x, y, w, h} in percentage
    prs: Presentation object
    """
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    return {
        "left": pct_to_emu(pos_pct["x"], slide_width),
        "top": pct_to_emu(pos_pct["y"], slide_height),
        "width": pct_to_emu(pos_pct["w"], slide_width),
        "height": pct_to_emu(pos_pct["h"], slide_height),
    }


def hex_to_rgb(hex_color):
    """HEXカラーコードをRGBタプルに変換する"""
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(r, g, b):
    """RGBタプルをHEXカラーコードに変換する"""
    return "#{:02X}{:02X}{:02X}".format(
        max(0, min(255, int(r))),
        max(0, min(255, int(g))),
        max(0, min(255, int(b))),
    )


def tint_color(hex_color, factor=0.20):
    """色を明るく（factor>0）/暗く（factor<0）する"""
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)

    def _blend(c, toward):
        return max(0, min(255, int(round(c + (toward - c) * abs(factor)))))

    if factor >= 0:
        nr, ng, nb = _blend(r, 255), _blend(g, 255), _blend(b, 255)
    else:
        nr, ng, nb = _blend(r, 0), _blend(g, 0), _blend(b, 0)
    return rgb_to_hex(nr, ng, nb)


def mix_with_white(hex_color, ratio=0.70):
    """色を白と混合してパステル調にする"""
    r, g, b = hex_to_rgb(hex_color)
    r = r + (255 - r) * ratio
    g = g + (255 - g) * ratio
    b = b + (255 - b) * ratio
    return rgb_to_hex(r, g, b)


def mix_rgb_with_white(rgb, ratio_white=0.96):
    """RGBタプルを白と線形補間して明度を上げる"""
    r, g, b = rgb
    rw = int(255 * ratio_white + r * (1 - ratio_white))
    gw = int(255 * ratio_white + g * (1 - ratio_white))
    bw = int(255 * ratio_white + b * (1 - ratio_white))
    return (min(255, rw), min(255, gw), min(255, bw))


def get_theme_color(theme, key, fallback):
    """テーマから色を取得してRGBタプルに変換"""
    raw = (theme or {}).get(key, fallback)
    return normalize_color(raw)


def create_text_box(
    shapes,
    left,
    top,
    width,
    height,
    text,
    font_size=10,
    bold=False,
    color="#000000",
    align=PP_ALIGN.LEFT,
    v_anchor=MSO_ANCHOR.MIDDLE,
):
    """共通のテキストボックス作成ヘルパー"""
    tb = shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = v_anchor
    try:
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = 0
    p.space_after = 0
    p.text = str(text).strip()
    f = p.runs[0].font if p.runs else p.font
    f.size = Pt(font_size)
    f.bold = bold
    r, g, b = normalize_color(color)
    f.color.rgb = RGBColor(r, g, b)
    return tb

def setup_chart_title(chart, title, font_size=16):
    """チャートのタイトルを設定"""
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(font_size)

def normalize_color(color_val):
    """HEX(#RRGGBB) or (r,g,b) or [r,g,b] -> (r,g,b)"""
    if isinstance(color_val, (tuple, list)) and len(color_val) == 3:
        return tuple(int(v) for v in color_val)
    if isinstance(color_val, str):
        return hex_to_rgb(color_val)
    return (0, 0, 0)


def get_contrast_color(hex_bg_color):
    """背景色に基づいてコントラストの高い文字色（黒または白）を返す"""
    r, g, b = hex_to_rgb(hex_bg_color)
    # 輝度を計算 (ITU-R BT.709)
    luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255
    if luminance > 0.5:
        return RGBColor(0, 0, 0)  # 黒
    else:
        return RGBColor(255, 255, 255)  # 白
```

```py
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml import parse_xml

from ..utils import hex_to_rgb
from ..schemas.component_diagram import Schema as ComponentsSchema

NODE_SHAPES = {
    "user": MSO_SHAPE.ROUND_1_RECTANGLE,
    "rect": MSO_SHAPE.RECTANGLE,
}


# ------------------------------------------------------------
# 位置計算ユーティリティ
# ------------------------------------------------------------
def _rel_to_abs(geom, rel):
    left, top, width, height = geom["left"], geom["top"], geom["width"], geom["height"]
    ax = left + int(width * (rel.x / 100.0))
    ay = top + int(height * (rel.y / 100.0))
    aw = int(width * (rel.w / 100.0))
    ah = int(height * (rel.h / 100.0))
    return {"x": ax, "y": ay, "w": aw, "h": ah}


def _calc_conn_points_and_sites(from_g, to_g, margin=0):
    """
    図形 from_g → to_g の位置関係を解析し、
    (start_xy, end_xy), (begin_site, end_site), connector_type を返す。
    site: 0=上,1=左,2=下,3=右
    """
    fx, fy, fw, fh = from_g["x"], from_g["y"], from_g["w"], from_g["h"]
    tx, ty, tw, th = to_g["x"], to_g["y"], to_g["w"], to_g["h"]

    f_left, f_right = fx, fx + fw
    f_top, f_bottom = fy, fy + fh
    t_left, t_right = tx, tx + tw
    t_top, t_bottom = ty, ty + th

    fc_x, fc_y = fx + fw / 2, fy + fh / 2
    tc_x, tc_y = tx + tw / 2, ty + th / 2

    dx = abs(tc_x - fc_x)
    dy = abs(tc_y - fc_y)

    # ほぼ水平かどうかの閾値（図形の高さの平均の半分より縦のズレが小さいか）
    is_horizontal_aligned = dy < (fh + th) / 4
    # ほぼ垂直かどうかの閾値（図形の幅の平均の半分より横のズレが小さいか）
    is_vertical_aligned = dx < (fw + tw) / 4

    # 1) ほぼ水平に並んでいて、かつ左右に分離している場合 → 直線
    if is_horizontal_aligned and (f_right <= t_left or t_right <= f_left):
        connector_type = MSO_CONNECTOR.STRAIGHT
        if f_right <= t_left:  # from -> to
            return (
                ((f_right + margin, fc_y), (t_left - margin, tc_y)),
                (3, 1),
                connector_type,
            )
        else:  # to -> from
            return (
                ((f_left - margin, fc_y), (t_right + margin, tc_y)),
                (1, 3),
                connector_type,
            )

    # 2) ほぼ垂直に並んでいて、かつ上下に分離している場合 → 直線
    if is_vertical_aligned and (f_bottom <= t_top or t_bottom <= f_top):
        connector_type = MSO_CONNECTOR.STRAIGHT
        if f_bottom <= t_top:  # from -> to
            return (
                ((fc_x, f_bottom + margin), (tc_x, t_top - margin)),
                (2, 0),
                connector_type,
            )
        else:  # to -> from
            return (
                ((fc_x, f_top - margin), (tc_x, t_bottom + margin)),
                (0, 2),
                connector_type,
            )

    # 3) 上記以外（斜め、または重なっている）→ 常にカギ型
    connector_type = MSO_CONNECTOR.ELBOW
    if dx >= dy:  # 横方向優勢
        if tc_x > fc_x:  # to が右側
            return ((fc_x, fc_y), (tc_x, tc_y)), (3, 1), connector_type
        else:  # to が左側
            return ((fc_x, fc_y), (tc_x, tc_y)), (1, 3), connector_type
    else:  # 縦方向優勢
        if tc_y > fc_y:  # to が下側
            return ((fc_x, fc_y), (tc_x, tc_y)), (2, 0), connector_type
        else:  # to が上側
            return ((fc_x, fc_y), (tc_x, tc_y)), (0, 2), connector_type


# ------------------------------------------------------------
# コネクタ生成（矢印向き修正版）
# ------------------------------------------------------------
def _add_connector(
    slide,
    start,
    end,
    style,
    connector_type,
    from_shape=None,
    to_shape=None,
    site_pair=None,
):
    conn = slide.shapes.add_connector(
        connector_type, int(start[0]), int(start[1]), int(end[0]), int(end[1])
    )

    if from_shape and to_shape and site_pair:
        try:
            conn.begin_connect(from_shape, site_pair[0])
            conn.end_connect(to_shape, site_pair[1])
        except Exception:
            pass

    line = conn.line
    line.width = Pt(style.pt)
    line.color.rgb = RGBColor(*hex_to_rgb(style.color))
    if style.dash == "dashed":
        line.dash_style = MSO_LINE_DASH_STYLE.DASH

    # Default to end arrow style for now
    ln = line._get_or_add_ln()
    ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ln.append(parse_xml(f'<a:tailEnd type="arrow" {ns}/>'))

    return conn


# ------------------------------------------------------------
# ラベル描画
# ------------------------------------------------------------
def _draw_label(slide, text, center_xy, context):
    w = Pt(100)
    h = Pt(28)
    left = center_xy[0] - int(w / 2)
    top = center_xy[1] - int(h / 2)
    tb = slide.shapes.add_textbox(left, top, w, h)
    tf = tb.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(10)

    default_font_color = context.get("theme", {}).get("font_color", "#000000")
    r.font.color.rgb = RGBColor(*hex_to_rgb(default_font_color))

    return tb


# ------------------------------------------------------------
# メイン描画関数
# ------------------------------------------------------------
def render(slide, data: ComponentsSchema, geom: dict, context: dict):
    node_map = {}
    default_font_color = context.get("theme", {}).get("font_color", "#000000")

    # --- ノード ---
    for nd in data.nodes:
        kind = nd.kind
        if kind not in NODE_SHAPES:
            kind = "rect"

        g = _rel_to_abs(geom, nd.pos)
        shape = slide.shapes.add_shape(
            NODE_SHAPES[kind], g["x"], g["y"], g["w"], g["h"]
        )

        st = nd.style
        fill_hex = (
            st.fill or ("#F0F4FF" if kind == "user" else "#FFFFFF")
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*hex_to_rgb(fill_hex))
        if st.stroke:
            shape.line.color.rgb = RGBColor(*hex_to_rgb(st.stroke))
        shape.line.width = Pt(st.stroke_pt)

        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = nd.label or ""
        r.font.size = Pt(st.font_size or 11)

        font_color_hex = st.font_color or default_font_color
        r.font.color.rgb = RGBColor(*hex_to_rgb(font_color_hex))

        node_map[nd.id] = {"geom": g, "shape": shape}

    # --- コネクタ ---
    for c in data.connectors:
        fn = node_map.get(c.from_)
        tn = node_map.get(c.to)
        if not fn or not tn:
            continue

        (start, end), sites, connector_type = _calc_conn_points_and_sites(
            fn["geom"], tn["geom"]
        )
        conn = _add_connector(
            slide,
            start,
            end,
            c.style,
            connector_type,
            from_shape=fn["shape"],
            to_shape=tn["shape"],
            site_pair=sites,
        )

        if c.label:
            mid = ((start[0] + end[0]) / 2, (start[1] + end[1]) / 2)
            _draw_label(slide, c.label, mid, context)
```

```py
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Pt
from pptx.dml.color import RGBColor

from ..schemas.decompose_boxes import DecomposeBoxesSchema


def render(slide, data: DecomposeBoxesSchema, geom: dict, context: dict):
    """
    決定論的レイアウト：
    - 列は最大深度+1。
    - 各列は等幅、列間ギャップは固定率。
    - 列ごとに上部ヘッダー（テキストボックス、背景なし）。
    - 親高さ = 子合計高さ（兄弟間の垂直ギャップは親にも子にも含めない）。
    - コネクタは描画しない。
    - 色は列0: 薄青 (#DDEBF7) / 列1以降: 薄灰 (#F2F2F2)、枠線なし。
    """
    shapes = context.get("shapes_target", slide.shapes)
    # Normalize root: allow DecomposeBoxNode or list[DecomposeBoxNode];
    # convert pydantic models to plain dicts for renderer convenience.
    raw_root = data.root
    if isinstance(raw_root, list):
        converted = []
        for r in raw_root:
            if hasattr(r, "model_dump"):
                converted.append(r.model_dump())
            else:
                converted.append(r)
        root_data = converted
    else:
        if hasattr(raw_root, "model_dump"):
            root_data = raw_root.model_dump()
        else:
            root_data = raw_root
    headers = data.column_headers

    # -------------------------
    # 既定パラメータ（必要に応じて utils へ移せます）
    # -------------------------
    col_gap_pct = 0.03  # 列間ギャップ（描画領域幅に対する比）
    row_gap_pct = 0.01  # 兄弟間ギャップ（描画領域高に対する比）
    header_band_pct = 0.08  # 列ヘッダー高さ（描画領域高に対する比）
    font_title = Pt(12)
    font_header = Pt(10)

    # 既定色（ご要望の通り）
    LIGHT_BLUE = RGBColor(0xDD, 0xEB, 0xF7)  # #DDEBF7
    LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)  # #F2F2F2

    # -------------------------
    # 深さ＆列数
    # -------------------------
    def depth(node):
        # node may be a dict-like node or a list of nodes (top-level root list)
        if isinstance(node, list):
            if not node:
                return 0
            return max(depth(n) for n in node)
        # guard: if node isn't a mapping-like object, treat as leaf
        if not isinstance(node, dict):
            return 0
        children = node.get("children")
        if not children:
            return 0
        return 1 + max(depth(ch) for ch in children)

    D = depth(root_data)
    cols = D + 1

    # -------------------------
    # 幾何（列幅・位置・コンテンツ帯）
    # -------------------------
    col_gap = int(geom["width"] * col_gap_pct) if cols > 1 else 0
    total_col_gaps = col_gap * (cols - 1)

    # 左側の列（col=0）をデフォルトで狭くする
    if cols > 1:
        # 左側の列を全体の20%、残りの列を等分
        first_col_ratio = 0.2
        remaining_width = geom["width"] - total_col_gaps
        first_col_w = int(remaining_width * first_col_ratio)
        other_col_w = (remaining_width - first_col_w) // (cols - 1)
        col_widths = [first_col_w] + [other_col_w] * (cols - 1)
    else:
        # 列が1つしかない場合は従来通り
        col_w = geom["width"]
        col_widths = [col_w]

    # Only reserve header band if explicit column_headers are provided
    header_h = int(geom["height"] * header_band_pct) if headers else 0
    content_top = geom["top"] + header_h
    content_h = geom["height"] - header_h

    def col_x(c):
        if c == 0:
            return geom["left"]
        else:
            x = geom["left"] + col_widths[0] + col_gap
            for i in range(1, c):
                x += col_widths[i] + col_gap
            return x

    def col_width(c):
        return col_widths[c] if c < len(col_widths) else col_widths[-1]

    # 兄弟間の縦ギャップ（固定）
    gap_y = int(geom["height"] * row_gap_pct)

    # -------------------------
    # 列ヘッダー（背景なしのテキスト）
    # -------------------------
    # Draw column headers only when the user provided them.
    # If the provided `column_headers` list is shorter than the number of
    # columns, do not auto-fill or draw placeholders for the missing headers.
    if headers:
        for c in range(cols):
            if c >= len(headers):
                # No header string supplied for this column — skip drawing.
                continue
            tb = shapes.add_textbox(col_x(c), geom["top"], col_width(c), header_h)
            tf = tb.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = headers[c]
            p.font.size = font_header
            p.font.bold = True

    # -------------------------
    # レイアウト：親→子に高さを分配（合計一致）
    # -------------------------
    boxes = []  # (node, col, x, y, w, h)

    def layout(node, col, x, y, w, h):
        # この node の箱（列 col のコンテンツ帯内）
        y = max(y, content_top)
        bottom = content_top + content_h
        h = min(h, bottom - y)
        boxes.append((node, col, x, y, w, h))

        children = node.get("children") or []
        if not children:
            return

        n = len(children)
        # 重み（省略時は等分）
        weights = [float(ch.get("weight", 1.0)) for ch in children]
        total_w = sum(wi for wi in weights if wi > 0) or float(n)
        # 子同士のギャップを除外した可用高さ
        usable = h - gap_y * (n - 1)
        if usable < 0:
            # ギャップが大きすぎる異常系：ギャップなしで強制配分
            usable = h
            local_gap = 0
        else:
            local_gap = gap_y

        cursor_y = y
        for i, ch in enumerate(children):
            wi = max(weights[i], 0.0)
            frac = (wi / total_w) if total_w > 0 else (1.0 / n)
            # 端数は最後の子に吸収して合計一致
            if i < n - 1:
                ch_h = int(round(usable * frac))
            else:
                ch_h = y + h - cursor_y
            child_col = min(col + 1, cols - 1)
            child_x = col_x(child_col)
            layout(ch, child_col, child_x, cursor_y, col_width(child_col), ch_h)
            cursor_y += ch_h + (local_gap if i < n - 1 else 0)

    # ルートを全高で配置
    # もし root_data がリストで渡されたら、左列 (col=0) に複数のトップレベル箱を分割して配置する
    # （各要素に 'weight' があれば比率配分に使う）
    if isinstance(root_data, list):
        n = len(root_data)
        if n > 0:
            weights = [float(item.get("weight", 1.0)) for item in root_data]
            total_w = sum(w for w in weights if w > 0) or float(n)
            # 子同士のギャップを除いた可用高さ
            usable = content_h - gap_y * (n - 1)
            if usable < 0:
                usable = content_h
                local_gap = 0
            else:
                local_gap = gap_y

            cursor_y = content_top
            for i, item in enumerate(root_data):
                wi = max(weights[i], 0.0)
                frac = (wi / total_w) if total_w > 0 else (1.0 / n)
                # 最後の要素は端数を吸収して合計一致させる
                if i < n - 1:
                    h_i = int(round(usable * frac))
                else:
                    h_i = content_top + content_h - cursor_y
                layout(item, 0, col_x(0), cursor_y, col_width(0), h_i)
                cursor_y += h_i + (local_gap if i < n - 1 else 0)
    else:
        layout(root_data, 0, col_x(0), content_top, col_width(0), content_h)

    # -------------------------
    # 描画：矩形＋テキスト（枠線なし）
    # -------------------------
    def add_box(node, col, x, y, w, h):
        shp = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        shp.fill.solid()
        # 列 0 は薄青、それ以降は薄灰
        shp.fill.fore_color.rgb = LIGHT_BLUE if col == 0 else LIGHT_GRAY
        # 枠線なし
        shp.line.fill.background()
        # 角丸（やりすぎない程度）※環境により 0〜1 相当の調整値
        try:
            shp.adjustments[0] = 0.08
        except Exception:
            pass

        # テキスト（スキーマの 'name' を使用）
        tf = shp.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = node.get("name", "")
        p.font.size = font_title
        p.font.bold = True if col <= 1 else False
        tf.word_wrap = True

    for node, col, x, y, w, h in boxes:
        add_box(node, col, x, y, w, h)
```
