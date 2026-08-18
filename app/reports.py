# -*- coding: utf-8 -*-
"""报表引擎：模板解析（行=实体×字段，列=期间组×口径）→ 数值表格 → Excel 导出

模板两种列结构：
- 旧式 "columns"：列 = 字段×口径×期间（月度报表用）
- 新式 "column_blocks"：列 = 期间组×口径，字段在行（经营预测3 用，与原表一致）
  行 = rows(实体) × fields(字段) 展开，每实体连续 3 行共享 A 列标签

期间表达式：'7月实际' / '1-7月实际' / '8-12月预测' / '1-7月实际+8-12月预测' / '当月' / '1-当月'
占位符：{y} 年份、{m} 查询月、{m+1}/{m+2} 查询月+N、{k} 月份循环块内当前月
"""
import json
import math
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import calc

PERIOD_SEG_RE = re.compile(r"(\d{1,2})(?:-(\d{1,2}))?月(实际|BP|预测|同期)")


def parse_period(expr, report_month):
    """期间表达式 -> [(m1, m2, dtype名)]（相对项以报表月为基准）"""
    segments = []
    for part in str(expr).split("+"):
        part = part.strip()
        m_rel = re.match(r"^(1-当月|当月)(实际|BP|预测|同期)?$", part)
        if m_rel:
            dtype = m_rel.group(2) or "实际"
            if m_rel.group(1) == "当月":
                segments.append((report_month, report_month, dtype))
            else:
                segments.append((1, report_month, dtype))
            continue
        m = PERIOD_SEG_RE.match(part)
        if not m:
            raise ValueError(f"期间表达式无法解析: {part!r} (来自 {expr!r})")
        m1 = int(m.group(1))
        m2 = int(m.group(2) or m.group(1))
        if m1 > 12:
            continue  # 超范围段（如 12月查询时 {m+1}-12月 = 13-12月）→ 无此期间，丢弃
        if m1 > m2:
            raise ValueError(f"期间范围错误: {part!r}")
        segments.append((m1, m2, m.group(3)))
    return segments


def _subst(s, year, month, extra=None):
    """模板占位符替换：{y}/{m}/{m+1}/{m+2} + 块级占位（{k}/{q}/{qs}/{qe}/{h}/{hs}/{he}/{h_label}）"""
    s = (str(s).replace("{y}", str(year)).replace("{m}", str(month))
         .replace("{m+1}", str(month + 1)).replace("{m+2}", str(month + 2))
         .replace("{m_q}", str((month - 1) // 3 + 1))   # 查询月所在季度
         .replace("{m_h}", "1" if month <= 6 else "2"))  # 查询月所在半年
    if extra:
        for k, v in extra.items():
            s = s.replace("{" + k + "}", str(v))
    return s


def resolve_selector(conn, sel):
    """选择器 -> 节点 id 列表"""
    if sel.get("all"):
        roots = [r["id"] for r in conn.execute(
            "SELECT cn.id FROM channel_nodes cn JOIN product_lines p ON p.id=cn.product_id "
            "WHERE cn.parent_id IS NULL")]
        if "path" in sel:
            resolved = []
            for nid in roots:
                cur = nid
                for seg in sel["path"]:
                    row = conn.execute(
                        "SELECT id FROM channel_nodes WHERE parent_id=? AND name=?",
                        (cur, seg)).fetchone()
                    if row is None:
                        cur = None
                        break
                    cur = row["id"]
                if cur is not None:
                    resolved.append(cur)
            return resolved
        return roots
    nid = calc.resolve_path(conn, sel["p"], tuple(sel.get("path", [])))
    if nid is None:
        return []
    if sel.get("leaf"):
        return calc.subtree_leaves(conn, nid)
    return [nid]


def resolve_rows(conn, row_defs):
    """旧式行定义（实体展开）：-> [(显示名, 节点id列表)]"""
    rows = []
    for rd in row_defs:
        if "expand" in rd:
            sel = rd["expand"]
            nids = resolve_selector(conn, sel)
            for nid in nids:
                path = calc.node_path(conn, nid)
                name = rd.get("name_template", "{path}").format(path="-".join(path))
                rows.append((name, [nid]))
        else:
            nids = resolve_selector(conn, rd["sel"])
            if nids:
                rows.append((rd["name"], nids))
    return rows


# 自然季度/半年度的首末月映射（集中魔法数）
_Q_MONTHS = {1: (1, 3), 2: (4, 6), 3: (7, 9), 4: (10, 12)}
_H_MONTHS = {1: (1, 6), 2: (7, 12)}
_H_LABELS = {1: "上半年", 2: "下半年"}


def _block_pairs(blk, year, month):
    """块定义 -> [(group, period), ...]：各类块的期间展开统一出口"""
    title = _subst(blk["title"], year, month)
    period_tpl = _subst(blk["period"], year, month)
    if "months" in blk:  # 单月组：{k} 从 start 到 end
        start = int(_subst(blk["months"]["start"], year, month))
        return [(title.replace("{k}", str(k)), period_tpl.replace("{k}", str(k)))
                for k in range(start, blk["months"]["end"] + 1)]
    if "quarters" in blk:  # 季度组：{q} 从 start（{m_q}=查询月所在季度）到 Q4
        start_q = int(_subst(blk["quarters"]["start"], year, month))
        return [(_subst(title, year, month, {"q": q, "qs": _Q_MONTHS[q][0], "qe": _Q_MONTHS[q][1]}),
                 _subst(period_tpl, year, month, {"q": q, "qs": _Q_MONTHS[q][0], "qe": _Q_MONTHS[q][1]}))
                for q in range(start_q, 5)]
    if "halves" in blk:  # 半年组：{h} 从 start（{m_h}=查询月所在半年）起
        start_h = int(_subst(blk["halves"]["start"], year, month))
        pairs = []
        for h in range(start_h, 3):
            ex = {"h": h, "hs": _H_MONTHS[h][0], "he": _H_MONTHS[h][1], "h_label": _H_LABELS[h]}
            pairs.append((_subst(title, year, month, ex), _subst(period_tpl, year, month, ex)))
        return pairs
    if "segments" in blk:  # 任意显式段列表：[{"title":..., "period":...}]
        return [(_subst(seg["title"], year, month), _subst(seg["period"], year, month))
                for seg in blk["segments"]]
    if not parse_period(period_tpl, month):
        return []  # 期间为空（12月查询的 {m+1}-12月）→ 该块不存在，跳过
    return [(title, period_tpl)]


def _expand_columns(conn, cfg, year, month):
    """列展开：column_blocks（新式，字段在行）或 columns（旧式）-> [{group, measure, period, measure_row}]
    块类型：months（单月循环 {k}）/ quarters（自然季度循环 {q}{qs}{qe}）/ halves（半年循环 {h}{hs}{he}{h_label}）
            / segments（显式段列表）/ channels（渠道在列，每渠道一列）/ 静态块"""
    cols = []
    if "column_blocks" in cfg:
        for blk in cfg["column_blocks"]:
            if "channels" in blk:
                # 渠道在列：每渠道 × 每口径 一列（列带渠道名与节点集；行取数按渠道）
                period = _subst(blk["period"], year, month)
                for ch in blk["channels"]:
                    nids = resolve_selector(conn, ch["sel"])
                    if not nids:
                        continue
                    for mea in blk["measures"]:
                        cols.append(_mk_col(conn, blk["title"], period,
                                            mea, channel_name=ch["name"], channel_nids=nids))
                continue
            for group, period in _block_pairs(blk, year, month):
                for mea in blk["measures"]:
                    cols.append(_mk_col(conn, group, period, mea))
    else:
        for c in cfg["columns"]:
            cols.append(_mk_col(conn, _subst(c.get("group", c["field"]), year, month),
                                _subst(c["period"], year, month), c["measure"], c["field"]))
    return cols


def _mk_col(conn, group, period_expr, measure_name, field=None, channel_name=None, channel_nids=None):
    """单列：期间 + 口径（新式字段在行；旧式带 field；渠道列带 channel 信息）"""
    measure = conn.execute("SELECT * FROM measures WHERE name=?", (measure_name,)).fetchone()
    if measure is None:
        raise ValueError(f"未知口径: {measure_name}")
    return {"group": group, "measure": measure_name, "period": period_expr,
            "measure_row": measure, "field": field,
            "channel_name": channel_name, "channel_nids": channel_nids}


def _drill_for(conn, segments, measure_row, year):
    """单元格钻取条件：{note, segments:[{year, months:[m1,m2], dtype_ids}]}（前端双击跳明细用）
    复合期间按段拆分（全年累计实际 = 1-7实际 + 8-12预测 两段，各按段自身 dtype）；
    formula 口径变量含虚类型（同期）→ 去年实际段（同比不丢分母）；同年月段合并类型。
    空期间（12月查询的 {m+1}-12月 等）→ None（该列无数据，不可钻取）"""
    if not segments:
        return None
    m1 = min(s[0] for s in segments)
    m2 = max(s[1] for s in segments)
    dtypes = {r["name"]: r for r in conn.execute(
        "SELECT id, name, base_dtype_id, year_shift FROM data_types")}

    def merge(pairs):
        """[(year, [m1,m2], [dtype_ids])] -> 同年月段合并类型去重"""
        merged = {}
        for y, months, ids in pairs:
            key = (y, months[0], months[1])
            merged.setdefault(key, []).extend(ids)
        return [{"year": y, "months": [m1_, m2_], "dtype_ids": list(dict.fromkeys(ids))}
                for (y, m1_, m2_), ids in sorted(merged.items())]

    if measure_row["kind"] == "dtype":
        if measure_row["year_shift"]:  # 同期口径（实际 + year_shift=-1）：钻取去年实际
            return {"note": "同期 = 去年实际",
                    "segments": merge([(year + measure_row["year_shift"], [m1, m2],
                                        [measure_row["dtype_id"]])])}
        # 实际口径：复合期间每段按段自身 dtype（1-7实际+8-12预测 -> 两段）；
        # 其余口径（BP/预测）：段名可能是默认「实际」（period 无后缀），统一用口径自身的 dtype_id
        if measure_row["name"] == "实际":
            pairs = [(year, [s[0], s[1]],
                      [dtypes[s[2]]["id"] if s[2] in dtypes else measure_row["dtype_id"]]) for s in segments]
        else:
            pairs = [(year, [s[0], s[1]], [measure_row["dtype_id"]]) for s in segments]
        return {"note": measure_row["name"], "segments": merge(pairs)}
    # formula 口径：变量 = 数据类型名；虚类型（同期）→ 去年实际段；变量=实际 → 复合期间按段拆分
    notes, pairs = [], []
    for name in re.findall(r"[一-鿿A-Za-z0-9_]+", measure_row["formula"] or ""):
        d = dtypes.get(name)
        if d is None or name in notes:
            continue
        notes.append(name)
        if d["base_dtype_id"] is not None:  # 虚类型（同期 = 去年实际）
            pairs.append((year + (d["year_shift"] or 0), [m1, m2], [d["base_dtype_id"]]))
        elif name == "实际":  # 混合上下文：按段自身 dtype 拆分
            for s in segments:
                seg = dtypes.get(s[2])
                pairs.append((year, [s[0], s[1]], [seg["id"] if seg else d["id"]]))
        else:
            pairs.append((year, [m1, m2], [d["id"]]))
    return {"note": f"{measure_row['name']}（{' / '.join(notes)}）", "segments": merge(pairs)}


def layout(conn, template_row, year, month):
    """报表结构（不算单元格值）：{columns, rows(骨架), title}
    列带 segments/drill/channel_nids（渠道在列时），行骨架带 entity/field/label/nids/measure_overrides。
    render 在其上填充 cells；explain 接口只取结构，免整表重算。
    measure_overrides 键 = 列索引（int）：进程内 Python 直取；经 JSON 序列化后前端按字符串访问；
    两处取值需保持同一索引语义（勿改键类型）"""
    cfg = json.loads(template_row["config_json"])
    cols = _expand_columns(conn, cfg, year, month)
    for c in cols:
        c["segments"] = parse_period(c["period"], month)
        c["drill"] = _drill_for(conn, c["segments"], c["measure_row"], year)
    out_rows = []
    field_ids = {r["name"]: r["id"] for r in conn.execute("SELECT id, name FROM fields")}
    measures_by_name = {r["name"]: r for r in conn.execute("SELECT * FROM measures")}

    def field_of(fdef):
        """字段定义 -> (fid, name, label)；未知字段返回 None"""
        fname = fdef["name"] if isinstance(fdef, dict) else fdef
        label = fdef.get("label", fname) if isinstance(fdef, dict) else fname
        fid = field_ids.get(fname)
        return None if fid is None else (fid, fname, label)

    def row_overrides(fdef, cols):
        """行级口径覆盖：{列索引: 覆盖口径名}；口径名校验（未知报错）"""
        overrides = {}
        mmap = fdef.get("measure_map") if isinstance(fdef, dict) else None
        if not mmap:
            return overrides
        for ci, c in enumerate(cols):
            if c["measure"] in mmap:
                mea = mmap[c["measure"]]
                if measures_by_name.get(mea) is None:
                    raise ValueError(f"未知口径: {mea}")
                overrides[ci] = mea
        return overrides

    if "row_groups" in cfg:
        # 渠道在列模式：行 = 分组 × 字段（A 列=分组标签），列 = 渠道×口径（列带 channel_nids）
        for grp in cfg["row_groups"]:
            for fdef in grp["fields"]:
                fo = field_of(fdef)
                if fo is None:
                    continue
                fid, fname, label = fo
                for c in cols:
                    if c.get("channel_nids") is None:
                        raise ValueError(f"row_groups 模板要求全部列带 channels: {fname!r}")
                out_rows.append({"entity": grp["label"], "field": fname, "label": label,
                                 "measure_overrides": row_overrides(fdef, cols)})
    elif "fields" in cfg and "rows" in cfg:
        # 新式：行 = 实体 × 字段 展开（每实体连续 N 行，共享实体标签；label=字段显示名）
        for entity, nids in resolve_rows(conn, cfg["rows"]):
            for fdef in cfg["fields"]:
                fo = field_of(fdef)
                if fo is None:
                    continue
                fid, fname, label = fo
                out_rows.append({"entity": entity, "field": fname, "label": label,
                                 "nids": nids, "measure_overrides": row_overrides(fdef, cols)})
    else:
        # 旧式：行 = 单字段实体（月度报表）；列字段未知/缺失时跳过该列（与新式 field_of 优雅跳过一致）
        cols = [c for c in cols if field_ids.get(c.get("field") or "")]
        for name, nids in resolve_rows(conn, cfg["rows"]):
            out_rows.append({"entity": name, "field": None, "label": None, "nids": nids})
    return {"columns": cols, "rows": out_rows, "title": cfg.get("title", template_row["name"])}


def render(conn, template_row, year, month):
    """渲染报表：layout 结构 + 填充单元格值。列带 segments/drill/channel_nids（渠道在列时），
    行带 nids（实体在行时）与 measure_overrides（行级口径覆盖）——供显示/钻取/计算过程"""
    t = layout(conn, template_row, year, month)
    cols, rows = t["columns"], t["rows"]
    field_ids = {r["name"]: r["id"] for r in conn.execute("SELECT id, name FROM fields")}
    measures_by_name = {r["name"]: r for r in conn.execute("SELECT * FROM measures")}
    for row in rows:
        fname = row.get("field")
        fid = field_ids.get(fname) if fname else None
        cells = []
        for ci, c in enumerate(cols):
            mea = (row.get("measure_overrides") or {}).get(ci) or c["measure"]
            mrow = measures_by_name.get(mea)  # 覆盖口径必已校验存在
            nids = row.get("nids") or c.get("channel_nids")
            if fid is None:
                # 旧式：字段在列上
                fid2 = field_ids[c["field"]]
                cells.append(calc.period_measure(conn, year, c["segments"], nids, fid2, mrow))
            else:
                cells.append(calc.period_measure(conn, year, c["segments"], nids, fid, mrow))
        row["cells"] = cells
    return t


def api_columns(cols):
    """渲染列 -> API 序列化。渠道在列时透传 channel_name/channel_nids（前端表头显示渠道名），
    其余字段（group/field/measure/period/drill）原样；此前遗漏 channel_name 导致费用分析表头丢失渠道名"""
    out = []
    for c in cols:
        col = {"group": c.get("group"), "field": c["field"], "measure": c["measure"],
               "period": c["period"], "drill": c.get("drill")}
        if c.get("channel_name") is not None:
            col["channel_name"] = c["channel_name"]
        if c.get("channel_nids") is not None:
            col["channel_nids"] = c["channel_nids"]
        out.append(col)
    return out


# ---------- Excel 导出（样式规格模式） ----------
_HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
_FONT = Font(name="黑体", size=8)
_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _numfmt(unit, measure_name):
    if measure_name in ("BP达成率", "同比") or (unit or "") == "%":
        return "0.0%"  # 比率类口径 + 比率字段（毛利率/费用率/利润率）用百分比格式
    scale = {"万元": 0, "万台": 0, "元": 0}.get(unit or "", 2)
    return f"0.{'0' * scale}" if scale else "0"


def _measure_label(name):
    """列口径显示名：差值同比类（ASP同比/毛利率同比/净利同比）统一显示「同比」，与收入/销量块一致"""
    return "同比" if name.endswith("同比") else name


def _round_away(x, nd=0):
    """半数远离零取整（Excel ROUND / JS Math.round 语义；Python round 为银行家舍入，边界会差 1）"""
    mul = 10 ** nd
    y = x * mul
    return (math.floor(y + 0.5) if y >= 0 else math.ceil(y - 0.5)) / mul


def _margin_yoy_text(v):
    """毛利率同比文本：0.032 -> 增加3.2个点（原表 26年累计报表 U 列公式）"""
    d = abs(_round_away(v * 100, 1))
    s = f"{d:g}"
    return f"增加{s}个点" if v >= 0 else f"减少{s}个点"


def _profit_yoy_text(v, yoy):
    """净利同比文本（原表 Z 列公式）：四态 增亏/减亏/增盈/减盈 + |差|万
    按今年盈亏状态 + 差值方向；无去年同期值（yoy=None）退化按原表 IF(D>=0,减亏,增亏)"""
    d = _round_away(v)
    if yoy is None or yoy + v < 0:  # 今年亏损（或未知）
        verb = "减亏" if d >= 0 else "增亏"
    else:  # 今年盈利
        verb = "增盈" if d >= 0 else "减盈"
    return f"{verb}{abs(d):g}万"


def export_report(conn, template_row, year, month, out_path):
    """渲染并导出 Excel（两行表头：期间组头合并 + 口径行；A 列实体标签跨字段行合并，B 列=项目）"""
    table = render(conn, template_row, year, month)
    cols, rows = table["columns"], table["rows"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}年{month}月报表"

    ws.column_dimensions["A"].width = 14  # 实体
    ws.column_dimensions["B"].width = 12  # 项目（字段显示名）
    for i, c in enumerate(cols):
        # 渠道列（经营费用分析）：列名长，加宽
        ws.column_dimensions[get_column_letter(i + 3)].width = 11 if c.get("channel_name") else 8.5

    # 表头：A=品类(合并) B=项目(合并) 数据从 C 起
    ws["A1"] = table["title"]
    ws.merge_cells("A1:A2")
    ws["B1"] = "项目"
    ws.merge_cells("B1:B2")
    g_start = 3
    for i, c in enumerate(cols):
        col = i + 3
        c1 = ws.cell(row=1, column=col, value=c["group"])
        # 渠道在列：行 2 显示渠道名（非口径名）
        c2 = ws.cell(row=2, column=col, value=c.get("channel_name") or _measure_label(c["measure"]))
        for cell in (c1, c2):
            cell.font = _FONT
            cell.alignment = _CENTER
            cell.border = _BORDER
        if i > 0 and c["group"] != cols[i - 1]["group"]:
            if col - 1 > g_start:
                ws.merge_cells(start_row=1, start_column=g_start, end_row=1, end_column=col - 1)
            g_start = col
    if len(cols) + 2 > g_start:
        ws.merge_cells(start_row=1, start_column=g_start, end_row=1, end_column=len(cols) + 2)
    for col in range(1, len(cols) + 3):
        cell = ws.cell(row=1, column=col)
        if col > 2:
            cell.fill = _HEADER_FILL
        cell.font = _FONT
        cell.alignment = _CENTER
        cell.border = _BORDER
        ws.cell(row=2, column=col).fill = _HEADER_FILL

    # 数据行：A=实体 B=项目 数据从 C 起；字段单位按行取
    unit_by_field = {r["name"]: r["unit"] for r in conn.execute("SELECT name, unit FROM fields")}
    # 净利同比列 -> 同块（group）同期列（盈亏文本需去年同期判断今年盈亏状态）。
    # 含「同比」列：fields 模式（预测3）列名仍为 同比，行级覆盖后才成净利同比，
    # 该列同样需要同期值（收入行百分比同比不使用该映射，无害）
    yoy_col = {}
    for i, c in enumerate(cols):
        if c["measure"] in ("净利同比", "同比"):
            for j, c2 in enumerate(cols):
                if j != i and c2.get("group") == c.get("group") and c2["measure"] == "同期":
                    yoy_col[i] = j
                    break
    for r, row in enumerate(rows, start=3):
        ws.cell(row=r, column=1, value=row["entity"])
        ws.cell(row=r, column=2, value=row.get("label") or "")
        for i, v in enumerate(row["cells"]):
            col = i + 3
            c = cols[i]
            cell = ws.cell(row=r, column=col)
            # 行级口径覆盖后的实际口径：毛利率/净利 同比 按原表写文本（个点差 / 盈亏万差）
            mea = (row.get("measure_overrides") or {}).get(i) or c["measure"]
            if mea == "毛利率同比":
                cell.value = "" if v is None else _margin_yoy_text(v)
            elif mea == "净利同比":
                yoy = row["cells"][yoy_col[i]] if i in yoy_col else None
                cell.value = "" if v is None else _profit_yoy_text(v, yoy)
            else:
                # 无数据写空串（与文本分支一致，不伪装成 0）；单位取行字段，旧式模板行无字段时回退列字段
                cell.value = "" if v is None else round(v, 6)
                cell.number_format = _numfmt(unit_by_field.get(row.get("field") or c.get("field")),
                                             c["measure"])
            cell.font = _FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = _BORDER
        for cc in (1, 2):
            ws.cell(row=r, column=cc).font = _FONT
            ws.cell(row=r, column=cc).alignment = _CENTER
            ws.cell(row=r, column=cc).border = _BORDER
        ws.row_dimensions[r].height = 13
    # A 列：同实体连续行合并标签
    start = 3
    for i in range(1, len(rows)):
        if rows[i]["entity"] != rows[i - 1]["entity"]:
            if i + 2 > start:
                ws.merge_cells(start_row=start, start_column=1, end_row=i + 2, end_column=1)
            start = i + 3
    if len(rows) + 2 > start:
        ws.merge_cells(start_row=start, start_column=1, end_row=len(rows) + 2, end_column=1)

    ws.row_dimensions[1].height = 13
    ws.row_dimensions[2].height = 23
    ws.freeze_panes = "C3"
    wb.save(out_path)
    return out_path
