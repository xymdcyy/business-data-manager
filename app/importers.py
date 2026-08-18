# -*- coding: utf-8 -*-
"""导入引擎：宽表/明细两类来源解析 → 预览 → 入库（upsert）

- 宽表（wide）：主文件 26年实际/26年预算/25实际，行2'合计'分块、行1块标签（KA-智屏-连锁-连锁）、A列字段名、块内1..12月
  · 聚合块（非叶）不写事实，用于校验：对比 叶子和 vs 块值（差值>0.01 报警）
  · 月份标签与位置不符时报警（防"源表忘了改月份名"类错误），但按位置导入
- 明细（detail）：品类源库长表，行=年月×品类×渠道，列=指标
"""
import json
import math

import openpyxl
from openpyxl.utils import column_index_from_string

from . import calc
from .seed import PRODUCT_ALIASES

MONTH_LABELS = [f"{m}月" for m in range(1, 13)]


class Resolver:
    """导入期元数据缓存（字段/数据类型名 -> id）"""

    def __init__(self, conn):
        self.field_id = {r["name"]: r["id"] for r in conn.execute("SELECT id, name FROM fields")}
        self.dtype_id = {r["name"]: r["id"] for r in conn.execute("SELECT id, name FROM data_types")}
        self.field_kind = {r["name"]: r["kind"] for r in conn.execute("SELECT name, kind FROM fields")}
        self.field_formula = {r["name"]: r["formula"] for r in conn.execute("SELECT name, formula FROM fields")}
        self.node = {}
        for r in conn.execute(
            "SELECT cn.id, p.name AS prod, cn.name, cn.parent_id, cn.is_leaf "
            "FROM channel_nodes cn JOIN product_lines p ON p.id = cn.product_id"
        ):
            self.node[r["id"]] = dict(r)


def _resolve_label(conn, label):
    """块标签（KA-智屏-连锁-连锁）-> node_id；跨产品聚合/未知返回 None"""
    if not label or not isinstance(label, str) or not label.startswith("KA-"):
        return None
    parts = [p.strip() for p in label.split("-")[1:]]
    if not parts or parts[0] not in PRODUCT_ALIASES:
        return None
    return calc.resolve_path(conn, PRODUCT_ALIASES[parts[0]], tuple(parts[1:]))


def parse_wide(conn, cfg, file_path, sheet_name):
    """宽表解析：返回 (items, warnings)；items=(year,month,node_id,field_id,dtype_id,value)"""
    res = Resolver(conn)
    year = cfg["year"]
    dtype_id = res.dtype_id[cfg["dtype"]]
    warnings = []

    wb = openpyxl.load_workbook(file_path, data_only=True)  # 源表都很小，普通模式更快
    ws = wb[sheet_name]
    starts = [c for c in range(2, ws.max_column + 1) if ws.cell(row=2, column=c).value == "合计"]
    if not starts:
        wb.close()
        raise ValueError(f"工作表 {sheet_name} 未找到'合计'分块")

    items = []
    leaf_by_node = {}          # node_id -> [(field_id, month, value)] 聚合校验用
    agg_blocks = []            # (node_id, 块起始列) 校验用
    # 受保护单元格：产品 -> 月份集合（如 智屏 1-6月 以明细源为权威，宽表重导入不覆盖）
    protected = {pr["product"]: set(pr.get("months", [])) for pr in cfg.get("protected", [])}
    protected_skipped = 0
    for s in starts:
        label = ws.cell(row=1, column=s).value
        label = cfg.get("label_fix", {}).get(str(s), label)  # 源表标签纠错（按块起始列）
        node_id = _resolve_label(conn, label)
        if node_id is None:
            warnings.append(f"列{s} 块标签 {label!r}: 未匹配到渠道树节点（跨产品聚合/未知，跳过）")
            continue
        is_leaf = bool(res.node[node_id]["is_leaf"])
        # 月份列 -> 值。默认按位置导入（数据位置为准，标签不符仅报警——防"忘了改月份名"）；
        # month_mode=label 的模板（如规划2 同期表：位置错位、标签为真）按标签导入
        col_month = {}
        if cfg.get("month_mode") == "label":
            for off in range(1, 13):
                mc = s + off
                hdr = ws.cell(row=2, column=mc).value
                if not isinstance(hdr, str) or "月" not in hdr:
                    continue
                m = int(hdr.replace("月", "").strip())
                if not (1 <= m <= 12):
                    continue
                if m in col_month:
                    # 同块重复月份标签（复制粘贴类缺陷）：后列覆盖前列，与入库 upsert 结果一致
                    warnings.append(f"列{s} 块 {label} 月份标签重复 {hdr!r}（同月多列，后写覆盖）")
                col_month[m] = mc
        else:
            for off in range(1, 13):
                mc = s + off
                hdr = ws.cell(row=2, column=mc).value
                label_ok = hdr is not None and str(hdr).strip() == MONTH_LABELS[off - 1]
                if not label_ok:
                    warnings.append(f"列{s} 块 {label} 列{mc} 月份标签 {hdr!r} != {MONTH_LABELS[off-1]}（按位置导入）")
                col_month[off] = mc  # 位置即月份
        for month, mc in sorted(col_month.items()):
            for r in range(3, ws.max_row + 1):
                name = ws.cell(row=r, column=1).value
                if not name or not isinstance(name, str):
                    continue
                field_name = name.strip().lstrip("+")  # A 列 '+后台利润' 等带符号前缀
                if field_name not in res.field_id:
                    continue
                v = ws.cell(row=r, column=mc).value
                if not isinstance(v, (int, float)):
                    continue
                if not math.isfinite(v):
                    warnings.append(f"列{s} 块 {label} 行{r} {name}: 单元格值 {v!r} 非有限数，跳过")
                    continue  # NaN/Inf 入库会导致报表/明细序列化 500
                if res.field_kind.get(field_name) != "raw":
                    continue  # 公式字段不落库（报表/网格统一按公式重算；源值仅用于校验）
                if is_leaf:
                    if month in protected.get(res.node[node_id]["prod"], ()):
                        protected_skipped += 1  # 明细源权威，宽表不覆盖（含 skip_zero 前的 0 值）
                        continue
                    if cfg.get("skip_zero", True) and v == 0:
                        continue  # 0 = 无数据（未发生月份/占位），不写事实
                    items.append((year, month, node_id, res.field_id[field_name], dtype_id, v))
                else:
                    # 聚合块值全部参与校验（含 0：块值 0 而叶子非零正是校验要捕获的源表缺陷）
                    leaf_by_node.setdefault(node_id, []).append((res.field_id[field_name], month, v))
    wb.close()

    if protected_skipped:
        warnings.append(f"受保护单元格跳过 {protected_skipped} 个（{cfg.get('dtype')}：明细源为权威，宽表不覆盖）")

    # 聚合校验：聚合块值 vs 叶子和（仅 raw 字段可加；formula 字段跳过，报表期再校验）
    for node_id, vals in leaf_by_node.items():
        leaves = set(calc.subtree_leaves(conn, node_id))
        for field_id, month, v in vals:
            field_name = _name_of(conn, field_id)
            if res.field_kind.get(field_name) != "raw":
                continue
            leaf_sum = sum(x[5] for x in items if x[1] == month and x[2] in leaves
                           and x[3] == field_id)
            if abs(leaf_sum - v) > 0.01:
                warnings.append(
                    f"聚合校验 KA-{'-'.join(calc.node_path(conn, node_id))} {field_name} {month}月: "
                    f"块值={v:,.2f} vs 叶子和={leaf_sum:,.2f}（差 {v - leaf_sum:+,.2f}）"
                )
    return items, warnings


def _name_of(conn, field_id):
    row = conn.execute("SELECT name FROM fields WHERE id=?", (field_id,)).fetchone()
    return row["name"] if row else "?"


def find_conflicts(conn, items, tol=0.01):
    """items 中与库中现有值差异超容差的记录（导入前冲突检查）"""
    conflicts = []
    for it in items:
        old = conn.execute(
            "SELECT value FROM facts WHERE year=? AND month=? AND node_id=? AND field_id=? AND dtype_id=?",
            it[:5],
        ).fetchone()
        if old and abs(old["value"] - it[5]) > tol:
            conflicts.append({"item": list(it[:6]), "old": old["value"]})
    return conflicts


def _to_int(v):
    """容错数字：int/整 float/数字字符串('07'/'2026') -> int；非整 float（7.9 公式残留）拒绝"""
    if isinstance(v, float) and not v.is_integer():
        return None  # 非整月/年（公式产物/手误），拒绝导入而非静默截断
    if isinstance(v, (int, float)):
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return None


def parse_detail(conn, cfg, file_path, sheet_name):
    """明细长表解析：行=年月×品类×渠道，列=指标（col_map 支持 A+B 多列求和）"""
    res = Resolver(conn)
    warnings = []

    def coli(letter):
        return column_index_from_string(str(letter).upper())

    year_c, month_c = coli(cfg["year_col"]), coli(cfg["month_col"])
    dtype_c = coli(cfg["dtype_col"]) if cfg.get("dtype_col") else None
    prod_c, chan_c = coli(cfg["product_col"]), coli(cfg["channel_col"])
    filters = {coli(k): v for k, v in cfg.get("filters", {}).items()}
    metric_cols = [( [coli(c) for c in expr.split("+")], name)
                   for expr, name in cfg["col_map"].items()]
    dtype_default = cfg.get("dtype_default")
    seen_unknown = set()

    wb = openpyxl.load_workbook(file_path, data_only=True)  # 明细源表很小，普通模式更快
    ws = wb[sheet_name]
    items = []
    for r in range(2, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value != v for c, v in filters.items()):
            continue
        yv, mv = ws.cell(row=r, column=year_c).value, ws.cell(row=r, column=month_c).value
        year, month = _to_int(yv), _to_int(mv)
        if year is None or month is None:
            warnings.append(f"行{r}: 年份/月份无效 {yv!r}/{mv!r}（非整数），跳过")
            continue
        dv = ws.cell(row=r, column=dtype_c).value if dtype_c else None
        dv = dv if dv is not None else dtype_default
        if not dv or dv not in res.dtype_id:
            warnings.append(f"行{r}: 未知数据类型 {dv!r}，跳过")
            continue
        dti = res.dtype_id[dv]
        pl = ws.cell(row=r, column=prod_c).value
        if pl not in cfg["product_map"]:
            if pl not in seen_unknown:
                seen_unknown.add(pl)
                warnings.append(f"行{r}: 未知品类 {pl!r}，跳过（首次）")
            continue
        prod = cfg["product_map"][pl]
        cl = ws.cell(row=r, column=chan_c).value
        if cl not in cfg["channel_map"]:
            if cl not in seen_unknown:
                seen_unknown.add(cl)
                warnings.append(f"行{r}: 未知渠道 {cl!r}，跳过（首次）")
            continue
        path = tuple(cfg["channel_map"][cl])
        node_id = calc.resolve_path(conn, prod, path)
        if node_id is None:
            warnings.append(f"行{r}: 渠道路径 {prod}/{path} 未找到节点")
            continue
        for cols, field_name in metric_cols:
            if field_name not in res.field_id:
                if field_name not in seen_unknown:
                    seen_unknown.add(field_name)
                    warnings.append(f"列映射字段 {field_name!r} 不在字段表，跳过（首次）")
                continue
            vals = [ws.cell(row=r, column=c).value for c in cols]
            vals = [v for v in vals if isinstance(v, (int, float)) and math.isfinite(v)]
            if not vals:
                continue
            total = sum(vals)
            if cfg.get("skip_zero", True) and total == 0:
                continue  # 0 = 无数据，不写事实
            items.append((year, month, node_id, res.field_id[field_name], dti, total))
    wb.close()
    return items, warnings


def parse_source(conn, template_row, file_path):
    """按模板解析来源文件：返回 {items, warnings, summary}（不改库）"""
    cfg = json.loads(template_row["config_json"])
    res = Resolver(conn)  # 节点->产品名缓存（解析时已加载，免逐条回查）
    if cfg["style"] == "wide":
        items, warnings = parse_wide(conn, cfg, file_path, template_row["sheet_name"])
    elif cfg["style"] == "detail":
        items, warnings = parse_detail(conn, cfg, file_path, template_row["sheet_name"])
    else:
        raise ValueError(f"未知导入形态: {cfg['style']}")
    summary = {
        "条数": len(items),
        "产品分布": {},
        "月份分布": {},
    }
    for it in items:
        pname = res.node.get(it[2], {}).get("prod", "?")
        summary["产品分布"][pname] = summary["产品分布"].get(pname, 0) + 1
        ym = f"{it[0]}-{it[1]:02d}"
        summary["月份分布"][ym] = summary["月份分布"].get(ym, 0) + 1
    return {"items": items, "warnings": warnings, "summary": summary}


def commit_import(conn, parsed, source_note):
    """入库（upsert）：同键已存在且值不同则覆盖并计数；返回 {写入, 覆盖}"""
    cur = conn.cursor()
    written = overwritten = 0
    for year, month, node_id, field_id, dtype_id, value in parsed["items"]:
        old = cur.execute(
            "SELECT id FROM facts WHERE year=? AND month=? AND node_id=? AND field_id=? AND dtype_id=?",
            (year, month, node_id, field_id, dtype_id),
        ).fetchone()
        cur.execute(
            "INSERT INTO facts (year, month, node_id, field_id, dtype_id, value, source) VALUES (?,?,?,?,?,?,?) "
            "ON CONFLICT(year, month, node_id, field_id, dtype_id) DO UPDATE SET value=excluded.value, source=excluded.source",
            (year, month, node_id, field_id, dtype_id, value, source_note),
        )
        if old:
            overwritten += 1
        else:
            written += 1
    conn.commit()
    return {"写入": written, "覆盖": overwritten}
