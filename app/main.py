# -*- coding: utf-8 -*-
"""FastAPI 入口：静态前端 + JSON API（本地单机，无需鉴权）"""
import pathlib

import openpyxl
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from . import calc, importers, reports
from .db import BASE_DIR, DATA_DIR, get_conn, init_db

app = FastAPI(title="经营数据管理系统", docs_url=None, redoc_url=None)
STATIC = BASE_DIR / "static"
OUT_DIR = DATA_DIR / "输出"


def _rows(conn, sql, args=()):
    return [dict(r) for r in conn.execute(sql, args)]


@app.on_event("startup")
def _startup():
    init_db()
    OUT_DIR.mkdir(exist_ok=True)


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


app.mount("/static", StaticFiles(directory=STATIC), name="static")


# ---------- 元数据 ----------
@app.get("/api/meta")
def api_meta():
    conn = get_conn()
    return {
        "products": _rows(conn, "SELECT id, name FROM product_lines"),
        "data_types": _rows(conn, "SELECT id, name, year_shift, base_dtype_id FROM data_types"),
        "fields": _rows(conn, "SELECT id, name, kind, formula, unit, scale FROM fields"),
        "measures": _rows(conn, "SELECT id, name, kind, dtype_id, formula FROM measures"),
        "channel_trees": _rows(
            conn,
            "SELECT cn.id, cn.parent_id, cn.name, cn.is_leaf, p.name AS product "
            "FROM channel_nodes cn JOIN product_lines p ON p.id=cn.product_id "
            "ORDER BY cn.product_id, cn.id",
        ),
        "source_templates": _rows(
            conn, "SELECT id, name, path_pattern, sheet_name, note FROM source_templates WHERE enabled=1"
        ),
        "report_templates": _rows(
            conn, "SELECT id, name, note FROM report_templates WHERE enabled=1"
        ),
    }


# ---------- 手工录入 ----------
@app.get("/api/facts")
def api_facts(year: int, month: int, product_id: int):
    """某产品某月的叶子事实：{node_id: {field_id: {dtype_id: value}}}"""
    conn = get_conn()
    leaves = _rows(
        conn,
        "SELECT id FROM channel_nodes WHERE product_id=? AND is_leaf=1",
        (product_id,),
    )
    if not leaves:
        return {}
    ph = ",".join("?" * len(leaves))
    facts = {}
    for r in conn.execute(
        f"SELECT node_id, field_id, dtype_id, value FROM facts "
        f"WHERE year=? AND month=? AND node_id IN ({ph})",
        (year, month, *[l["id"] for l in leaves]),
    ):
        facts.setdefault(r["node_id"], {}).setdefault(r["field_id"], {})[r["dtype_id"]] = r["value"]
    # 全部叶子都要返回（无数据 = 空值，前端显示空输入框）
    return {l["id"]: facts.get(l["id"], {}) for l in leaves}


@app.post("/api/fact")
def api_fact(payload: dict):
    """写入/删除一个叶子值（value=None 删除）；公式字段不允许手工写入"""
    conn = get_conn()
    year, month = payload["year"], payload["month"]
    node_id, field_id, dtype_id = payload["node_id"], payload["field_id"], payload["dtype_id"]
    kind = conn.execute("SELECT kind FROM fields WHERE id=?", (field_id,)).fetchone()
    if kind and kind["kind"] == "formula":
        raise HTTPException(400, "公式字段由引擎计算，不允许手工录入")
    value = payload.get("value")
    try:
        value = None if value is None or value == "" else float(value)
    except (TypeError, ValueError):
        raise HTTPException(400, "值必须为数字")
    if value is None:
        conn.execute(
            "DELETE FROM facts WHERE year=? AND month=? AND node_id=? AND field_id=? AND dtype_id=?",
            (year, month, node_id, field_id, dtype_id),
        )
    else:
        conn.execute(
            "INSERT INTO facts (year, month, node_id, field_id, dtype_id, value, source) "
            "VALUES (?,?,?,?,?,?, '手工录入') "
            "ON CONFLICT(year, month, node_id, field_id, dtype_id) "
            "DO UPDATE SET value=excluded.value, source=excluded.source",
            (year, month, node_id, field_id, dtype_id, value),
        )
    conn.commit()
    return {"ok": True}


# ---------- 导入 ----------
def _parse_for_api(conn, template_id, file_path):
    tpl = conn.execute("SELECT * FROM source_templates WHERE id=?", (template_id,)).fetchone()
    if not tpl:
        raise HTTPException(404, "导入模板不存在")
    try:
        parsed = importers.parse_source(conn, tpl, file_path)
    except Exception as e:
        raise HTTPException(400, f"解析失败: {e}")
    return tpl, parsed


@app.post("/api/import/preview")
def api_import_preview(payload: dict):
    """解析预览（不改库）：汇总 + 警告 + 冲突 + 前 50 条示例"""
    conn = get_conn()
    tid, path = payload.get("template_id"), payload.get("file_path")
    if not tid or not path:
        raise HTTPException(400, "缺少 template_id 或 file_path")
    tpl, parsed = _parse_for_api(conn, tid, path)
    conflicts = importers.find_conflicts(conn, parsed["items"])
    return {
        "template_name": tpl["name"],
        "summary": parsed["summary"],
        "warnings": parsed["warnings"][:100],
        "warnings_total": len(parsed["warnings"]),
        "conflicts": conflicts[:100],
        "conflicts_total": len(conflicts),
        "items_sample": [list(it) for it in parsed["items"][:50]],
    }


@app.post("/api/import/commit")
def api_import_commit(payload: dict):
    conn = get_conn()
    tid, path = payload.get("template_id"), payload.get("file_path")
    if not tid or not path:
        raise HTTPException(400, "缺少 template_id 或 file_path")
    tpl, parsed = _parse_for_api(conn, tid, path)
    res = importers.commit_import(conn, parsed, payload.get("source_note") or tpl["name"])
    return {"ok": True, **res}


# ---------- 数据明细 ----------
_SORT_COLS = {
    "ym": ("year", "month"), "product": ("product",), "channel": ("channel",),
    "field": ("field",), "dtype": ("dtype",), "value": ("value",),
}


def _details_rows(conn, filters, include_formula=False):
    """明细长表行：存储事实 + 可选公式字段计算值（含渠道路径）
    多值筛选：month_from/month_to 区间、product_ids/channel_ids 多节点并集、dtype_ids 多类型、
    field_ids 多字段（raw 查存储行 / formula 算计算值，可混合）"""
    where, args = [], []
    for key, col in (("year", "f.year"),):
        if filters.get(key):
            where.append(f"{col}=?")
            args.append(filters[key])
    if filters.get("month_from"):
        where.append("f.month>=?")
        args.append(filters["month_from"])
    if filters.get("month_to"):
        where.append("f.month<=?")
        args.append(filters["month_to"])
    if filters.get("product_ids"):
        ph = ",".join("?" * len(filters["product_ids"]))
        where.append(f"cn.product_id IN ({ph})")
        args.extend(filters["product_ids"])
    elif filters.get("product_id"):
        where.append("cn.product_id=?")
        args.append(filters["product_id"])
    if filters.get("dtype_ids"):
        ph = ",".join("?" * len(filters["dtype_ids"]))
        where.append(f"f.dtype_id IN ({ph})")
        args.extend(filters["dtype_ids"])
    elif filters.get("dtype_id"):
        where.append("f.dtype_id=?")
        args.append(filters["dtype_id"])
    channel_ids = filters.get("channel_ids") or ([filters["channel_id"]] if filters.get("channel_id") else None)
    if channel_ids:
        # 节点集并集筛选：各自展开为叶子（聚合节点 = 其子树全部叶子；叶子节点 = 精确匹配）
        leaves = set()
        for nid in channel_ids:
            leaves.update(calc.subtree_leaves(conn, nid))
        if leaves:
            ph = ",".join("?" * len(leaves))
            where.append(f"f.node_id IN ({ph})")
            args.extend(leaves)
        else:
            where.append("1=0")
    base_w = " WHERE " + " AND ".join(where) if where else ""
    # 字段过滤（多值）：raw 查存储行；formula 不落库 → 仅算计算值；raw+formula 混合可同时存在
    stored_w, stored_args = base_w, args
    fid_filters = filters.get("field_ids") or ([filters["field_id"]] if filters.get("field_id") else None)
    formula_ids = None
    if fid_filters:
        ph = ",".join("?" * len(fid_filters))
        raw_ids = [r["id"] for r in conn.execute(
            f"SELECT id FROM fields WHERE id IN ({ph}) AND kind='raw'", fid_filters)]
        formula_ids = [r["id"] for r in conn.execute(
            f"SELECT id FROM fields WHERE id IN ({ph}) AND kind='formula'", fid_filters)]
        if raw_ids:
            ph2 = ",".join("?" * len(raw_ids))
            stored_w += f" AND f.field_id IN ({ph2})"
            stored_args = args + raw_ids
        else:
            stored_w += " AND 1=0"
    rows = [dict(r) for r in conn.execute(
        f"SELECT f.year, f.month, cn.id AS node_id, p.name AS product, "
        f"f.field_id, fl.name AS field, f.dtype_id, d.name AS dtype, f.value, f.source "
        f"FROM facts f JOIN channel_nodes cn ON cn.id=f.node_id "
        f"JOIN product_lines p ON p.id=cn.product_id "
        f"JOIN fields fl ON fl.id=f.field_id JOIN data_types d ON d.id=f.dtype_id{stored_w}",
        stored_args)]
    # 渠道路径（一次构建节点映射）
    nodes = {r["id"]: r for r in conn.execute("SELECT id, name, parent_id FROM channel_nodes")}

    def path_of(nid):
        p, cur = [], nid
        while cur and cur in nodes:
            p.append(nodes[cur]["name"])
            cur = nodes[cur]["parent_id"]
        return "KA-" + "-".join(reversed(p))

    for r in rows:
        r["channel"] = path_of(r["node_id"])
    if include_formula:
        # 公式字段计算值：对每个 年/月/节点/数据类型 组合，按同 dtype 上下文计算
        # 批量求值器缓存叶子/事实行/元数据，免逐组合重建 Resolver（O(组合×字段) SQL → 常数级）
        combos = [dict(r) for r in conn.execute(
            f"SELECT DISTINCT f.year, f.month, f.node_id, f.dtype_id, p.name AS product, d.name AS dtype "
            f"FROM facts f JOIN channel_nodes cn ON cn.id=f.node_id "
            f"JOIN product_lines p ON p.id=cn.product_id JOIN data_types d ON d.id=f.dtype_id{base_w}",
            args)]
        fids = conn.execute("SELECT id, name FROM fields WHERE kind='formula'").fetchall()
        if formula_ids is not None:
            fids = [r for r in fids if r["id"] in formula_ids]
        evaluator = calc.FormulaBulkEvaluator(conn)
        for c in combos:
            for fr in fids:
                v = evaluator.value(c["year"], c["month"], c["node_id"], fr["id"], c["dtype_id"])
                if v is not None:
                    rows.append({
                        "year": c["year"], "month": c["month"], "node_id": c["node_id"],
                        "product": c["product"], "field_id": fr["id"], "field": fr["name"],
                        "dtype_id": c["dtype_id"], "dtype": c["dtype"], "value": v,
                        "source": "计算值", "channel": path_of(c["node_id"]),
                    })
    return rows


def _details_filters(year=None, month=None, month_from=None, month_to=None, product_id=None,
                     product_ids=None, dtype_id=None, dtype_ids=None, field_id=None,
                     field_ids=None, channel_id=None, channel_ids=None):
    """筛选参数字典（多值/区间 + 兼容单值旧参数；*_ids 为逗号分隔字符串）"""
    filters = {}
    if year:
        filters["year"] = year
    if month:
        filters["month_from"] = filters["month_to"] = month  # 旧单值参数 = 区间单月
    else:
        if month_from:
            filters["month_from"] = month_from
        if month_to:
            filters["month_to"] = month_to
    if product_ids:
        filters["product_ids"] = [int(x) for x in str(product_ids).split(",") if x]
    elif product_id:
        filters["product_id"] = product_id
    if dtype_ids:
        filters["dtype_ids"] = [int(x) for x in str(dtype_ids).split(",") if x]
    elif dtype_id:
        filters["dtype_id"] = dtype_id
    if channel_ids:
        filters["channel_ids"] = [int(x) for x in str(channel_ids).split(",") if x]
    elif channel_id:
        filters["channel_id"] = channel_id
    if field_ids:
        filters["field_ids"] = [int(x) for x in str(field_ids).split(",") if x]
    elif field_id:
        filters["field_id"] = field_id
    return filters


@app.get("/api/details")
def api_details(year: int = None, month: int = None, month_from: int = None, month_to: int = None,
                product_id: int = None, product_ids: str = None, dtype_id: int = None,
                dtype_ids: str = None, field_id: int = None, field_ids: str = None,
                channel_id: int = None, channel_ids: str = None,
                include_formula: int = 0,
                page: int = 1, page_size: int = 500, sort: str = "ym", sort_dir: str = "asc"):
    conn = get_conn()
    filters = _details_filters(year, month, month_from, month_to, product_id, product_ids,
                               dtype_id, dtype_ids, field_id, field_ids, channel_id, channel_ids)
    rows = _details_rows(conn, filters, bool(include_formula))
    sort_keys = _SORT_COLS.get(sort, _SORT_COLS["ym"])
    rows.sort(key=lambda r: tuple(r[k] for k in sort_keys), reverse=(sort_dir == "desc"))
    total = len(rows)
    page = max(1, page)  # 页码钳制
    page_rows = rows[(page - 1) * page_size: page * page_size]
    return {"total": total, "page": page, "page_size": page_size, "rows": page_rows}


@app.get("/api/details/export")
def api_details_export(year: int = None, month: int = None, month_from: int = None,
                       month_to: int = None, product_id: int = None, product_ids: str = None,
                       dtype_id: int = None, dtype_ids: str = None, field_id: int = None,
                       field_ids: str = None, channel_id: int = None, channel_ids: str = None,
                       include_formula: int = 0):
    conn = get_conn()
    filters = _details_filters(year, month, month_from, month_to, product_id, product_ids,
                               dtype_id, dtype_ids, field_id, field_ids, channel_id, channel_ids)
    rows = _details_rows(conn, filters, bool(include_formula))
    rows.sort(key=lambda r: (r["year"], r["month"], r["product"], r["channel"], r["field"], r["dtype"]))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细数据"
    headers = ["年", "月", "产品", "渠道", "字段", "数据类型", "值", "来源"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="1E3A5F")
        cell.alignment = openpyxl.styles.Alignment(horizontal="center")
    widths = [8, 6, 10, 24, 14, 10, 14, 12]
    for i, wdt in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wdt
    for r in rows:
        ws.append([r["year"], r["month"], r["product"], r["channel"], r["field"],
                   r["dtype"], r["value"], r["source"]])
    OUT_DIR.mkdir(exist_ok=True)
    from datetime import datetime
    out = OUT_DIR / f"明细数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(str(out))
    return FileResponse(str(out), filename=out.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------- 报表 ----------
@app.get("/api/report/render")
def api_report_render(template_id: int, year: int, month: int):
    conn = get_conn()
    tpl = conn.execute("SELECT * FROM report_templates WHERE id=?", (template_id,)).fetchone()
    if not tpl:
        raise HTTPException(404, "报表模板不存在")
    t = reports.render(conn, tpl, year, month)
    return {"title": t["title"], "columns": reports.api_columns(t["columns"]), "rows": t["rows"]}


@app.get("/api/report/explain")
def api_report_explain(template_id: int, year: int, month: int, row: int, col: int):
    """单元格计算过程求值树（计算指标弹窗用）"""
    conn = get_conn()
    tpl = conn.execute("SELECT * FROM report_templates WHERE id=?", (template_id,)).fetchone()
    if not tpl:
        raise HTTPException(404, "报表模板不存在")
    t = reports.layout(conn, tpl, year, month)  # 只算结构不算值，免整表重渲染
    if not (0 <= row < len(t["rows"]) and 0 <= col < len(t["columns"])):
        raise HTTPException(400, f"行列越界: row={row} col={col}")
    r, c = t["rows"][row], t["columns"][col]
    # 旧式模板（月度/累计报表）行无字段，字段在列上；新式（预测/费用分析）行有字段
    fname = r.get("field") or c.get("field")
    field_id = conn.execute("SELECT id FROM fields WHERE name=?", (fname,)).fetchone()
    if field_id is None:
        raise HTTPException(400, "该单元格无对应字段")
    nids = r.get("nids") or c.get("channel_nids")
    if not nids:
        raise HTTPException(400, "该单元格无渠道信息")
    # 行级口径覆盖（费用分析收入行 BP比 → BP 原值）：弹窗与单元格显示保持一致
    ov = (r.get("measure_overrides") or {}).get(col)
    mrow = c["measure_row"]
    if ov:
        mrow = conn.execute("SELECT * FROM measures WHERE name=?", (ov,)).fetchone()
        if mrow is None:
            raise HTTPException(400, f"未知口径: {ov}")
    tree = calc.explain_period_measure(conn, year, c["segments"], nids, field_id["id"],
                                       mrow, mixed=(mrow["name"] == "实际"))
    return tree


@app.get("/api/report/export")
def api_report_export(template_id: int, year: int, month: int):
    conn = get_conn()
    tpl = conn.execute("SELECT * FROM report_templates WHERE id=?", (template_id,)).fetchone()
    if not tpl:
        raise HTTPException(404, "报表模板不存在")
    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / f"{year}年{month}月-{tpl['name']}.xlsx"
    reports.export_report(conn, tpl, year, month, str(out))
    return FileResponse(
        str(out),
        filename=out.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
