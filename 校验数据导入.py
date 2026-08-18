# -*- coding: utf-8 -*-
"""数据导入独立校验：库中 facts vs 源表逐格核对
1. 月份映射：所有块 表头 vs 位置
2. 抽样核对：库值 vs 宽表（智屏 1-6月 除外——明细库修订值覆盖宽表旧值，属已知源差异）
3. 聚合不变式：聚合块值 vs 叶子和（分类：冰洗5月=源表不一致；智屏1-6月=明细修订差异）
4. 智屏明细源 vs 库（全量，明细为权威）
5. 报表级：月度报表 15行×25列 vs 26年7月报表（ASP 块达成率/同比为差值口径，与公式字段重算一致）
6. 同期完整性
"""
import random
import openpyxl

from app import calc, importers
from app.db import get_conn

MAIN = r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx"
DETAIL = r"E:\10 经营分析报告\7月报告\各品类指标表\1 智屏指标表\26年7月KA渠道智屏TCL数据库.xlsx"

conn = get_conn()
ok = fail = 0
notes = []


def chk(label, actual, expect, tol=0.02, allow_zero_none=True):
    """核对；allow_zero_none：源表 0 vs 库 None 视为一致（skip_zero 设计）"""
    global ok, fail
    if isinstance(expect, str):
        notes.append(f"源表异常值 {label}: 源表={expect!r}，跳过")
        return
    if expect is None:
        good = actual is None
    elif allow_zero_none and expect == 0 and actual is None:
        good = True
    else:
        good = actual is not None and abs(actual - expect) <= tol
    if good:
        ok += 1
    else:
        fail += 1
        print(f"  ✗ {label}: 库值={actual if actual is None else round(actual, 4)} vs 源表={expect if expect is None else round(expect, 4)}")


wb = openpyxl.load_workbook(MAIN, data_only=True)
field_ids = {r["name"]: r["id"] for r in conn.execute("SELECT id, name FROM fields")}
from app.importers import _resolve_label

# ---------- 1. 月份映射 ----------
print("=== 1. 月份映射（表头 vs 位置）===")
bad_map = 0
for sheet in ("26年实际", "26年预算", "25实际"):
    ws = wb[sheet]
    starts = [c for c in range(2, ws.max_column + 1) if ws.cell(row=2, column=c).value == "合计"]
    for s in starts:
        for off in range(1, 13):
            hdr = ws.cell(row=2, column=s + off).value
            if hdr != f"{off}月":
                bad_map += 1
                if bad_map <= 10:
                    print(f"  ⚠ {sheet} 块{s} 列{s+off}: 表头={hdr!r} != {off}月位")
print(f"  主文件三表 表头错位共 {bad_map} 处（应为 0）")

# ---------- 2. 抽样核对 ----------
print("\n=== 2. 抽样核对（库值 vs 宽表，80 条）===")
random.seed(42)
ws26 = wb["26年实际"]
starts26 = [c for c in range(2, ws26.max_column + 1) if ws26.cell(row=2, column=c).value == "合计"]
leaf_blocks = []
for s in starts26:
    nid = _resolve_label(conn, ws26.cell(row=1, column=s).value)
    if nid is not None and conn.execute("SELECT is_leaf FROM channel_nodes WHERE id=?", (nid,)).fetchone()["is_leaf"]:
        leaf_blocks.append((s, nid))
row_names = {r: ws26.cell(row=r, column=1).value for r in range(3, 40)}
samples = []
for s, nid in leaf_blocks:
    prod = conn.execute("SELECT p.name FROM channel_nodes cn JOIN product_lines p ON p.id=cn.product_id "
                        "WHERE cn.id=?", (nid,)).fetchone()["name"]
    for off in range(1, 8):
        if prod == "智屏" and off <= 6:
            continue  # 智屏 1-6月 已知明细修订差异，跳过
        for r, name in row_names.items():
            if name in field_ids and isinstance(ws26.cell(row=r, column=s + off).value, (int, float)):
                samples.append((nid, off, field_ids[name], ws26.cell(row=r, column=s + off).value))
sampled = random.sample(samples, 80)
for nid, month, fid, expect in sampled:
    chk(f"node{nid} {month}月 字段{fid}", calc.field_value(conn, 2026, month, [nid], fid, 1), expect)
print(f"  抽样 {len(sampled)} 条")

# ---------- 3. 聚合不变式 ----------
print("\n=== 3. 聚合不变式（聚合块值 vs 叶子和）===")
agg_total = agg_bad = 0
for s in starts26:
    nid = _resolve_label(conn, ws26.cell(row=1, column=s).value)
    if nid is None:
        continue
    if conn.execute("SELECT is_leaf FROM channel_nodes WHERE id=?", (nid,)).fetchone()["is_leaf"]:
        continue
    label = ws26.cell(row=1, column=s).value
    prod = label.split("-")[1] if "-" in label else ""
    for off in range(1, 8):
        for r, name in row_names.items():
            if name not in field_ids:
                continue
            cell = ws26.cell(row=r, column=s + off).value
            if not isinstance(cell, (int, float)):
                continue
            agg_total += 1
            v = calc.field_value(conn, 2026, off, [nid], field_ids[name], 1)
            if abs((v or 0) - cell) > 0.02:
                agg_bad += 1
                if agg_bad <= 10:
                    print(f"  ⚠ {label} {name} {off}月: 叶子和={(v or 0):.2f} vs 块值={cell:.2f}")
print(f"  聚合核对 {agg_total} 处，不一致 {agg_bad} 处")
print("  （分类：冰洗 5月 4 处 = 源表数据不一致；智屏 1-6月 = 明细库修订值 vs 宽表旧值）")

# ---------- 4. 智屏明细源全量核对 ----------
print("\n=== 4. 智屏明细源 vs 库（明细为权威）===")
tpl_d = conn.execute("SELECT * FROM source_templates WHERE id=4").fetchone()
pd = importers.parse_source(conn, tpl_d, DETAIL)
d_bad = 0
for it in pd["items"]:
    v = calc.field_value(conn, it[0], it[1], [it[2]], it[3], it[4])
    if v is None or abs(v - it[5]) > 0.02:
        d_bad += 1
        if d_bad <= 5:
            print(f"  ✗ {it[0]}-{it[1]} node{it[2]} 字段{it[3]}: 库={v} vs 明细源={it[5]}")
print(f"  明细 {len(pd['items'])} 条，不一致 {d_bad} 条")

# ---------- 5. 报表级全表 ----------
print("\n=== 5. 月度报表 15行×25列 vs 26年7月报表 ===")
from app import reports
tpl_r = conn.execute("SELECT * FROM report_templates WHERE id=1").fetchone()
t = reports.render(conn, tpl_r, 2026, 7)
ws7 = wb["26年7月报表"]
row_map = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]
checked = cell_bad = 0
for ri, (row, src_r) in enumerate(zip(t["rows"], row_map)):
    for ci, v in enumerate(row["cells"]):
        src = ws7.cell(row=src_r, column=2 + ci).value
        if not isinstance(src, (int, float)) and src is not None:
            notes.append(f"原表 {src_r}行 列{2+ci} 非数值: {src!r}")
            continue
        if src is None and v is None:
            continue
        if (src == 0 and v is None) or (src is None and v is None):
            continue
        src = src if src is not None else 0
        v = v if v is not None else 0
        tol = 0.005 if ci % 5 in (2, 4) else 0.02
        checked += 1
        if abs(v - src) > tol:
            cell_bad += 1
            if cell_bad <= 10:
                print(f"  ✗ 行{src_r} 列{ci} ({row['entity']}): 渲染={round(v, 4)} vs 原表={round(src, 4)}")
print(f"  核对 {checked} 格，不一致 {cell_bad} 格")

# ---------- 6. 同期完整性 ----------
print("\n=== 6. 同期完整性 ===")
m_spl = conn.execute("SELECT * FROM measures WHERE name='同期'").fetchone()
for prod, row7 in [("智屏", 6), ("空调", 9), ("冰洗", 12), ("CIoT", 15)]:
    nid = calc.resolve_path(conn, prod, ())
    v = calc.measure_value(conn, 2026, 7, [nid], 1, m_spl)
    chk(f"{prod} 7月同期", v, ws7.cell(row=row7, column=5).value, 0.05)

print("\n=== 报告 ===")
for n in notes:
    print("  ·", n)
print(f"\n结果: {ok} 通过, {fail} 失败")
# 已知基线（用户已确认忽略，见第3/4节分类注释）：智屏 1-6月 明细修订差异 d_bad=24、
# 聚合差异 agg_bad=16（含冰洗5月源表不一致）；超出基线（源表更新/回归）才失败；表头错位必须为 0
if fail or cell_bad or bad_map or d_bad > 24 or agg_bad > 16:
    raise SystemExit(1)
print("数据导入校验：通过 ✓（基线内已知差异见报告）")
