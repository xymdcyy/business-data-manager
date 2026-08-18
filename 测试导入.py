# -*- coding: utf-8 -*-
"""导入引擎测试：宽表 26年实际/25实际 解析→入库→与 26年7月报表 缓存值交叉验证（临时库副本）"""
import openpyxl

from app import calc, importers
from 测试helpers import chk, result, use_test_db

MAIN = r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx"
conn = use_test_db()


# ---------- 1. 解析预览（不改库） ----------
tpl = conn.execute("SELECT * FROM source_templates WHERE name='26年实际'").fetchone()
parsed = importers.parse_source(conn, tpl, MAIN)
print("=== 26年实际 解析预览 ===")
print("  条数:", parsed["summary"]["条数"], "| 产品分布:", parsed["summary"]["产品分布"])
print("  月份分布:", parsed["summary"]["月份分布"])
print("  警告:", len(parsed["warnings"]))
for w in parsed["warnings"][:10]:
    print("   -", w)

# ---------- 2. 入库 ----------
res = importers.commit_import(conn, parsed, "底稿(1).xlsx 26年实际")
print("\n=== 入库:", res, "===")

# ---------- 3. 与 26年7月报表 交叉验证 ----------
print("\n=== 3. 交叉验证 vs 26年7月报表（7月 实际）===")
wbv = openpyxl.load_workbook(MAIN, data_only=True)
ws = wbv["26年7月报表"]
# 报表行: 3=KA合计 6=智屏 9=空调 12=冰洗 15=CIoT; B=收入 C=BP D=达成率 E=同期 F=同比
# 智屏根/空调根/冰洗根/CIoT根 节点 id = 1/7/13/19
expected = {
    "智屏": (6, "B"), "空调": (9, "B"), "冰洗": (12, "B"), "CIoT": (15, "B"),
}
for prod, (row, col) in expected.items():
    nid = calc.resolve_path(conn, prod, ())
    v = calc.field_value(conn, 2026, 7, [nid], 1, 1)  # 收入 实际
    expect = ws[f"{col}{row}"].value
    if prod == "空调" and (expect == 0 or expect is None):
        chk(f"空调 收入(万元) = 源表7月空（真实缺口，报表显示0）", v, None)
    else:
        chk(f"{prod} 收入(万元) = 报表 B{row}", v, expect, tol=0.02)
# 销量 G6
for prod, (row, col) in {"智屏": (6, "G"), "空调": (9, "G")}.items():
    nid = calc.resolve_path(conn, prod, ())
    v = calc.field_value(conn, 2026, 7, [nid], 2, 1)
    expect = ws[f"{col}{row}"].value
    if prod == "空调" and (expect == 0 or expect is None):
        chk("空调 销量 = 源表7月空", v, None)
    else:
        chk(f"{prod} 销量(万台) = 报表 {col}{row}", v, expect, tol=0.02)
# KA合计 = 四品类之和
ka = calc.selector_sum(conn, 2026, 7, [1, 7, 13, 19], 1, 1)
chk("KA合计 收入 = 报表 B3", ka, ws["B3"].value, tol=0.05)
# 端到端净利 V6（字段33）
nid = calc.resolve_path(conn, "智屏", ())
chk("智屏 端到端净利 = 报表 V6", calc.field_value(conn, 2026, 7, [nid], 33, 1), ws["V6"].value, tol=0.05)
# ASP 公式（字段3） vs 报表 L6
asp = calc.field_value(conn, 2026, 7, [nid], 3, 1)
chk("智屏 ASP = 报表 L6", asp, ws["L6"].value, tol=0.5)

# ---------- 4. 25实际（同期基础） ----------
print("\n=== 4. 25实际 导入 ===")
tpl25 = conn.execute("SELECT * FROM source_templates WHERE name='25实际'").fetchone()
p25 = importers.parse_source(conn, tpl25, MAIN)
print("  条数:", p25["summary"]["条数"], "| 警告:", len(p25["warnings"]))
for w in p25["warnings"][:8]:
    print("   -", w)
importers.commit_import(conn, p25, "底稿(1).xlsx 25实际")

# 同期/同比验证：智屏 7月 同期 vs 报表 E6；同比 vs F6
nid = calc.resolve_path(conn, "智屏", ())
m_spl = conn.execute("SELECT * FROM measures WHERE name='同期'").fetchone()
m_yoy = conn.execute("SELECT * FROM measures WHERE name='同比'").fetchone()
chk("智屏 同期 = 报表 E6", calc.measure_value(conn, 2026, 7, [nid], 1, m_spl), ws["E6"].value, tol=0.05)
chk("智屏 同比 = 报表 F6", calc.measure_value(conn, 2026, 7, [nid], 1, m_yoy), ws["F6"].value, tol=0.001)

# ---------- 5. 智屏明细源（长表）解析验证 ----------
print("\n=== 5. 智屏实际明细（长表）解析 ===")
tpl_d = conn.execute("SELECT * FROM source_templates WHERE name='智屏实际明细'").fetchone()
DETAIL = r"E:\10 经营分析报告\7月报告\各品类指标表\1 智屏指标表\26年7月KA渠道智屏TCL数据库.xlsx"
pd = importers.parse_source(conn, tpl_d, DETAIL)
print("  条数:", pd["summary"]["条数"], "| 产品分布:", pd["summary"]["产品分布"], "| 警告:", len(pd["warnings"]))
for w in pd["warnings"][:6]:
    print("   -", w)
# 明细 2026-07 智屏 收入 = 宽表导入值（同源应一致）
d_items = {it for it in pd["items"]}
wide_july = [it for it in d_items if it[0] == 2026 and it[1] == 7]
# 节点 3=连锁-连锁, 6=区域连锁
chk("明细源 智屏-连锁-连锁 7月收入 = 6446.93",
    next((v for y, m, n, f, d, v in wide_july if n == 3 and f == 1), None), 6446.93, tol=0.02)
chk("明细源 智屏-区域连锁 7月收入 = 3198.46",
    next((v for y, m, n, f, d, v in wide_july if n == 6 and f == 1), None), 3198.46, tol=0.02)
chk("明细源 智屏 7月收入合计(叶子) = 9645.39",
    sum(v for y, m, n, f, d, v in wide_july if f == 1), 9645.39, tol=0.05)

result("导入引擎")
