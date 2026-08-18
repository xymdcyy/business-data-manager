# -*- coding: utf-8 -*-
"""经营预测3 新模板验证：
1. 查询 7月：45 列（9组×5口径）× 15 行 vs 原表 C..AU 全表逐格
2. 查询 6月：标题动态（1-6月累计/7月预测..12月预测/7-12月/全年）+ 期间正确
"""
import openpyxl

from app import reports
from 测试helpers import chk, result, use_test_db

conn = use_test_db()
PLAN2 = r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx"

tpl = conn.execute("SELECT * FROM report_templates WHERE name='经营预测3'").fetchone()

# ---------- 1. 查询 7月 vs 原表 ----------
print("=== 1. 查询 2026-07 vs 原表全表 ===")
t = reports.render(conn, tpl, 2026, 7)
print(f"  列数: {len(t['columns'])}（期望 60）| 行数: {len(t['rows'])}（期望 15）")
groups = []
for c in t["columns"]:
    if not groups or groups[-1] != c["group"]:
        groups.append(c["group"])
print("  组头:", " | ".join(groups))
expect_groups7 = (["1-7月累计"] + [f"{k}月预测" for k in range(7, 13)]
                  + ["Q3(7-9月)", "Q4(10-12月)", "下半年(7-12月)", "8-12月", "全年累计"])
ok7 = groups == expect_groups7
print("  " + ("✓" if ok7 else "✗") + " 7月组头序列")
if not ok7:
    raise SystemExit(1)

wb = openpyxl.load_workbook(PLAN2, data_only=True)
ws = wb["经营预测 3"]


def src_col_of(ci):
    """渲染列 -> 原表列（重叠部分）：累计5+单月30 对齐原表前 35 列；8-12月/全年（渲染 50..59）对齐末 10 列
    渲染列序：累计5 + 单月30 + Q3/Q4/下半年 15 + 8-12月 5(50-54) + 全年 5(55-59)"""
    if ci <= 34:
        return 3 + ci
    if 50 <= ci <= 59:  # 8-12月 = 原表 AL..AP(38..42)，全年 = AQ..AU(43..47)
        return 3 + ci - 15
    return None  # Q3/Q4/下半年 是新增组，原表无对应


bad = checked = 0
for ri, row in enumerate(t["rows"]):
    src_r = 3 + ri  # 原表行 3-17（实体×字段展开顺序一致）
    for ci, v in enumerate(row["cells"]):
        sc = src_col_of(ci)
        if sc is None:
            continue
        src = ws.cell(row=src_r, column=sc).value
        if not isinstance(src, (int, float)):
            continue
        v = v if v is not None else 0
        if src == 0 and v == 0:
            checked += 1
            continue
        tol = 0.005 if ci % 5 in (2, 4) else 0.05  # 达成率/同比列放宽
        checked += 1
        if abs(v - src) > tol:
            bad += 1
            if bad <= 8:
                print(f"  ✗ 行{src_r} 列{ci} ({row['entity']}-{row['field']} {t['columns'][ci]['measure']}): "
                      f"渲染={round(v, 4)} vs 原表={round(src, 4)}")
print(f"  重叠列核对 {checked} 格，不一致 {bad} 格")
# 基线：源表聚合块缓存陈旧（合计行/智屏 1-6月明细修订等，见注释），当前 107 格已知；
# 差异数超出基线（源表更新为真值后应减少、或回归增多）即失败——不再吞回归信号
if bad > 130:
    raise SystemExit(f"重叠列核对差异 {bad} 超出基线 130，疑似回归（见上方 ✗ 明细）")

# 关键格抽查
row = {r["entity"] + "-" + r["field"]: r["cells"] for r in t["rows"]}
col = {f"{c['group']}|{c['measure']}": i for i, c in enumerate(t["columns"])}
chk("智屏-收入 8月预测 = 原表M6", row["智屏-收入"][col["8月预测|预测"]], ws["M6"].value)
chk("智屏-收入 8-12月累计 = 原表AL6", row["智屏-收入"][col["8-12月|预测"]], ws["AL6"].value)
# 全年：原表缓存含智屏 1-6月 旧值（明细修订差 0.29），tol 放宽；渲染以源表真值为准
chk("智屏-收入 全年 = 原表AQ6", row["智屏-收入"][col["全年累计|实际"]], ws["AQ6"].value, tol=0.35)
chk("智屏-毛利率 全年 = 原表AQ7(加权)", row["智屏-端到端毛利率"][col["全年累计|实际"]], ws["AQ7"].value, tol=0.005)
chk("智屏-净利 8月预测 = 原表M8", row["智屏-端到端利润"][col["8月预测|预测"]], ws["M8"].value)

# ---------- 2. 查询 6月：动态标题与期间 ----------
print("\n=== 2. 查询 2026-06 动态标题 ===")
t6 = reports.render(conn, tpl, 2026, 6)
groups6 = []
for c in t6["columns"]:
    if not groups6 or groups6[-1] != c["group"]:
        groups6.append(c["group"])
print("  组头:", " | ".join(groups6))
expect_groups = (["1-6月累计"] + [f"{k}月预测" for k in range(6, 13)]
                 + ["Q2(4-6月)", "Q3(7-9月)", "Q4(10-12月)",
                    "上半年(1-6月)", "下半年(7-12月)", "7-12月", "全年累计"])
ok6 = groups6 == expect_groups
print("  " + ("✓" if ok6 else "✗") + " 6月组头序列")
if not ok6:
    raise SystemExit(1)
row6 = {r["entity"] + "-" + r["field"]: r["cells"] for r in t6["rows"]}
col6 = {f"{c['group']}|{c['measure']}": i for i, c in enumerate(t6["columns"])}
# 6月查询：7月预测 = 源 26实际(预测) 7月值；全年 = 1-6实际 + 7-12预测
v_7y = row6["智屏-收入"][col6["7月预测|预测"]]
chk("智屏 6月查询的 7月预测 = 源7月预测(14998.4)", v_7y, 14998.4, tol=0.05)
v_12y = row6["智屏-收入"][col6["12月预测|预测"]]
chk("智屏 6月查询的 12月预测 = 源12月预测(31271.1)", v_12y, 31271.1, tol=0.05)

# ---------- 3. 季度/半年度块数值（7月查询，源表真值） ----------
print("\n=== 3. 季度/半年度块（7月查询，源表真值）===")
wb2 = openpyxl.load_workbook(PLAN2, data_only=True)
sp = wb2["26实际(预测)"]  # 块41 智屏: 44=7月 45=8月 46=9月 47=10月 48=11月 49=12月
row7 = {r["entity"] + "-" + r["field"]: r["cells"] for r in t["rows"]}
col7 = {f"{c['group']}|{c['measure']}": i for i, c in enumerate(t["columns"])}
q3 = sum(sp.cell(row=3, column=c).value or 0 for c in range(44, 47))   # 7-9月
q4 = sum(sp.cell(row=3, column=c).value or 0 for c in range(47, 50))   # 10-12月
h2 = sum(sp.cell(row=3, column=c).value or 0 for c in range(44, 50))   # 7-12月
chk("智屏 Q3预测 = 源7-9月和", row7["智屏-收入"][col7["Q3(7-9月)|预测"]], q3)
chk("智屏 Q4预测 = 源10-12月和", row7["智屏-收入"][col7["Q4(10-12月)|预测"]], q4)
chk("智屏 下半年预测 = 源7-12月和", row7["智屏-收入"][col7["下半年(7-12月)|预测"]], h2)
chk("Q3+Q4 = 下半年（勾稽，Q3 已含 7月）", q3 + q4, h2, tol=0.001)
chk("智屏 8-12月 = 源45-49和", row7["智屏-收入"][col7["8-12月|预测"]],
    sum(sp.cell(row=3, column=c).value or 0 for c in range(45, 50)))

result("经营预测3")
