# -*- coding: utf-8 -*-
"""报表引擎测试：
1. 月度报表（当月）vs 26年7月报表 缓存值
2. 经营预测3 复合期间 vs 修正后的口径（8-12月同期、全年）
3. Excel 导出检查
"""
import openpyxl

from app import reports
from 测试helpers import chk, result, use_test_db

conn = use_test_db()

# 确保 26年预算 / 25实际同期(规划2) 已导入（达成率、8-12月同期需要）
from app import importers
MAIN = r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx"
PLAN2 = r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx"
if conn.execute("SELECT COUNT(*) FROM facts WHERE year=2026 AND dtype_id=2").fetchone()[0] == 0:
    tpl_bp = conn.execute("SELECT * FROM source_templates WHERE name='26年预算'").fetchone()
    p = importers.parse_source(conn, tpl_bp, MAIN)
    importers.commit_import(conn, p, "底稿(1).xlsx 26年预算")
    print(f"[自动导入 26年预算: {p['summary']['条数']} 条, 警告 {len(p['warnings'])}]")
if conn.execute("SELECT COUNT(*) FROM facts WHERE year=2025 AND month=8").fetchone()[0] == 0:
    tpl_spl = conn.execute("SELECT * FROM source_templates WHERE name='25实际同期(规划2)'").fetchone()
    p2 = importers.parse_source(conn, tpl_spl, PLAN2)
    importers.commit_import(conn, p2, "规划2 25实际同期")
    print(f"[自动导入 25实际同期(规划2): {p2['summary']['条数']} 条, 警告 {len(p2['warnings'])}]")
    for w in p2["warnings"][:5]:
        print("   -", w)


# ---------- 1. 月度报表渲染 ----------
tpl = conn.execute("SELECT * FROM report_templates WHERE name='月度报表'").fetchone()
t = reports.render(conn, tpl, 2026, 7)
print("=== 月度报表 7月: 行数", len(t["rows"]), "列数", len(t["columns"]), "===")
# 行序：0=KA合计 1=全国连锁 2=区域连锁 3=智屏 4=智屏-全国连锁 5=智屏-区域连锁 ...（重名行按位置取）
row_i = {r["entity"]: i for i, r in enumerate(t["rows"])}
col = {f"{c['field']}|{c['measure']}": i for i, c in enumerate(t["columns"])}


def cells_of(name):
    return t["rows"][row_i[name]]["cells"]

wbv = openpyxl.load_workbook(r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx", data_only=True)
ws = wbv["26年7月报表"]

# 原表: 行6=智屏, B=收入实际, D=收入达成率, E=同期, F=同比, G=销量实际, L=ASP实际, Q=毛利率实际, V=净利实际
checks = [
    ("智屏 收入 实际", cells_of("智屏")[col["收入|实际"]], ws["B6"].value),
    ("智屏 收入 BP达成率", cells_of("智屏")[col["收入|BP达成率"]], ws["D6"].value, 0.005),
    ("智屏 收入 同期", cells_of("智屏")[col["收入|同期"]], ws["E6"].value),
    ("智屏 收入 同比", cells_of("智屏")[col["收入|同比"]], ws["F6"].value, 0.005),
    ("智屏 销量 实际", cells_of("智屏")[col["销量|实际"]], ws["G6"].value),
    ("智屏 ASP 实际", cells_of("智屏")[col["ASP|实际"]], ws["L6"].value, 0.5),
    ("智屏 端到端毛利率 实际", cells_of("智屏")[col["端到端毛利率|实际"]], ws["Q6"].value, 0.0005),
    ("智屏 端到端净利 实际", cells_of("智屏")[col["端到端利润|实际"]], ws["V6"].value, 0.05),
    ("KA合计 收入 实际", cells_of("KA合计")[col["收入|实际"]], ws["B3"].value, 0.05),
    ("全国连锁 收入 实际", cells_of("全国连锁")[col["收入|实际"]], ws["B4"].value, 0.05),
    ("区域连锁 收入 实际", cells_of("区域连锁")[col["收入|实际"]], ws["B5"].value, 0.05),
    ("智屏-全国连锁 收入 实际", t["rows"][row_i["智屏"] + 1]["cells"][col["收入|实际"]], ws["B7"].value, 0.05),
]
for c in checks:
    label, actual, expect = c[0], c[1], c[2]
    tol = c[3] if len(c) > 3 else 0.02
    chk(label, actual, expect, tol)

# ---------- 2. 经营预测3 复合期间 ----------
print("\n=== 经营预测3（复合期间）===")
tpl3 = conn.execute("SELECT * FROM report_templates WHERE name='经营预测3'").fetchone()
t3 = reports.render(conn, tpl3, 2026, 7)
row3 = {r["entity"] + "-" + r["field"]: r["cells"] for r in t3["rows"]}
col3 = {f"{c['group']}|{c['measure']}": i for i, c in enumerate(t3["columns"])}

# 智屏 收入 全年(1-7实际+8-12预测) = 累计实际 + 8-12预测
cum = row3["智屏-收入"][col3["1-7月累计|实际"]]
full = row3["智屏-收入"][col3["全年累计|实际"]]
s812 = row3["智屏-收入"][col3["8-12月|预测"]]
chk("智屏 全年 = 1-7累计 + 8-12预测", full, (cum if cum else 0) + (s812 if s812 else 0))
# 智屏 全年BP达成率 = 全年实际/全年BP（复合期间口径，非各月达成率之和）
rate = row3["智屏-收入"][col3["全年累计|BP达成率"]]
bp12 = row3["智屏-收入"][col3["全年累计|BP"]]
if full and bp12:
    chk("智屏 全年BP达成率 = 全年实际/全年BP", rate, full / bp12, 1e-9)
# 8-12月同期 = 25年8-12月实际（真值只在规划2 的 25实际 AS..AW）
w25 = openpyxl.load_workbook(PLAN2, data_only=True)["25实际"]
spl_812 = sum(w25[f"{c}3"].value or 0 for c in ["AS", "AT", "AU", "AV", "AW"])
chk("智屏 8-12月同期 = 规划2 25实际 AS..AW 之和", row3["智屏-收入"][col3["8-12月|同期"]], spl_812, 0.05)
# 8-12月同比 = 8-12月预测/同期 - 1
yoy812 = row3["智屏-收入"][col3["8-12月|同比"]]
if s812 and spl_812:
    chk("智屏 8-12月同比 = 预测/同期-1", yoy812, s812 / spl_812 - 1, 1e-9)

# ---------- 3. Excel 导出 ----------
print("\n=== Excel 导出 ===")
import os
out = os.path.join(os.path.dirname(__file__), "data", "输出", "2026年7月月度报表.xlsx")
os.makedirs(os.path.dirname(out), exist_ok=True)
reports.export_report(conn, tpl, 2026, 7, out)
print("  已导出:", out)
wbo = openpyxl.load_workbook(out)
wso = wbo.active
print(f"  维度: {wso.max_row}行 x {wso.max_column}列 | 冻结: {wso.freeze_panes}")
print("  行1:", [wso.cell(row=1, column=c).value for c in range(1, 6)])
print("  行2:", [wso.cell(row=2, column=c).value for c in range(1, 6)])
print("  行3:", [wso.cell(row=3, column=c).value for c in range(1, 6)])
print("  合并:", list(wso.merged_cells.ranges)[:6])
print("  C3 格式:", wso["C3"].number_format, "| E3 格式:", wso["E3"].number_format, "| A3:", wso["A3"].value, "| B3:", wso["B3"].value)
chk("导出 C3(KA合计收入) = 原表 B3", wso["C3"].value, ws["B3"].value, 0.05)


# ---------- 4. 经营费用分析（渠道在列模式 + 费率区 + measure_map） ----------
tpl4 = conn.execute("SELECT * FROM report_templates WHERE name='经营费用分析-累计'").fetchone()
t4 = reports.render(conn, tpl4, 2026, 7)
chk("费用分析 列数 18 (9渠道×2口径)", len(t4["columns"]), 18)
chk("费用分析 行数 18 (收入+金额8+费率9)", len(t4["rows"]), 18)
chk("费用分析 首行实体=收入", t4["rows"][0]["entity"], "收入")
# 期望值为引擎对拍源表后确认的正确值（源表聚合列缓存旧不可靠，如 26年经营费用 B55 与其成分矛盾；
# 智屏 1 月收入按用户裁决以明细文件为准，故 166777.56 ≠ 源表 C4 缓存 166777.26）
chk("收入行 实际合计", t4["rows"][0]["cells"][0], 166777.55905954013, 0.01)
chk("收入行 BP比合计 = BP原值(非差值)", t4["rows"][0]["cells"][9], 190918.38259982655, 0.01)
# 汇总费用 = 固定性+变动性+专利认证（固定性=固定+研发+财务，变动性=变动-研发-财务-专利认证；
# 自洽；源表 经营费用分析 C5 引用聚合块旧值 41003.18，叶子明细真值 = 41019.13）
pa = conn.execute(
    "SELECT COALESCE(SUM(f.value),0) FROM facts f JOIN fields fl ON fl.id=f.field_id "
    "WHERE fl.name='专利认证' AND f.year=2026").fetchone()[0]
chk("汇总费用 = 固定性+变动性+专利认证", t4["rows"][1]["cells"][0],
    t4["rows"][2]["cells"][0] + t4["rows"][8]["cells"][0] + pa, 0.01)
chk("费率_汇总费用 = 费用总额÷收入", t4["rows"][9]["cells"][0], t4["rows"][1]["cells"][0] / t4["rows"][0]["cells"][0], 1e-9)
chk("费率_固定性费用 智屏连锁", t4["rows"][10]["cells"][1], 6381.7476791488 / 90737.185830787, 0.000001)
# 导出：表头行2=渠道名、行1=口径组、B列=科目名、A列=分组
out4 = os.path.join(os.environ.get("TEMP", "."), "费用分析测试.xlsx")
reports.export_report(conn, tpl4, 2026, 7, out4)
wso4 = openpyxl.load_workbook(out4)["2026年7月报表"]
chk("导出 表头行1 组=实际", wso4["C1"].value, "实际")
chk("导出 表头行2 首列=合计", wso4["C2"].value, "合计")
chk("导出 表头行2 渠道列名", wso4["D2"].value, "智屏-连锁渠道")
chk("导出 A列 分组=收入", wso4["A3"].value, "收入")
chk("导出 A列 费率分组合并", wso4["A12"].value, "费率")
chk("导出 B列 费率行科目=汇总费用", wso4["B12"].value, "汇总费用")
chk("导出 费率单元格 百分比格式", wso4["C12"].number_format, "0.0%")


# ---------- 5. 计算过程求值树 + 钻取信息 ----------
from app import calc as calcm

def explain_of(tpl_name, row_name, field_name, measure_name):
    """渲染报表定位单元格 -> explain 求值树（支持 实体在行/渠道在列 两种模板）"""
    tt = conn.execute("SELECT * FROM report_templates WHERE name=?", (tpl_name,)).fetchone()
    tab = reports.render(conn, tt, 2026, 7)
    rows_have_field = tab["rows"][0].get("field") is not None
    if rows_have_field:
        # 新式/渠道在列：行 = 实体×字段，列按口径
        ri = next(i for i, r in enumerate(tab["rows"])
                  if r["entity"] == row_name and r.get("field") == field_name)
        ci = next(i for i, c in enumerate(tab["columns"]) if c["measure"] == measure_name)
    else:
        # 旧式：行 = 实体，字段在列
        ri = next(i for i, r in enumerate(tab["rows"]) if r["entity"] == row_name)
        ci = next(i for i, c in enumerate(tab["columns"])
                  if c.get("field") == field_name and c["measure"] == measure_name)
    fname = tab["rows"][ri].get("field") or tab["columns"][ci].get("field")
    fid = conn.execute("SELECT id FROM fields WHERE name=?", (fname,)).fetchone()["id"]
    nids = tab["rows"][ri].get("nids") or tab["columns"][ci].get("channel_nids")
    return calcm.explain_period_measure(conn, 2026, tab["columns"][ci]["segments"], nids, fid,
                                        tab["columns"][ci]["measure_row"],
                                        mixed=(tab["columns"][ci]["measure"] == "实际"))

tree = explain_of("月度报表", "智屏", "端到端利润", "实际")
chk("explain 端到端利润 10 成分", len(tree["children"]), 10)
chk("explain 费用总额子项 = 固定+变动", len([c for c in tree["children"] if c["name"] == "费用总额"][0]["children"]), 2)
chk("explain 端到端利润 值 = 引擎渲染值", tree["value"],
    [r["cells"][20] for r in t["rows"] if r["entity"] == "智屏"][0], 0.01)
# 公式口径：达成率 = 实际 / BP
tree2 = explain_of("月度报表", "KA合计", "收入", "BP达成率")
chk("explain 达成率 公式", tree2["formula"], "实际 / BP")
chk("explain 达成率 = 实际÷BP", tree2["value"],
    tree2["children"][0]["value"] / tree2["children"][1]["value"], 1e-9)
# 渠道在列（费用分析）：费率_固定性费用 = 固定性费用/收入
tree3 = explain_of("经营费用分析-累计", "费率", "费率_固定性费用", "实际")
chk("explain 费率_固定性费用 = 固定性/收入", tree3["value"],
    tree3["children"][0]["value"] / tree3["children"][1]["value"], 1e-9)
# 钻取信息（多段结构）：同期列 -> 去年实际段；达成率 -> 实际+BP 双类型合并段
tplc = conn.execute("SELECT * FROM report_templates WHERE name='当年累计报表'").fetchone()
tc = reports.render(conn, tplc, 2026, 7)
dc = tc["columns"][3]["drill"]
chk("累计报表 同期列 drill 段数", len(dc["segments"]), 1)
chk("累计报表 同期列 drill.year = 2025", dc["segments"][0]["year"], 2025)
chk("累计报表 同期列 drill 精确区间 [1,7]", f"{dc['segments'][0]['months'][0]}-{dc['segments'][0]['months'][1]}", "1-7")
chk("累计报表 达成率 drill 双类型合并", len(tc["columns"][2]["drill"]["segments"]), 1)
chk("累计报表 达成率 drill 类型数", len(tc["columns"][2]["drill"]["segments"][0]["dtype_ids"]), 2)
chk("月度报表 行带 nids", len(t["rows"][0]["nids"]), 4)
# 复合期间钻取：预测3 全年累计|实际 = 1-7实际 + 8-12预测 两段（各按段自身 dtype）
tplp = conn.execute("SELECT * FROM report_templates WHERE name='经营预测3'").fetchone()
tp = reports.render(conn, tplp, 2026, 7)
d_year = tp["columns"][55]["drill"]
chk("预测3 全年实际 drill 段数=2", len(d_year["segments"]), 2)
chk("预测3 全年实际 段1 区间+类型", f"{d_year['segments'][0]['year']}|{d_year['segments'][0]['months'][0]}-{d_year['segments'][0]['months'][1]}|{d_year['segments'][0]['dtype_ids']}",
    "2026|1-7|[1]")
chk("预测3 全年实际 段2 区间+类型", f"{d_year['segments'][1]['year']}|{d_year['segments'][1]['months'][0]}-{d_year['segments'][1]['months'][1]}|{d_year['segments'][1]['dtype_ids']}",
    "2026|8-12|[3]")
# 同比钻取：公式变量 实际 + 虚类型 同期 -> 今年实际段 + 去年实际段（不丢分母）
d_yoy = tp["columns"][54]["drill"]
chk("预测3 全年同比 drill 段数=2", len(d_yoy["segments"]), 2)
chk("预测3 同比 含去年段", any(s["year"] == 2025 and s["dtype_ids"] == [1] for s in d_yoy["segments"]), True)
chk("预测3 同比 含今年混合段", any(s["year"] == 2026 and s["dtype_ids"] == [3] for s in d_yoy["segments"]), True)

# ---------- 6. 列 API 序列化（回归：费用分析表头曾丢失渠道名，全部显示成口径名） ----------
tplf = conn.execute("SELECT * FROM report_templates WHERE name='经营费用分析-当月'").fetchone()
tf = reports.render(conn, tplf, 2026, 7)
af = reports.api_columns(tf["columns"])
channels = [c["channel_name"] for c in af if "channel_name" in c]
chk("费用分析 18 列全部带渠道名", len(channels), 18)
chk("渠道序: 合计/智屏-连锁/智屏-区域", channels[0] + "|" + channels[1] + "|" + channels[2],
    "合计|智屏-连锁渠道|智屏-区域连锁")
chk("渠道序: BP比组首列同为 合计", channels[9], "合计")
chk("费用分析 列带 channel_nids", len(af[0]["channel_nids"]), len(tf["columns"][0]["channel_nids"]), 0)
tpl1 = conn.execute("SELECT * FROM report_templates WHERE name='月度报表'").fetchone()
a1 = reports.api_columns(reports.render(conn, tpl1, 2026, 7)["columns"])
chk("月度报表 列不带 channel_name", all("channel_name" not in c for c in a1), True)
chk("月度报表 列含口径与期间", a1[0]["measure"] + "|" + a1[0]["period"], "实际|当月")

result("报表引擎")
