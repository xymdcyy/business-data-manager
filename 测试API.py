# -*- coding: utf-8 -*-
"""API 层测试（不经 HTTP，直接调用 FastAPI handler 函数——覆盖接口参数/序列化类 bug，
而引擎层测试全部绕过 HTTP 会漏检此类问题）：
1. 导入预览/确认的键名契约（回归：前端发 {template_id, file_path}，曾错配 KeyError 500）
2. 报表渲染/计算过程/导出入口
3. 明细多值筛选 + 包含计算值
"""
from app import main
from 测试helpers import chk, result, use_test_db

conn = use_test_db()
MAIN = r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx"

# ---------- 1. 导入接口：键名契约（前端发 template_id/file_path） ----------
tpl = conn.execute("SELECT id, name FROM source_templates WHERE enabled=1").fetchone()
preview = main.api_import_preview({"template_id": tpl["id"], "file_path": MAIN})
chk("import 预览 解析条数 > 0", preview["summary"]["条数"] > 0, True)
res = main.api_import_commit({"template_id": tpl["id"], "file_path": MAIN})
chk("import 入库 写入+覆盖 > 0", res["写入"] + res["覆盖"] > 0, True)
# 缺键（旧前端 {tpl,path}）→ 友好 400 而非 KeyError 500
try:
    main.api_import_commit({"tpl": tpl["id"], "path": MAIN})
    chk("import 缺键报 400", True, False)
except Exception as e:
    chk("import 缺键报 400", getattr(e, "status_code", 0), 400)

# ---------- 2. 报表接口 ----------
t = main.api_report_render(template_id=1, year=2026, month=7)
chk("报表 render 列数=25", len(t["columns"]), 25)
chk("报表 render 行数 > 10", len(t["rows"]) > 10, True)
tree = main.api_report_explain(template_id=1, year=2026, month=7, row=0, col=0)
chk("报表 explain 有值", tree.get("value") is not None, True)
# 费用分析（渠道在列）explain：渠道列单元格可解释
t4 = main.api_report_render(template_id=4, year=2026, month=7)
tree4 = main.api_report_explain(template_id=4, year=2026, month=7, row=0, col=1)
chk("费用分析 explain 渠道列有值", tree4.get("value") is not None, True)

# ---------- 3. 明细接口：多值筛选 + 包含计算值 ----------
r = main.api_details(year=2026, month=7, product_ids="1", dtype_ids="1",
                     channel_ids="2,3", field_ids="1", page_size=10)
chk("明细 多值筛选 total >= 0", r["total"] >= 0, True)
r_all = main.api_details(year=2026, month=7, page_size=1)
r_calc = main.api_details(year=2026, month=7, include_formula=1, page_size=1)
chk("明细 包含计算值 行数更多", r_calc["total"] > r_all["total"], True)

# ---------- 4. 批量求值器 vs 逐格 field_value 对拍（缓存优化不得改变语义） ----------
from app import calc as calcmod
ev = calcmod.FormulaBulkEvaluator(conn)
r_full = main.api_details(year=2026, month=7, include_formula=1, page_size=500)
bad = checked = 0
for d in r_full["rows"]:
    if d["source"] != "计算值":
        continue
    v = calcmod.field_value(conn, d["year"], d["month"], [d["node_id"]], d["field_id"], d["dtype_id"])
    checked += 1
    if v is None or abs(v - d["value"]) > 1e-9:
        bad += 1
chk("批量求值器 vs field_value 一致", bad, 0, 0)
chk("对拍样本数 > 0", checked > 0, True)

result("API")
