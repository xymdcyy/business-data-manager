# -*- coding: utf-8 -*-
"""数据迁移：一次导入全部来源模板（25 全年 + 26 至今）。
用法：uv run python 迁移数据.py [--yes]
不带 --yes 时仅预览；带 --yes 确认入库。
"""
import sys

from app import importers, seed
from app.db import get_conn, init_db

FILES = {
    1: r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx",
    2: r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx",
    3: r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx",
    4: r"E:\10 经营分析报告\7月报告\各品类指标表\1 智屏指标表\26年7月KA渠道智屏TCL数据库.xlsx",
    5: r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx",
    6: r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx",
    7: r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx",
}

confirm = "--yes" in sys.argv
init_db()
conn = get_conn()
# 空库时先播种（幂等：元数据/模板/渠道树），保证全新环境一条命令可迁移
seed.seed_all(conn, FILES[1])
print("数据迁移：", "【确认入库】" if confirm else "【仅预览】\n")
total = 0
for tid, path in FILES.items():
    tpl = conn.execute("SELECT * FROM source_templates WHERE id=?", (tid,)).fetchone()
    parsed = importers.parse_source(conn, tpl, path)
    total += parsed["summary"]["条数"]
    print(f"[{tpl['name']}] 解析 {parsed['summary']['条数']} 条 | "
          f"产品分布 {parsed['summary']['产品分布']} | 警告 {len(parsed['warnings'])}")
    for w in parsed["warnings"][:8]:
        print(f"    ⚠ {w}")
    if confirm:
        res = importers.commit_import(conn, parsed, tpl["name"])
        print(f"    ✅ 入库: 写入 {res['写入']} 覆盖 {res['覆盖']}")
    else:
        conflicts = importers.find_conflicts(conn, parsed["items"])
        print(f"    ⚠ 与库中冲突 {len(conflicts)} 条（--yes 后覆盖）")
print(f"\n合计解析 {total} 条事实。")
if not confirm:
    print("如需入库请加 --yes 重跑。")
