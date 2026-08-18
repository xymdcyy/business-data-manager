# -*- coding: utf-8 -*-
"""计算引擎冒烟测试：手工造数验证 聚合/公式/口径/任意组合 四类核心逻辑（临时库副本，不碰生产库）"""
from app import calc
from 测试helpers import chk, result, use_test_db

conn = use_test_db()
cur = conn.cursor()

# ---- 造数：智屏叶子 2026-07 收入(字段1)/销量(字段2)，实际(类型1)/BP(类型2) ----
# 节点: 1=智屏(根) 2=连锁(聚合) 3=连锁-连锁 4=连锁-苏宁零售云 5=连锁-20渠道 6=区域连锁
cur.execute("DELETE FROM facts")
test_data = [
    # (年, 月, 节点, 字段, 类型, 值)
    (2026, 7, 3, 1, 1, 10.0),   # 连锁-连锁 收入 实际
    (2026, 7, 4, 1, 1, 20.0),   # 苏宁零售云 收入 实际
    (2026, 7, 5, 1, 1, 30.0),   # 20渠道 收入 实际
    (2026, 7, 6, 1, 1, 40.0),   # 区域连锁 收入 实际
    (2026, 7, 3, 2, 1, 1.0),    # 销量
    (2026, 7, 4, 2, 1, 2.0),
    (2026, 7, 5, 2, 1, 3.0),
    (2026, 7, 6, 2, 1, 4.0),
    (2026, 7, 3, 1, 2, 8.0),    # 收入 BP（只给两个叶子，测缺数据按0）
    (2026, 7, 4, 1, 2, 12.0),
    (2026, 7, 3, 2, 2, 1.0),    # 销量 BP
    (2026, 7, 4, 2, 2, 2.0),
]
cur.executemany(
    "INSERT INTO facts (year, month, node_id, field_id, dtype_id, value) VALUES (?,?,?,?,?,?)",
    test_data,
)
conn.commit()

print("=== 1. 树聚合（收入 实际）===")
chk("智屏根(1) = 10+20+30+40", calc.field_value(conn, 2026, 7, [1], 1, 1), 100)
chk("连锁聚合(2) = 10+20+30", calc.field_value(conn, 2026, 7, [2], 1, 1), 60)
chk("区域连锁(6) 单叶", calc.field_value(conn, 2026, 7, [6], 1, 1), 40)
chk("空调(7) 无数据", calc.field_value(conn, 2026, 7, [7], 1, 1), None)

print("=== 2. 比率公式（ASP = 收入/销量，聚合后除）===")
chk("智屏 ASP = 100/10", calc.field_value(conn, 2026, 7, [1], 3, 1), 10)
chk("连锁 ASP = 60/6", calc.field_value(conn, 2026, 7, [2], 3, 1), 10)
# 加权正确性：区域连锁收入40 销量4 -> 10；但若只给 苏宁零售云(收入20,销量2)+20渠道(收入30,销量3) -> 50/5=10
chk("单叶 ASP", calc.field_value(conn, 2026, 7, [4], 3, 1), 10)

print("=== 3. 口径（measure）===")
m_actual = conn.execute("SELECT * FROM measures WHERE name='实际'").fetchone()
m_rate = conn.execute("SELECT * FROM measures WHERE name='BP达成率'").fetchone()
m_yoy = conn.execute("SELECT * FROM measures WHERE name='同比'").fetchone()
chk("收入 实际(口径dtype)", calc.measure_value(conn, 2026, 7, [1], 1, m_actual), 100)
# 达成率 = 实际/BP：收入 100/20 = 5.0
chk("收入 BP达成率 = 100/20", calc.measure_value(conn, 2026, 7, [1], 1, m_rate), 5.0)
# 同期无数据 -> None
chk("收入 同比（同期缺失）", calc.measure_value(conn, 2026, 7, [1], 1, m_yoy), None)

print("=== 4. 任意组合（selector_sum）===")
# 智屏合计 + 空调合计（空调无数据） = 100
chk("KA合计(智屏1+空调7)", calc.selector_sum(conn, 2026, 7, [1, 7], 1, 1), 100)
# 全国连锁（连锁聚合2 + 空调连锁8，空调无数据） = 60
chk("KA-全国连锁(2+8)", calc.selector_sum(conn, 2026, 7, [2, 8], 1, 1), 60)
# 路径解析
nid = calc.resolve_path(conn, "智屏", ("连锁", "苏宁零售云"))
chk("resolve_path 智屏/连锁/苏宁零售云", nid, 4)

print("=== 5. 同期口径（year_shift=-1）===")
cur.execute(
    "INSERT INTO facts (year, month, node_id, field_id, dtype_id, value) VALUES (2025,7,3,1,1,50.0)"
)
conn.commit()
m_spl = conn.execute("SELECT * FROM measures WHERE name='同期'").fetchone()
chk("收入 同期 = 2025-07 实际", calc.measure_value(conn, 2026, 7, [1], 1, m_spl), 50.0)
chk("收入 同比 = 100/50-1", calc.measure_value(conn, 2026, 7, [1], 1, m_yoy), 1.0)
chk("同比 无同期数据（2025 空调）", calc.measure_value(conn, 2026, 7, [7], 1, m_yoy), None)

result("计算引擎")
