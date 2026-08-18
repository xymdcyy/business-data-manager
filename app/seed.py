# -*- coding: utf-8 -*-
"""种子数据：数据类型 / 产品线 / 指标字段 / 口径 + 从底稿块标签自动发现渠道树
幂等：固定 ID 写入，重复运行安全。
用法：uv run python -m app.seed [底稿路径]
"""
import json
import sys

import openpyxl

from .db import get_conn, init_db

# 底稿块标签 -> 渠道树来源（月度经营主文件）
DEFAULT_SRC = r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿(1).xlsx"
SRC_SHEET = "26年实际"

# 源表块标签中的产品代号 -> 系统产品名
PRODUCT_ALIASES = {"智屏": "智屏", "KT": "空调", "BX": "冰洗", "CIOT": "CIoT"}

# ---------------- 固定种子（ID 稳定，供代码引用） ----------------
# 同期为虚类型：base=实际(1) + year_shift=-1，查询去年同月实际，不存储
DATA_TYPES = [
    (1, "实际", None, 0, "实际发生数据"),
    (2, "BP", None, 0, "预算数据"),
    (3, "预测", None, 0, "预测数据"),
    (4, "同期", 1, -1, "虚类型：取去年同月实际（同比分母）"),
]

PRODUCTS = [
    (1, "智屏", ""),
    (2, "空调", ""),
    (3, "冰洗", ""),
    (4, "CIoT", ""),
]

MEASURES = [
    # dtype 类：直接取字段的某数据类型（year_shift 支持同期）
    (1, "实际", "dtype", 1, 0, None),
    (2, "BP", "dtype", 2, 0, None),
    (3, "预测", "dtype", 3, 0, None),
    (4, "同期", "dtype", 1, -1, None),  # 去年同月实际
    # formula 类：跨数据类型公式（变量=数据类型名，作用于该列字段）
    (5, "BP达成率", "formula", None, 0, "实际 / BP"),
    (6, "同比", "formula", None, 0, "实际 / 同期 - 1"),
    # ASP 专用差值口径（原表 L..P 块：达成率/同比都是单价差，不是百分比）
    (7, "ASP达成率", "formula", None, 0, "实际 - BP"),
    (8, "ASP同比", "formula", None, 0, "实际 - 同期"),
    # 费用分析表「BP比」列 = 实际-BP 差值（同 ASP 差值语义，原表 26年经营费用 D列）
    (9, "BP达成差额", "formula", None, 0, "实际 - BP"),
    # 端到端毛利率/净利 块的同比 = 差值口径（原表 26年累计报表 U/Z 列：个点差、盈亏万差，非百分比）
    (10, "毛利率同比", "formula", None, 0, "实际 - 同期"),
    (11, "净利同比", "formula", None, 0, "实际 - 同期"),
    # 缺口：用户后定，留白
]

# 字段：raw（导入）+ formula（派生，导入校验/明细源计算用）。
# 以主文件 26年实际 A列（行3-40）为权威：源宽表有值的行都是 raw；比率行公式化用于校验
FIELDS = [
    # (id, 名称, kind, formula, 单位, 小数位, 备注)
    (1, "收入", "raw", None, "万元", 0, "26年实际 行3"),
    (2, "销量", "raw", None, "万台", 0, "26年实际 行4"),
    (3, "ASP", "formula", "收入 / 销量", "元", 0, "26年实际 行5；源有值用于校验"),
    (4, "销司端毛利", "raw", None, "万元", 0, "26年实际 行6"),
    (5, "销司端毛利率", "formula", "销司端毛利 / 收入", "%", 1, "26年实际 行7；校验"),
    (6, "端到端毛利", "raw", None, "万元", 0, "26年实际 行8"),
    (7, "端到端毛利率", "formula", "端到端毛利 / 收入", "%", 1, "26年实际 行9；校验"),
    (8, "税金及附加", "raw", None, "万元", 0, "26年实际 行10"),
    (9, "固定费用", "formula", "人力成本 + 日常费用 + 管理费用", "万元", 0, "原表公式 =AV12+AV13+AV14"),
    (10, "人力成本", "raw", None, "万元", 0, "26年实际 行12"),
    (11, "日常费用", "raw", None, "万元", 0, "26年实际 行13"),
    (12, "管理费用", "raw", None, "万元", 0, "26年实际 行14"),
    (13, "变动费用", "formula",
     "激励及佣金 + 储运费 + 市场推广 + 售后服务 + 安装费用 + 品牌费用 + 研发费用 + 财务费用 + 专利认证",
     "万元", 0, "原表公式 =SUM(AV16:AV24)"),
    (14, "激励及佣金", "raw", None, "万元", 0, "26年实际 行16"),
    (15, "储运费", "raw", None, "万元", 0, "26年实际 行17"),
    (16, "市场推广", "raw", None, "万元", 0, "26年实际 行18"),
    (17, "售后服务", "raw", None, "万元", 0, "26年实际 行19"),
    (18, "安装费用", "raw", None, "万元", 0, "26年实际 行20"),
    (19, "品牌费用", "raw", None, "万元", 0, "26年实际 行21"),
    (20, "研发费用", "raw", None, "万元", 0, "26年实际 行22"),
    (21, "财务费用", "raw", None, "万元", 0, "26年实际 行23"),
    (22, "专利认证", "raw", None, "万元", 0, "26年实际 行24"),
    (23, "费用总额", "formula", "固定费用 + 变动费用", "万元", 0, "原表公式 =AV25(固定+变动)"),
    (24, "费用率", "formula", "费用总额 / 收入", "%", 1, "26年实际 行26；校验"),
    (25, "资产减值损失", "raw", None, "万元", 0, "26年实际 行27"),
    (26, "信用减值损失", "raw", None, "万元", 0, "26年实际 行28"),
    (27, "其它业务收支", "raw", None, "万元", 0, "26年实际 行29"),
    (28, "营业外收支", "raw", None, "万元", 0, "26年实际 行30"),
    (29, "所得税", "raw", None, "万元", 0, "26年实际 行31"),
    (30, "销司净利润", "raw", None, "万元", 0, "26年实际 行32"),
    (31, "后台利润", "raw", None, "万元", 0, "26年实际 行33（+后台利润）"),
    (32, "UIP增加利润", "raw", None, "万元", 0, "26年实际 行34（+UIP增加利润）"),
    (33, "端到端利润", "formula",
     "端到端毛利 - 税金及附加 - 费用总额 - 资产减值损失 - 信用减值损失"
     " + 其它业务收支 + 营业外收支 - 所得税 + UIP增加利润 + 后台利润",
     "万元", 0, "原表公式 =BV8-BV10-BV25-BV27-BV28+BV29+BV30-BV31+BV34+BV33"),
    (34, "端到端利润率", "formula", "端到端利润 / 收入", "%", 1, "26年实际 行36；校验"),
    (35, "展台费用", "raw", None, "万元", 0, "26年实际 行37"),
    # 经营费用分析表专用口径：原表固定性=人力/日常/研发/财务/管理（5项），变动性=激励/储运/市场/售后/安装/品牌（6项）
    # 与通用 固定费用(3项)/变动费用(9项) 不同，需独立公式字段
    (36, "固定性费用", "formula", "固定费用 + 研发费用 + 财务费用", "万元", 0, "经营费用分析 固定性=人力/日常/研发/财务/管理"),
    (37, "变动性费用", "formula", "变动费用 - 研发费用 - 财务费用 - 专利认证", "万元", 0, "经营费用分析 变动性=激励/储运/市场/售后/安装/品牌"),
    # 费率区：各费用科目 ÷ 收入（原表 26年经营费用 费率区；单位 % 导出百分比）
    # 注：费率-汇总费用 复用 24「费用率」（同式 费用总额/收入），模板 label 覆盖为「汇总费用」
    (38, "费率_固定性费用", "formula", "固定性费用 / 收入", "%", 1, "经营费用分析 费率-固定性费用"),
    (39, "费率_变动性费用", "formula", "变动性费用 / 收入", "%", 1, "经营费用分析 费率-变动性费用"),
    (40, "费率_激励及佣金", "formula", "激励及佣金 / 收入", "%", 1, "经营费用分析 费率-激励及佣金"),
    (41, "费率_储运费", "formula", "储运费 / 收入", "%", 1, "经营费用分析 费率-储运费"),
    (42, "费率_市场推广", "formula", "市场推广 / 收入", "%", 1, "经营费用分析 费率-市场推广"),
    (43, "费率_售后服务", "formula", "售后服务 / 收入", "%", 1, "经营费用分析 费率-售后服务"),
    (44, "费率_安装费用", "formula", "安装费用 / 收入", "%", 1, "经营费用分析 费率-安装费"),
    (45, "费率_品牌费用", "formula", "品牌费用 / 收入", "%", 1, "经营费用分析 费率-品牌费用"),
]


# ---------------- 导入来源模板（宽表/明细） ----------------
MAIN_FILE = r"E:\10 经营分析报告\{month}月报告\2026年{month}月KA经营分析报告底稿(1).xlsx"
DETAIL_FILES = {
    "智屏": r"E:\10 经营分析报告\7月报告\各品类指标表\1 智屏指标表\26年7月KA渠道智屏TCL数据库.xlsx",
}

SOURCE_TEMPLATES = [
    (1, "26年实际", MAIN_FILE, "26年实际",
     json.dumps({"style": "wide", "year": 2026, "dtype": "实际",
                 # 智屏 1-6月 以明细源为权威（用户裁决修订值），宽表重导入不得覆盖
                 "protected": [{"product": "智屏", "months": [1, 2, 3, 4, 5, 6]}]},
                ensure_ascii=False),
     "主文件实际宽表（1-12月块；智屏 1-6月 受保护=明细源权威）"),
    (2, "26年预算", MAIN_FILE, "26年预算",
     json.dumps({"style": "wide", "year": 2026, "dtype": "BP"}, ensure_ascii=False),
     "主文件预算宽表"),
    (3, "25实际", MAIN_FILE, "25实际",
     json.dumps({"style": "wide", "year": 2025, "dtype": "实际"}, ensure_ascii=False),
     "主文件25年实际宽表（同比/同期基础）"),
    (5, "25实际同期(规划2)",
     r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx", "25实际",
     json.dumps({"style": "wide", "year": 2025, "dtype": "实际", "month_mode": "label"},
                ensure_ascii=False),
     "规划2 的 25实际 同期专用表（位置错位、标签为真，month_mode=label；8-12月真值）"),
    (6, "26实际(预测)",
     r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx", "26实际(预测)",
     json.dumps({"style": "wide", "year": 2026, "dtype": "预测", "month_mode": "label"},
                ensure_ascii=False),
     "规划2 的 26实际(预测) 预测宽表（位置错位、标签为真；7-12月预测值）"),
    (7, "26年预算(规划2)",
     r"E:\10 经营分析报告\7月报告\2026年7月KA经营分析报告底稿（规划2）-副本.xlsx", "26年预算",
     json.dumps({"style": "wide", "year": 2026, "dtype": "BP", "month_mode": "label",
                 # 源表标签纠错：冰洗 223 列标签误写为 KA-BX-连锁（应为 KA-BX-连锁-连锁，
                 # 值与 210 列相同系复制粘贴），否则连锁-连锁的 BP 无法入库
                 "label_fix": {"223": "KA-BX-连锁-连锁"}},
                ensure_ascii=False),
     "规划2 的 26年预算 宽表（位置错位、标签为真；主文件预算只到7月，8-12月真值在此）"),
    (4, "智屏实际明细", DETAIL_FILES["智屏"], "底稿",
     json.dumps({
         "style": "detail",
         "year_col": "A", "month_col": "B", "dtype_col": "C",
         "product_col": "D", "channel_col": "G",
         "product_map": {"智屏-TCL": "智屏"},
         "channel_map": {"连锁": ["连锁", "连锁"], "区域连锁": ["区域连锁"]},
         "filters": {"J": "KA"},
         "col_map": {
             "K": "收入", "L": "销量", "M": "销司端毛利", "N": "端到端毛利",
             "U+Y": "激励及佣金", "Z+AA": "储运费", "AB": "市场推广",
             "AC+AE": "品牌费用", "AG": "人力成本", "AI+AJ+AK": "日常费用",
             "AL": "研发费用", "AM": "财务费用", "AN": "管理费用",
             "AT": "税金及附加", "AU": "信用减值损失", "AV": "资产减值损失",
             "AW": "其它业务收支", "AX": "营业外收支", "AY": "所得税",
         },
     }, ensure_ascii=False),
     "智屏实际明细长表（各品类指标表）"),
]


# ---------------- 报表模板 ----------------
# 月度报表/当年累计报表 共享的行实体（15 行：KA合计/全国连锁/区域连锁 + 4产品×其中渠道）
MONTHLY_ROWS = [
    {"name": "KA合计", "sel": {"all": True}},
    {"name": "全国连锁", "sel": {"all": True, "path": ["连锁"]}},
    {"name": "区域连锁", "sel": {"all": True, "path": ["区域连锁"]}},
    {"name": "智屏", "sel": {"p": "智屏"}},
    {"name": "其中：全国连锁", "sel": {"p": "智屏", "path": ["连锁"]}},
    {"name": "其中：区域连锁", "sel": {"p": "智屏", "path": ["区域连锁"]}},
    {"name": "空调", "sel": {"p": "空调"}},
    {"name": "其中：全国连锁", "sel": {"p": "空调", "path": ["连锁"]}},
    {"name": "其中：区域连锁", "sel": {"p": "空调", "path": ["区域连锁"]}},
    {"name": "冰洗", "sel": {"p": "冰洗"}},
    {"name": "其中：全国连锁", "sel": {"p": "冰洗", "path": ["连锁"]}},
    {"name": "其中：区域连锁", "sel": {"p": "冰洗", "path": ["区域连锁"]}},
    {"name": "CIoT", "sel": {"p": "CIoT"}},
    {"name": "其中：全国连锁", "sel": {"p": "CIoT", "path": ["连锁"]}},
    {"name": "其中：区域连锁", "sel": {"p": "CIoT", "path": ["区域连锁"]}},
]


def _monthly_columns(period):
    """月度报表/当年累计报表 列：5 指标块 × 5 口径（period 区分 当月/累计）
    端到端毛利率/净利 的同比 = 差值口径（原表：个点差 / 盈亏万差），收入/销量 同比保持百分比"""
    return [
        *[{"group": "收入（万元）", "field": "收入", "measure": m, "period": period}
          for m in ("实际", "BP", "BP达成率", "同期", "同比")],
        *[{"group": "销量（万台）", "field": "销量", "measure": m, "period": period}
          for m in ("实际", "BP", "BP达成率", "同期", "同比")],
        *[{"group": "ASP（元）", "field": "ASP", "measure": m, "period": period}
          for m in ("实际", "BP", "ASP达成率", "同期", "ASP同比")],
        *[{"group": "端到端毛利率", "field": "端到端毛利率", "measure": m, "period": period}
          for m in ("实际", "BP", "BP达成率", "同期", "毛利率同比")],
        *[{"group": "端到端净利（万元）", "field": "端到端利润", "measure": m, "period": period}
          for m in ("实际", "BP", "BP达成率", "同期", "净利同比")],
    ]


# 经营费用分析 渠道列（9 组：合计 + 4产品×{连锁渠道, 区域连锁}），当月/累计模板共用
EXPENSE_CHANNELS = [
    {"name": "合计", "sel": {"all": True}},
    {"name": "智屏-连锁渠道", "sel": {"p": "智屏", "path": ["连锁"]}},
    {"name": "智屏-区域连锁", "sel": {"p": "智屏", "path": ["区域连锁"]}},
    {"name": "空调-连锁渠道", "sel": {"p": "空调", "path": ["连锁"]}},
    {"name": "空调-区域连锁", "sel": {"p": "空调", "path": ["区域连锁"]}},
    {"name": "冰洗-连锁渠道", "sel": {"p": "冰洗", "path": ["连锁"]}},
    {"name": "冰洗-区域连锁", "sel": {"p": "冰洗", "path": ["区域连锁"]}},
    {"name": "CIoT-连锁渠道", "sel": {"p": "CIoT", "path": ["连锁"]}},
    {"name": "CIoT-区域连锁", "sel": {"p": "CIoT", "path": ["区域连锁"]}},
]


def _expense_blocks(period):
    """经营费用分析 列块：实际 / BP比（=实际-BP 差值）两组，渠道在列"""
    return [
        {"title": "实际", "period": period, "measures": ["实际"], "channels": EXPENSE_CHANNELS},
        {"title": "BP比", "period": period, "measures": ["BP达成差额"], "channels": EXPENSE_CHANNELS},
    ]


def _expense_row_groups():
    """经营费用分析 行：收入 / 金额区 / 费率区（费率_汇总费用 复用通用 费用率 字段，label 覆盖）"""
    return [
        {"label": "收入", "fields": [{"name": "收入", "label": "", "measure_map": {"BP达成差额": "BP"}}]},
        {"label": "金额", "fields": [
            {"name": "费用总额", "label": "汇总费用"},
            {"name": "固定性费用", "label": "固定性费用"},
            {"name": "人力成本", "label": "人力成本"},
            {"name": "日常费用", "label": "日常费用"},
            {"name": "研发费用", "label": "研发费用"},
            {"name": "财务费用", "label": "财务费用"},
            {"name": "管理费用", "label": "管理费用"},
            {"name": "变动性费用", "label": "变动性费用"},
        ]},
        {"label": "费率", "fields": [
            {"name": "费用率", "label": "汇总费用"},
            {"name": "费率_固定性费用", "label": "固定性费用"},
            {"name": "费率_变动性费用", "label": "变动性费用"},
            {"name": "费率_激励及佣金", "label": "激励及佣金"},
            {"name": "费率_储运费", "label": "储运费"},
            {"name": "费率_市场推广", "label": "市场推广"},
            {"name": "费率_售后服务", "label": "售后服务"},
            {"name": "费率_安装费用", "label": "安装费"},
            {"name": "费率_品牌费用", "label": "品牌费用"},
        ]},
    ]


REPORT_TEMPLATES = [
    (1, "月度报表",
     json.dumps({
         "title": "品类/渠道",
         "rows": MONTHLY_ROWS,
         "columns": _monthly_columns("当月"),
     }, ensure_ascii=False),
     "月度报表（当月/5列口径×5指标，对应 26年7月报表 结构）"),
    (2, "经营预测3",
     json.dumps({
         "title": "品类",
         # 行 = 实体 × 字段展开（每实体连续 3 行：收入/毛利率/净利，A 列共享实体标签）
         "rows": [
             {"name": "合计", "sel": {"all": True}},
             {"name": "智屏", "sel": {"p": "智屏"}},
             {"name": "空调", "sel": {"p": "空调"}},
             {"name": "冰洗", "sel": {"p": "冰洗"}},
             {"name": "CIoT", "sel": {"p": "CIoT"}},
         ],
         # 字段显示名：源 A 列叫"端到端利润"，报表沿用原表习惯显示"端到端净利"
         # 毛利率/净利行的 同比 列按原表为差值口径（个点差/盈亏万差），经 measure_map 行级覆盖
         "fields": [
             {"name": "收入", "label": "收入"},
             {"name": "端到端毛利率", "label": "端到端毛利率", "measure_map": {"同比": "毛利率同比"}},
             {"name": "端到端利润", "label": "端到端净利", "measure_map": {"同比": "净利同比"}},
         ],
         # 列 = 期间组 × 5 口径。用户要求：累计实际 + 单月预测 + 季度预测 + 半年度预测 + 全年预测
         # （本质 = 任意月份段×数据类型组合：months/quarters/halves/segments 块类型通用）
         # {m}=查询月 {k}=单月 {q}/{qs}/{qe}=季度 {h}/{hs}/{he}/{h_label}=半年
         "column_blocks": [
             {"title": "1-{m}月累计", "period": "1-{m}月实际",
              "measures": ["实际", "BP", "BP达成率", "同期", "同比"]},
             {"title": "{k}月预测", "period": "{k}月预测",
              "measures": ["预测", "BP", "BP达成率", "同期", "同比"],
              "months": {"start": "{m}", "end": 12}},  # 单月预测：从当月预测组起（原表含 7月预测组）
             {"title": "Q{q}({qs}-{qe}月)", "period": "{qs}-{qe}月预测",
              "measures": ["预测", "BP", "BP达成率", "同期", "同比"],
              "quarters": {"start": "{m_q}"}},         # 季度预测：从查询月所在季度起
             {"title": "{h_label}({hs}-{he}月)", "period": "{hs}-{he}月预测",
              "measures": ["预测", "BP", "BP达成率", "同期", "同比"],
              "halves": {"start": "{m_h}"}},           # 半年度预测：从查询月所在半年起
             {"title": "{m+1}-12月", "period": "{m+1}-12月预测",
              "measures": ["预测", "BP", "BP达成率", "同期", "同比"]},
             {"title": "全年累计", "period": "1-{m}月实际+{m+1}-12月预测",
              "measures": ["实际", "BP", "BP达成率", "同期", "同比"]},
         ],
     }, ensure_ascii=False),
     "经营预测3（原表结构：5实体×3字段行，期间组动态化，列=期间×5口径）"),
    (3, "当年累计报表",
     json.dumps({
         # 与模板1「月度报表」同样格式（行/列结构一致），仅期间改为当年 1~{m} 累计
         "title": "品类/渠道",
         "rows": MONTHLY_ROWS,
         "columns": _monthly_columns("1-{m}月实际"),
     }, ensure_ascii=False),
     "当年累计报表（与月度报表同格式，期间=当年 1-查询月累计，对应 26年累计报表 sheet）"),
    # 经营费用分析（对应「经营费用分析 (当月)」/「经营费用分析（累计）」sheet）：
    # 列 = 渠道×口径（渠道在列，9 组：合计+4产品×{连锁渠道,区域连锁}），行 = 费用科目（金额区+费率区）
    # 口径两组：实际 / BP比（=BP达成差额 = 实际-BP 差值；收入行经 measure_map 显示 BP 原值）
    # row_groups 模式：A 列=分组标签（收入/金额/费率），B 列=科目名；列块 channels 展开渠道
    (4, "经营费用分析-当月",
     json.dumps({
         "title": "费用项目",
         "row_groups": _expense_row_groups(),
         "column_blocks": _expense_blocks("当月"),
     }, ensure_ascii=False),
     "经营费用分析-当月（列=9渠道×2口径，行=费用科目+费率，对应 经营费用分析 (当月) sheet）"),
    (5, "经营费用分析-累计",
     json.dumps({
         "title": "费用项目",
         "row_groups": _expense_row_groups(),
         "column_blocks": _expense_blocks("1-{m}月实际"),
     }, ensure_ascii=False),
     "经营费用分析-累计（列=9渠道×2口径，行=费用科目+费率，对应 经营费用分析（累计）sheet）"),
]


def seed_reports(conn):
    cur = conn.cursor()
    cur.executemany(
        "INSERT OR REPLACE INTO report_templates (id, name, config_json, note) VALUES (?,?,?,?)",
        REPORT_TEMPLATES,
    )
    conn.commit()


def seed_templates(conn):
    cur = conn.cursor()
    cur.executemany(
        "INSERT OR REPLACE INTO source_templates (id, name, path_pattern, sheet_name, config_json, note) "
        "VALUES (?,?,?,?,?,?)",
        SOURCE_TEMPLATES,
    )
    conn.commit()


def seed_meta(conn):
    cur = conn.cursor()
    cur.executemany(
        "INSERT OR REPLACE INTO data_types (id, name, base_dtype_id, year_shift, note) VALUES (?,?,?,?,?)",
        DATA_TYPES,
    )
    cur.executemany("INSERT OR REPLACE INTO product_lines (id, name, note) VALUES (?,?,?)", PRODUCTS)
    cur.executemany(
        "INSERT OR REPLACE INTO measures (id, name, kind, dtype_id, year_shift, formula) VALUES (?,?,?,?,?,?)",
        MEASURES,
    )
    cur.executemany(
        "INSERT OR REPLACE INTO fields (id, name, kind, formula, unit, scale, note) VALUES (?,?,?,?,?,?,?)",
        FIELDS,
    )
    # 清理已从种子移除的字段（仅 formula 字段可能被删——facts 不引用 formula，安全）
    cur.execute("DELETE FROM fields WHERE id NOT IN (%s)" % ",".join("?" * len(FIELDS)),
                [f[0] for f in FIELDS])
    conn.commit()


def discover_channel_trees(src_path):
    """扫描源表行2的'合计'分块，解析块标签（KA-智屏-连锁-苏宁零售云）→ 每产品一组渠道路径。
    返回 {产品: {tuple(渠道路径): 块起始列}} 与跳过列表。
    """
    wb = openpyxl.load_workbook(src_path, data_only=False, read_only=True)
    ws = wb[SRC_SHEET]
    # read_only 模式下 max_column 来自工作表 <dimension> 元素，程序化生成的 xlsx 可能缺失 → None
    max_col = ws.max_column or 0
    starts = [c for c in range(2, max_col + 1) if ws.cell(row=2, column=c).value == "合计"]
    trees = {}
    skipped = []  # 非产品线首段的块（跨产品聚合，如 KA-连锁 → 查询期聚合，不存储）
    for s in starts:
        label = ws.cell(row=1, column=s).value
        if not label or not isinstance(label, str) or not label.startswith("KA-"):
            skipped.append((s, label))
            continue
        path = [seg.strip() for seg in label.split("-")[1:]]
        if not path:
            skipped.append((s, label))
            continue
        prod = path[0]
        if prod not in PRODUCT_ALIASES:
            skipped.append((s, label))
            continue
        chan_path = tuple(path[1:])  # () = 产品合计块
        trees.setdefault(PRODUCT_ALIASES[prod], {})[chan_path] = s
    wb.close()
    return trees, skipped


def seed_trees(conn, trees):
    """按完整路径构建渠道树（幂等：按 (产品,父,名) 查现有节点，缺失才插入，不清空不重建）。
    保证 facts 引用的节点身份稳定；源表新增渠道时补节点，删渠道时旧节点保留（孤儿无害）。"""
    cur = conn.cursor()
    pid_by_name = {p[1]: p[0] for p in PRODUCTS}
    for prod, paths in trees.items():
        pid = pid_by_name[prod]
        root_name = prod
        row = cur.execute(
            "SELECT id FROM channel_nodes WHERE product_id=? AND parent_id IS NULL AND name=?",
            (pid, root_name),
        ).fetchone()
        if row:
            root_id = row["id"]
        else:
            cur.execute(
                "INSERT INTO channel_nodes (product_id, parent_id, name, is_leaf) VALUES (?,NULL,?,0)",
                (pid, root_name),
            )
            root_id = cur.lastrowid
        node_by_path = {(): root_id}
        for chan_path in paths:
            if not chan_path:
                continue
            parent_id = root_id
            for i in range(len(chan_path)):
                prefix = chan_path[: i + 1]
                if prefix in node_by_path:
                    parent_id = node_by_path[prefix]
                    continue
                row = cur.execute(
                    "SELECT id FROM channel_nodes WHERE product_id=? AND parent_id=? AND name=?",
                    (pid, parent_id, chan_path[i]),
                ).fetchone()
                if row:
                    node_id = row["id"]
                else:
                    cur.execute(
                        "INSERT INTO channel_nodes (product_id, parent_id, name, is_leaf) VALUES (?,?,?,1)",
                        (pid, parent_id, chan_path[i]),
                    )
                    node_id = cur.lastrowid
                node_by_path[prefix] = node_id
                parent_id = node_id
    # is_leaf 修正：有孩子的节点不是叶（先收集 ID 再更新，避免游标被重置）
    node_ids = [r["id"] for r in cur.execute("SELECT id FROM channel_nodes")]
    for nid in node_ids:
        cnt = cur.execute(
            "SELECT COUNT(*) FROM channel_nodes WHERE parent_id=?", (nid,)
        ).fetchone()[0]
        cur.execute("UPDATE channel_nodes SET is_leaf=? WHERE id=?", (0 if cnt > 0 else 1, nid))
    conn.commit()


def seed_all(conn, src_path):
    """幂等全量播种：元数据 + 来源/报表模板 + 渠道树（不清库，重复运行安全）"""
    seed_meta(conn)
    seed_templates(conn)
    seed_reports(conn)
    trees, skipped = discover_channel_trees(src_path)
    seed_trees(conn, trees)
    return trees, skipped


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
    init_db()
    conn = get_conn()
    trees, skipped = seed_all(conn, src)
    conn.close()

    print("=== 元数据种子 ===")
    print(f"数据类型 {len(DATA_TYPES)} | 产品线 {len(PRODUCTS)} | 口径 {len(MEASURES)} | 字段 {len(FIELDS)}")
    print("\n=== 渠道树（自动发现自", src, "）===")
    for prod, paths in trees.items():
        print(f"产品 {prod}: {len(paths)} 个块")
        for path, col in sorted(paths.items(), key=lambda x: x[1]):
            name = "KA-" + prod + ("-" + "-".join(path) if path else "")
            mark = "  <- 叶" if not _has_children(path, paths) else "  <- 聚合块(不存储)"
            print(f"  列{col}: {name}{mark}")
    print("\n=== 跳过块（跨产品聚合/未知，不存储）===")
    for s, label in skipped:
        print(f"  列{s}: {label!r}")


def _has_children(path, paths):
    return any(p[: len(path)] == path and len(p) > len(path) for p in paths)


if __name__ == "__main__":
    main()
