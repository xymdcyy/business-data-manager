-- 经营数据管理系统 schema
-- 原则：只存叶子数据（产品线×渠道×指标×数据类型×年月），所有合计/派生沿树与公式计算

-- 数据类型：实际 / BP / 预测 / 同期（同期为虚类型：base=实际 + year_shift=-1，不存储）
CREATE TABLE IF NOT EXISTS data_types (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    base_dtype_id INTEGER REFERENCES data_types(id),  -- 虚类型的实际存储类型（同期->实际）
    year_shift INTEGER NOT NULL DEFAULT 0,
    note TEXT
);

-- 产品线：智屏 / 空调 / 冰洗 / CIoT
CREATE TABLE IF NOT EXISTS product_lines (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    note TEXT
);

-- 渠道树节点（按产品线各一棵；产品节点为根，parent_id 为空）
CREATE TABLE IF NOT EXISTS channel_nodes (
    id INTEGER PRIMARY KEY,
    product_id INTEGER NOT NULL REFERENCES product_lines(id),
    parent_id INTEGER REFERENCES channel_nodes(id),
    name TEXT NOT NULL,
    is_leaf INTEGER NOT NULL DEFAULT 1,
    UNIQUE (product_id, parent_id, name)
);

-- 指标字段：raw=导入原始值；formula=由其他字段计算（同一数据类型上下文）
CREATE TABLE IF NOT EXISTS fields (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    kind TEXT NOT NULL DEFAULT 'raw' CHECK (kind IN ('raw', 'formula')),
    formula TEXT,          -- kind='formula' 时的表达式；裸字段名=同数据类型，可显式 {数据类型} 固定
    unit TEXT,             -- 万元 / 万台 / 元 / %
    scale INTEGER DEFAULT 2,  -- 导出数字格式小数位
    note TEXT
);

-- 口径（报表列）：dtype=直接取字段的某数据类型（year_shift 支持同期=-1年）；formula=跨数据类型公式（变量=数据类型名）
CREATE TABLE IF NOT EXISTS measures (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    kind TEXT NOT NULL CHECK (kind IN ('dtype', 'formula')),
    dtype_id INTEGER REFERENCES data_types(id),  -- kind='dtype'
    year_shift INTEGER NOT NULL DEFAULT 0,       -- kind='dtype' 时生效（同期=-1）
    formula TEXT                                 -- kind='formula'，如 实际/BP、实际/同期-1
);

-- 事实表：只存叶子节点数据
CREATE TABLE IF NOT EXISTS facts (
    id INTEGER PRIMARY KEY,
    year INTEGER NOT NULL,
    month INTEGER NOT NULL,           -- 1..12
    node_id INTEGER NOT NULL REFERENCES channel_nodes(id),
    field_id INTEGER NOT NULL REFERENCES fields(id),
    dtype_id INTEGER NOT NULL REFERENCES data_types(id),
    value REAL NOT NULL,
    source TEXT,                      -- 导入来源描述 / 手工
    updated_at TEXT DEFAULT (datetime('now', 'localtime')),
    UNIQUE (year, month, node_id, field_id, dtype_id)
);
CREATE INDEX IF NOT EXISTS idx_facts_lookup ON facts(year, month, node_id, field_id, dtype_id);

-- 导入来源模板：每个来源文件类型一份配置
CREATE TABLE IF NOT EXISTS source_templates (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    path_pattern TEXT,      -- 文件路径（含月份/年份变量）
    sheet_name TEXT,
    config_json TEXT,       -- 块识别规则 + 行→字段映射 + 标签→节点规则
    enabled INTEGER NOT NULL DEFAULT 1,
    note TEXT
);

-- 报表模板
CREATE TABLE IF NOT EXISTS report_templates (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    config_json TEXT,       -- 行/列/期间/样式模式
    style_json TEXT,        -- 样式快照或规格
    enabled INTEGER NOT NULL DEFAULT 1,
    note TEXT
);
