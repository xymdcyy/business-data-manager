# -*- coding: utf-8 -*-
"""计算引擎：公式求值 + 渠道树聚合 + 任意组合行取值

- 公式语言：四则运算，裸字段名 = 上下文数据类型，可显式 字段名{数据类型} 固定数据类型
  （例：ASP = 收入 / 销量；BP达成率 = 实际 / BP）
- 聚合：沿渠道树对叶子求和（缺数据叶子按 0 计；整组无数据 = None——顶层 raw 取值语义）
- 比率字段（formula 型）：聚合时先对各成分字段求和，再套公式（ASP 加权正确）
- 公式求值内部：成分字段无数据按 0 参与加减乘（SUM/差值链语义，源表 0 被 skip_zero 跳过）；
  除法除零/缺分母 → None（达成率/同比无数据显示）。公式环显式 ValueError 报错。
"""
import ast
import functools
import sqlite3

# ---------- 公式语法 ----------
import re as _re

_BRACE_PAT = _re.compile(r"([一-鿿A-Za-z0-9_]+)\{([一-鿿A-Za-z0-9_]+)\}")
_ALLOWED_NODES = (
    ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Name, ast.Load,
    ast.Add, ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd,
)


def preprocess_formula(expr: str) -> str:
    """字段名{数据类型} -> 字段名__数据类型（合法标识符）；裸名保留"""
    return _BRACE_PAT.sub(lambda m: f"{m.group(1)}__{m.group(2)}", expr)


@functools.lru_cache(maxsize=512)
def compile_expr(expr: str) -> ast.Expression:
    """编译公式为 AST（缓存：同一公式每单元格重复 ast.parse 浪费；AST 只读不改，可安全共享）"""
    tree = ast.parse(preprocess_formula(expr), mode="eval")
    for node in ast.walk(tree):
        if not isinstance(node, _ALLOWED_NODES):
            raise ValueError(f"公式语法不支持: {type(node).__name__}，公式: {expr}")
    return tree


def eval_tree(tree: ast.Expression, resolver) -> float:
    """resolver(name) -> float；求值失败（引用缺失/除零）返回 None"""
    def _eval(n):
        if isinstance(n, ast.BinOp):
            l, r = _eval(n.left), _eval(n.right)
            if l is None or r is None:
                return None
            try:
                if isinstance(n.op, ast.Add):
                    return l + r
                if isinstance(n.op, ast.Sub):
                    return l - r
                if isinstance(n.op, ast.Mult):
                    return l * r
                return l / r
            except ZeroDivisionError:
                return None
        if isinstance(n, ast.UnaryOp):
            v = _eval(n.operand)
            return -v if isinstance(n.op, ast.USub) else v
        if isinstance(n, ast.Constant):
            return n.value
        if isinstance(n, ast.Name):
            try:
                return resolver(n.id)
            except KeyError:
                return None
        raise ValueError(f"不支持的语法节点: {type(n).__name__}")
    return _eval(tree.body if isinstance(tree, ast.Expression) else tree)


# ---------- 树聚合 ----------
def subtree_leaves(conn: sqlite3.Connection, node_id: int) -> list:
    """节点子树的所有叶子（含自身若是叶）"""
    rows = conn.execute(
        """
        WITH RECURSIVE sub(id) AS (
            SELECT ? UNION ALL
            SELECT cn.id FROM channel_nodes cn JOIN sub ON cn.parent_id = sub.id
        )
        SELECT cn.id FROM sub JOIN channel_nodes cn ON cn.id = sub.id WHERE cn.is_leaf = 1
        """,
        (node_id,),
    ).fetchall()
    return [r["id"] for r in rows]


def _expand(conn, node_ids):
    """节点集展开为叶子集（聚合节点 -> 其子树全部叶子）"""
    leaves = set()
    for nid in node_ids:
        leaves.update(subtree_leaves(conn, nid))
    return list(leaves)


def _sum_facts_range(conn, year, m1, m2, node_ids, field_id, dtype_id, extra_shift=0):
    """月份区间版（m1..m2 闭区间）：缺数据叶子按 0 计；整组无数据 = None
    虚类型解析：base_dtype_id 决定实际查库类型（同期->实际），year_shift 决定年份偏移（-1）"""
    drow = conn.execute(
        "SELECT base_dtype_id, year_shift FROM data_types WHERE id=?", (dtype_id,)
    ).fetchone()
    base = drow["base_dtype_id"] if drow and drow["base_dtype_id"] else dtype_id
    shift = extra_shift + (drow["year_shift"] if drow else 0)
    y = year + shift
    leaf_ids = _expand(conn, node_ids)
    if not leaf_ids:
        return None
    ph = ",".join("?" * len(leaf_ids))
    rows = conn.execute(
        f"SELECT value FROM facts WHERE year=? AND month BETWEEN ? AND ? AND node_id IN ({ph}) "
        "AND field_id=? AND dtype_id=?",
        (y, m1, m2, *leaf_ids, field_id, base),
    ).fetchall()
    if not rows:
        return None
    return sum(r["value"] for r in rows)


def _dtype_id_map(conn):
    """数据类型名 -> id（公式解析用，集中定义避免重复构建）"""
    return {r["name"]: r["id"] for r in conn.execute("SELECT id, name FROM data_types")}


def node_path(conn, node_id):
    """节点完整路径（根到自身）；返回 [产品根, ...渠道]"""
    names = []
    nid = node_id
    guard = 0
    while nid and guard < 20:
        row = conn.execute("SELECT name, parent_id FROM channel_nodes WHERE id=?", (nid,)).fetchone()
        if not row:
            break
        names.append(row["name"])
        nid = row["parent_id"]
        guard += 1
    return list(reversed(names))


class FieldResolverRange:
    """字段公式求值解析器（区间版）：NAME__DTYPE -> 字段@数据类型区间聚合；裸 NAME -> 上下文数据类型
    extra_shift 透传给成分字段（同期口径的年份偏移），保证公式字段的同期计算正确
    formula 字段递归求值（固定费用→费用总额→端到端利润 链式依赖）；深度保护防公式环"""

    def __init__(self, conn, year, m1, m2, node_ids, ctx_dtype_id, extra_shift=0):
        self.conn = conn
        self.year, self.m1, self.m2 = year, m1, m2
        self.node_ids = node_ids
        self.ctx_dtype = ctx_dtype_id
        self.extra_shift = extra_shift
        self._field_row = {r["name"]: r for r in conn.execute(
            "SELECT id, name, kind, formula FROM fields")}
        self._field_id = {r["name"]: r["id"] for r in self._field_row.values()}
        self._dtype_id = _dtype_id_map(conn)
        self._depth = 0

    def __call__(self, name: str) -> float:
        field_name, dtype_id = name, self.ctx_dtype
        if "__" in name:
            field_name, dtype_name = name.rsplit("__", 1)
            dtype_id = self._dtype_id[dtype_name]
        if field_name not in self._field_id:
            raise KeyError(name)
        return self._resolve(field_name, dtype_id)

    def _resolve(self, field_name, dtype_id):
        row = self._field_row[field_name]
        if row["kind"] == "raw":
            val = _sum_facts_range(self.conn, self.year, self.m1, self.m2, self.node_ids,
                                   row["id"], dtype_id, self.extra_shift)
            # 无数据按 0 参与公式运算（SUM/差值链语义，源表 0 值被 skip_zero 跳过）
            return val if val is not None else 0.0
        if self._depth > 10:
            raise ValueError(f"字段公式环: {field_name}")  # 显式报错，不被 eval_tree 吞掉
        self._depth += 1
        try:
            return eval_tree(compile_expr(row["formula"]), self)
        finally:
            self._depth -= 1


def field_value(conn, year, month, node_ids, field_id, dtype_id, extra_shift=0):
    """单月取值：field 为 raw -> 叶子求和；formula -> 按公式重算（成分先聚合再运算）"""
    return aggregate_range(conn, year, month, month, node_ids, field_id, dtype_id, extra_shift)


class FormulaBulkEvaluator:
    """批量公式求值器（明细「包含计算值」用）：一次读取字段/类型元数据，按 (节点,类型) 缓存
    事实行与叶子展开，避免逐 (组合×公式字段) 重建 Resolver + 重复 SQL。
    单节点取值语义与 field_value(conn, year, month, [leaf], ...) 一致（raw 缺 -> None，成分缺按 0）"""

    def __init__(self, conn):
        self.conn = conn
        self._field_row = {r["name"]: r for r in conn.execute(
            "SELECT id, name, kind, formula FROM fields")}
        self._id_to_name = {r["id"]: r["name"] for r in self._field_row.values()}
        self._field_id = {r["name"]: r["id"] for r in self._field_row.values()}
        self._dtype_id = _dtype_id_map(conn)
        self._leaves = {}
        self._facts = {}  # (node_id, dtype_id) -> {(year, month): {field_id: value}}

    def _month_facts(self, year, month, node_id, dtype_id):
        """节点当月事实行（叶子集展开缓存；单节点 = 该叶子自身）"""
        key = (node_id, dtype_id)
        bucket = self._facts.setdefault(key, {})
        if (year, month) not in bucket:
            leaf_ids = self._leaves.get(node_id)
            if leaf_ids is None:
                leaf_ids = self._leaves[node_id] = subtree_leaves(self.conn, node_id)
            month_rows = {}
            if leaf_ids:
                ph = ",".join("?" * len(leaf_ids))
                for r in self.conn.execute(
                        f"SELECT month, field_id, value FROM facts "
                        f"WHERE year=? AND node_id IN ({ph}) AND dtype_id=?",
                        (year, *leaf_ids, dtype_id)):
                    month_rows.setdefault(r["month"], {})[r["field_id"]] = r["value"]
            bucket[(year, month)] = month_rows
        return bucket[(year, month)]

    def value(self, year, month, node_id, field_id, dtype_id):
        """同 field_value(conn, year, month, [node_id], field_id, dtype_id) 语义"""
        name = self._id_to_name[field_id]
        return self._value_name(year, month, node_id, name, dtype_id, 0)

    def _value_name(self, year, month, node_id, field_name, dtype_id, depth):
        row = self._field_row[field_name]
        if row["kind"] == "raw":
            return self._month_facts(year, month, node_id, dtype_id).get(month, {}).get(row["id"])
        if depth > 10:
            raise ValueError(f"字段公式环: {field_name}")
        if not row["formula"]:
            return None

        def resolver(name):
            fname, d_id = name, dtype_id
            if "__" in name:
                fname, dname = name.rsplit("__", 1)
                d_id = self._dtype_id[dname]
            if fname not in self._field_id:
                raise KeyError(name)
            val = self._value_name(year, month, node_id, fname, d_id, depth + 1)
            return val if val is not None else 0.0  # 成分缺按 0（同 FieldResolverRange）

        return eval_tree(compile_expr(row["formula"]), resolver)


def aggregate_range(conn, year, m1, m2, node_ids, field_id, dtype_id, extra_shift=0):
    """区间取值（m1..m2）：raw -> 叶子区间求和；formula -> 公式重算（成分按区间先聚合）"""
    row = conn.execute("SELECT kind, formula FROM fields WHERE id=?", (field_id,)).fetchone()
    if row["kind"] == "raw":
        return _sum_facts_range(conn, year, m1, m2, node_ids, field_id, dtype_id, extra_shift)
    # formula 字段：求值于同一组叶子与区间（比率指标跨月加权正确）；extra_shift 透传（同期）
    resolver = FieldResolverRange(conn, year, m1, m2, node_ids, dtype_id, extra_shift)
    return eval_tree(compile_expr(row["formula"]), resolver)


class _PeriodResolver:
    """复合期间字段公式解析器：成分字段跨全部段聚合后再套公式
    （比率正确性：1-7实际+8-12预测 的 ASP = Σ收入/Σ销量，不是各段 ASP 之和）"""

    def __init__(self, conn, year, segments, node_ids, ctx_dtype_id, extra_shift, mixed):
        self.conn, self.year, self.segments = conn, year, segments
        self.node_ids, self.ctx_dtype = node_ids, ctx_dtype_id
        self.extra_shift, self.mixed = extra_shift, mixed
        self._field_row = {r["name"]: r for r in conn.execute(
            "SELECT id, name, kind, formula FROM fields")}
        self._field_id = {r["name"]: r["id"] for r in self._field_row.values()}
        self._dtype_id = _dtype_id_map(conn)
        self._depth = 0

    def __call__(self, name: str) -> float:
        field_name, dtype_id = name, self.ctx_dtype
        if "__" in name:
            field_name, dtype_name = name.rsplit("__", 1)
            dtype_id = self._dtype_id[dtype_name]
        if field_name not in self._field_id:
            raise KeyError(name)
        return self._resolve(field_name, dtype_id)

    def _resolve(self, field_name, dtype_id):
        row = self._field_row[field_name]
        if row["kind"] == "raw":
            total = 0.0
            for m1, m2, dname in self.segments:
                if self.mixed:
                    seg = self._dtype_id.get(dname)
                    if seg is None:
                        continue
                    v = _sum_facts_range(self.conn, self.year, m1, m2, self.node_ids,
                                         row["id"], seg)
                else:
                    v = _sum_facts_range(self.conn, self.year, m1, m2, self.node_ids,
                                         row["id"], dtype_id, self.extra_shift)
                if v is not None:
                    total += v
            # 成分字段缺数据按 0 参与公式运算（SUM/差值链语义，源表 0 被 skip_zero 跳过）；
            # 除零/缺分母由 eval_tree 处理为 None（比率类天然显示空）
            return total
        if self._depth > 10:
            raise ValueError(f"字段公式环: {field_name}")  # 显式报错，不被 eval_tree 吞掉
        self._depth += 1
        try:
            return eval_tree(compile_expr(row["formula"]), self)
        finally:
            self._depth -= 1


def period_aggregate(conn, year, segments, node_ids, field_id, dtype_id, extra_shift=0, mixed=False):
    """复合期间取值。segments=[(m1,m2,dtype名)]
    变量语义：mixed=True（口径=实际）→ 各段按段自身 dtype 取（1-7实际+8-12预测）；
    mixed=False → 各段按 dtype_id 取（BP/预测/同期：同期经 base+year_shift 解析为去年实际）
    raw 字段：跨段求和；formula 字段：成分字段跨段聚合后再套公式"""
    row = conn.execute("SELECT kind, formula FROM fields WHERE id=?", (field_id,)).fetchone()
    if row["kind"] == "raw":
        vals = []
        for m1, m2, dname in segments:
            if mixed:
                seg = conn.execute("SELECT id FROM data_types WHERE name=?", (dname,)).fetchone()
                if seg is None:
                    continue
                v = _sum_facts_range(conn, year, m1, m2, node_ids, field_id, seg["id"])
            else:
                v = _sum_facts_range(conn, year, m1, m2, node_ids, field_id, dtype_id, extra_shift)
            if v is not None:
                vals.append(v)
        return sum(vals) if vals else None
    resolver = _PeriodResolver(conn, year, segments, node_ids, dtype_id, extra_shift, mixed)
    return eval_tree(compile_expr(row["formula"]), resolver)


def period_measure(conn, year, segments, node_ids, field_id, measure_row):
    """复合期间 × 口径：dtype -> period_aggregate；formula -> 各数据类型变量取期间聚合后套公式"""
    if measure_row["kind"] == "dtype":
        mixed = measure_row["name"] == "实际"
        return period_aggregate(conn, year, segments, node_ids, field_id,
                                measure_row["dtype_id"], measure_row["year_shift"], mixed)
    dtype_id = _dtype_id_map(conn)

    def resolver(name):
        if name not in dtype_id:
            raise KeyError(name)
        val = period_aggregate(conn, year, segments, node_ids, field_id,
                               dtype_id[name], 0, mixed=name == "实际")
        if val is None:
            raise KeyError(name)
        return val

    return eval_tree(compile_expr(measure_row["formula"]), resolver)


def measure_value(conn, year, month, node_ids, field_id, measure_row):
    """口径取值：dtype -> 字段@该数据类型（含 year_shift）；formula -> 跨数据类型公式（变量=数据类型名）"""
    if measure_row["kind"] == "dtype":
        return field_value(conn, year, month, node_ids, field_id,
                           measure_row["dtype_id"], measure_row["year_shift"])
    # formula 口径：变量 = 数据类型名，作用于该列字段
    dtype_id = _dtype_id_map(conn)

    def resolver(name):
        if name not in dtype_id:
            raise KeyError(name)
        val = field_value(conn, year, month, node_ids, field_id, dtype_id[name])
        if val is None:
            raise KeyError(name)
        return val

    return eval_tree(compile_expr(measure_row["formula"]), resolver)


# ---------- 计算过程求值树（报表单元格钻取） ----------
def explain_period_measure(conn, year, segments, node_ids, field_id, measure_row, mixed=False):
    """单元格求值树（用于前端展示计算过程）：
    {name, formula?, value, raw?, children:[...]}；公式节点展开成分，raw 节点为叶子。
    语义与 period_measure 一致（None 传播、缺失成分按 0 参与加减乘、除零→None）。"""
    dtype_id_map = _dtype_id_map(conn)
    fields = {r["name"]: r for r in conn.execute("SELECT id, name, kind, formula FROM fields")}
    id_to_name = {r["id"]: r["name"] for r in fields.values()}

    def _field_sum(field_name, dtype_id, extra_shift=0, mixed_here=None):
        """成分字段在复合期间下的合计（同 _PeriodResolver 语义）"""
        row = fields[field_name]
        if row["kind"] == "raw":
            use_mixed = mixed if mixed_here is None else mixed_here
            total = None
            for m1, m2, dname in segments:
                if use_mixed:
                    seg_id = dtype_id_map.get(dname)
                    if seg_id is None:
                        continue
                    v = _sum_facts_range(conn, year, m1, m2, node_ids, row["id"], seg_id)
                else:
                    v = _sum_facts_range(conn, year, m1, m2, node_ids, row["id"], dtype_id, extra_shift)
                if v is not None:
                    total = (total or 0.0) + v
            return total
        return _eval_formula(row["formula"], dtype_id, extra_shift, 0, mixed_here)

    def _eval_formula(expr, dtype_id, extra_shift=0, depth=0, mixed_here=None):
        """公式求值并记录过程；返回 {name?, formula, value, children}"""
        if depth > 10:
            raise ValueError(f"字段公式环: {expr}")
        tree = compile_expr(expr)
        children = []

        def resolver(name):
            fname, f_dtype = name, dtype_id
            if "__" in name:
                fname, dname = name.rsplit("__", 1)
                f_dtype = dtype_id_map[dname]
            child = _explain_one(fname, f_dtype, extra_shift, depth + 1, mixed_here)
            children.append(child)
            return child["value"] if child["value"] is not None else (0.0 if fields[fname]["kind"] == "raw" else None)

        value = eval_tree(tree, resolver)
        return {"formula": expr, "value": value, "children": children}

    def _explain_one(field_name, dtype_id, extra_shift=0, depth=0, mixed_here=None):
        row = fields[field_name]
        if row["kind"] == "raw":
            return {"name": field_name, "value": _field_sum(field_name, dtype_id, extra_shift, mixed_here),
                    "raw": True}
        node = _eval_formula(row["formula"], dtype_id, extra_shift, depth, mixed_here)
        node["name"] = field_name
        return node

    if measure_row["kind"] == "dtype":
        # 虚类型（同期）经 base_dtype_id 解析为实际类型；year_shift 叠加（与 _sum_facts_range 同语义）
        drow = conn.execute("SELECT base_dtype_id, year_shift FROM data_types WHERE id=?",
                            (measure_row["dtype_id"],)).fetchone()
        base = drow["base_dtype_id"] if drow and drow["base_dtype_id"] else measure_row["dtype_id"]
        shift = (drow["year_shift"] if drow else 0) + measure_row["year_shift"]
        return _explain_one(id_to_name[field_id], base, shift)
    # formula 口径：变量 = 数据类型名（各变量 mixed = 名称为实际）
    dtypes = {r["name"]: r["id"] for r in conn.execute("SELECT id, name FROM data_types")}
    children = []

    def resolver(name):
        if name not in dtypes:
            raise KeyError(name)
        node = _explain_one(id_to_name[field_id], dtypes[name], 0, 1, mixed_here=(name == "实际"))
        node["name"] = name
        children.append(node)
        return node["value"] if node["value"] is not None else None

    value = eval_tree(compile_expr(measure_row["formula"]), resolver)
    return {"name": measure_row["name"], "formula": measure_row["formula"], "value": value,
            "children": children}


def resolve_path(conn, product_name, path: tuple):
    """产品名+渠道路径 -> 节点 id；找不到返回 None"""
    row = conn.execute(
        """SELECT cn.id FROM channel_nodes cn JOIN product_lines p ON p.id = cn.product_id
           WHERE p.name = ? AND cn.name = ? AND cn.parent_id IS NULL""",
        (product_name, product_name),
    ).fetchone()
    if row is None:
        return None
    node_id = row["id"]
    for seg in path:
        r2 = conn.execute(
            "SELECT id FROM channel_nodes WHERE parent_id=? AND name=?",
            (node_id, seg),
        ).fetchone()
        if r2 is None:
            return None
        node_id = r2["id"]
    return node_id


def selector_sum(conn, year, month, node_id_list, field_id, dtype_id):
    """任意组合行：多个节点取值后相加（None 跳过；全 None 为 None）"""
    vals = [field_value(conn, year, month, [nid], field_id, dtype_id) for nid in node_id_list]
    vals = [v for v in vals if v is not None]
    return sum(vals) if vals else None
