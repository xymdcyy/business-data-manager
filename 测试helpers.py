# -*- coding: utf-8 -*-
"""测试公共工具：临时库隔离（防清空生产数据）+ chk 断言 + 结果汇总"""
import pathlib
import shutil
import tempfile

import app.db as dbmod

_ok = _fail = 0


def use_test_db():
    """生产库复制到临时目录，测试只操作副本"""
    src = dbmod.DB_PATH
    if not src.exists():
        raise SystemExit(f"生产库不存在: {src}，请先运行: uv run python 迁移数据.py --yes")
    dst = pathlib.Path(tempfile.gettempdir()) / "jingying-测试库.db"
    shutil.copy2(src, dst)
    dbmod.DB_PATH = dst
    print(f"[测试用临时库] 副本: {dst}（生产库不受影响）")
    from app.db import get_conn
    return get_conn()


def chk(label, actual, expect, tol=0.02):
    """断言：expect=None 期望无数据；str 精确比对（表头/标签）；否则数值比对（±tol）"""
    global _ok, _fail
    if expect is None:
        good = actual is None
    elif isinstance(expect, str):
        good = actual == expect
    else:
        good = actual is not None and abs(actual - expect) <= tol

    def _show(v):
        return v if v is None or isinstance(v, str) else round(v, 4)

    if good:
        _ok += 1
        print(f"  ✓ {label}: {_show(actual)}")
    else:
        _fail += 1
        print(f"  ✗ {label}: {_show(actual)} vs {_show(expect)}")


def result(label="测试"):
    """打印汇总；有失败则退出码 1"""
    print(f"\n结果: {_ok} 通过, {_fail} 失败")
    if _fail:
        raise SystemExit(1)
    print(f"{label} OK")
