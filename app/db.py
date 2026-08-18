# -*- coding: utf-8 -*-
"""数据库连接与初始化（SQLite 单文件）"""
import pathlib
import sqlite3

BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "jingying.db"
SCHEMA_PATH = pathlib.Path(__file__).resolve().parent / "schema.sql"


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    """建库：目录 + 表结构（幂等，不会删数据）"""
    DATA_DIR.mkdir(exist_ok=True)
    conn = get_conn()
    conn.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))
    conn.commit()
    conn.close()
