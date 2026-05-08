"""
翻译历史记录持久化 — SQLite 数据库

翻译完成后自动保存原文、译文、分析报告、审校报告到本地数据库。
下次打开程序可以浏览历史记录，直接恢复翻译结果进行导出，无需重新翻译。

数据库文件: <程序目录>/data/translation_history.db
"""

import os
import sqlite3
import json
import time
import logging
from datetime import datetime

logger = logging.getLogger('TranslationAgent.history')

# ── 数据库路径 ──
if getattr(__import__('sys'), 'frozen', False):
    _BASE = os.path.dirname(__import__('sys').executable)
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))

_DATA_DIR = os.path.join(_BASE, 'data')
os.makedirs(_DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(_DATA_DIR, 'translation_history.db')


def _get_conn():
    """获取数据库连接，自动建表。"""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    _ensure_tables(conn)
    return conn


def _ensure_tables(conn):
    """如果表不存在则创建。"""
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS translation_records (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        title       TEXT NOT NULL DEFAULT '',
        source_lang TEXT NOT NULL DEFAULT '',
        target_lang TEXT NOT NULL DEFAULT '',
        style       TEXT DEFAULT '',
        audience    TEXT DEFAULT '',
        file_name   TEXT DEFAULT '',
        file_type   TEXT DEFAULT '',
        file_path   TEXT DEFAULT '',
        source_text TEXT DEFAULT '',
        result_text TEXT DEFAULT '',
        analysis    TEXT DEFAULT '',
        critique    TEXT DEFAULT '',
        record_type TEXT NOT NULL DEFAULT 'document',
        -- record_type: 'document' | 'subtitle'
        -- 字幕额外数据
        subtitle_data TEXT DEFAULT NULL,
        created_at   TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
    );

    CREATE INDEX IF NOT EXISTS idx_records_created ON translation_records(created_at DESC);
    CREATE INDEX IF NOT EXISTS idx_records_title  ON translation_records(title);
    """)


# ═══════════════════════════════════════════════════════════
# 写入操作
# ═══════════════════════════════════════════════════════════

def save_record(title, source_lang, target_lang,
                source_text, result_text,
                analysis='', critique='',
                style='', audience='',
                file_name='', file_type='', file_path='',
                record_type='document',
                subtitle_data=None):
    """保存一条翻译记录。

    Parameters
    ----------
    title : str — 记录标题（自动生成或用户指定）
    record_type : str — 'document'（文档翻译）或 'subtitle'（字幕翻译）
    subtitle_data : dict or None — 字幕翻译额外数据（序列化的字幕对象）
    """
    try:
        conn = _get_conn()
        conn.execute("""
            INSERT INTO translation_records
                (title, source_lang, target_lang, style, audience,
                 file_name, file_type, file_path,
                 source_text, result_text, analysis, critique,
                 record_type, subtitle_data)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            title, source_lang, target_lang, style, audience,
            file_name, file_type, file_path,
            source_text, result_text, analysis, critique,
            record_type,
            json.dumps(subtitle_data, ensure_ascii=False) if subtitle_data else None,
        ))
        conn.commit()
        conn.close()
        logger.info(f"翻译记录已保存: {title}")
    except Exception as e:
        logger.error(f"保存翻译记录失败: {e}")


def delete_record(record_id):
    """删除一条翻译记录。"""
    try:
        conn = _get_conn()
        conn.execute("DELETE FROM translation_records WHERE id = ?", (record_id,))
        conn.commit()
        conn.close()
        logger.info(f"翻译记录已删除: id={record_id}")
    except Exception as e:
        logger.error(f"删除翻译记录失败: {e}")


def delete_records_batch(record_ids):
    """批量删除翻译记录。"""
    try:
        conn = _get_conn()
        placeholders = ','.join('?' * len(record_ids))
        conn.execute(f"DELETE FROM translation_records WHERE id IN ({placeholders})", record_ids)
        conn.commit()
        conn.close()
        logger.info(f"批量删除翻译记录: {len(record_ids)} 条")
    except Exception as e:
        logger.error(f"批量删除翻译记录失败: {e}")


def clear_all_records():
    """清空所有翻译记录。"""
    try:
        conn = _get_conn()
        conn.execute("DELETE FROM translation_records")
        conn.execute("VACUUM")
        conn.commit()
        conn.close()
        logger.info("已清空所有翻译记录")
    except Exception as e:
        logger.error(f"清空翻译记录失败: {e}")


# ═══════════════════════════════════════════════════════════
# 查询操作
# ═══════════════════════════════════════════════════════════

def get_all_records(keyword='', limit=100, offset=0):
    """获取所有翻译记录（按时间倒序）。

    Returns list of dict。
    """
    try:
        conn = _get_conn()
        if keyword.strip():
            sql = """
                SELECT id, title, source_lang, target_lang, file_name, file_type,
                       file_path, record_type, created_at,
                       LENGTH(source_text) AS src_len,
                       LENGTH(result_text) AS res_len
                FROM translation_records
                WHERE title LIKE ? OR source_text LIKE ? OR result_text LIKE ?
                ORDER BY created_at DESC
                LIMIT ? OFFSET ?
            """
            kw = f"%{keyword.strip()}%"
            rows = conn.execute(sql, (kw, kw, kw, limit, offset)).fetchall()
        else:
            sql = """
                SELECT id, title, source_lang, target_lang, file_name, file_type,
                       file_path, record_type, created_at,
                       LENGTH(source_text) AS src_len,
                       LENGTH(result_text) AS res_len
                FROM translation_records
                ORDER BY created_at DESC
                LIMIT ? OFFSET ?
            """
            rows = conn.execute(sql, (limit, offset)).fetchall()
        conn.close()
        return [dict(r) for r in rows]
    except Exception as e:
        logger.error(f"查询翻译记录失败: {e}")
        return []


def get_record(record_id):
    """获取一条翻译记录的完整数据（含原文/译文）。

    Returns dict or None。
    """
    try:
        conn = _get_conn()
        row = conn.execute(
            "SELECT * FROM translation_records WHERE id = ?", (record_id,)
        ).fetchone()
        conn.close()
        if row:
            d = dict(row)
            # 反序列化字幕数据
            if d.get('subtitle_data'):
                try:
                    d['subtitle_data'] = json.loads(d['subtitle_data'])
                except Exception:
                    d['subtitle_data'] = None
            return d
        return None
    except Exception as e:
        logger.error(f"获取翻译记录失败: {e}")
        return None


def get_record_count():
    """获取翻译记录总数。"""
    try:
        conn = _get_conn()
        count = conn.execute("SELECT COUNT(*) FROM translation_records").fetchone()[0]
        conn.close()
        return count
    except Exception:
        return 0


def get_db_size_mb():
    """获取数据库文件大小（MB）。"""
    try:
        if os.path.exists(DB_PATH):
            return os.path.getsize(DB_PATH) / (1024 * 1024)
        return 0.0
    except Exception:
        return 0.0


# ═══════════════════════════════════════════════════════════
# 标题自动生成
# ═══════════════════════════════════════════════════════════

def auto_title(source_text, file_name='', source_lang='', target_lang=''):
    """根据原文内容自动生成记录标题。

    优先级：文件名 > 原文前30字 + 语言对 > 语言对 + 时间
    """
    if file_name:
        name = os.path.splitext(file_name)[0]
        return f"{name} ({source_lang}→{target_lang})"

    # 取原文前30字作为标题
    preview = source_text.strip().replace('\n', ' ')[:30]
    if len(source_text.strip()) > 30:
        preview += "..."
    if preview:
        return f"{preview} ({source_lang}→{target_lang})"

    return f"{source_lang}→{target_lang} {datetime.now().strftime('%m/%d %H:%M')}"
