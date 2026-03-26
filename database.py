import sqlite3
import os

DB_PATH = os.environ.get('DB_PATH', 'qc_data.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript('''
        CREATE TABLE IF NOT EXISTS qc_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            machine_id TEXT NOT NULL,
            shift TEXT NOT NULL,
            part_no TEXT,
            lot_no TEXT,
            submitted_by TEXT,
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            UNIQUE(date, machine_id, shift)
        );
        CREATE TABLE IF NOT EXISTS visual_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id INTEGER REFERENCES qc_entries(id),
            item_index INTEGER,
            item_name TEXT,
            result TEXT,
            rejected_lot TEXT
        );
        CREATE TABLE IF NOT EXISTS eol_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id INTEGER REFERENCES qc_entries(id),
            item_index INTEGER,
            item_name TEXT,
            result TEXT,
            rejected_lot TEXT
        );
        CREATE TABLE IF NOT EXISTS dim_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id INTEGER REFERENCES qc_entries(id),
            item_index INTEGER,
            item_name TEXT,
            result TEXT
        );
        CREATE TABLE IF NOT EXISTS machine_handover (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            machine_id TEXT NOT NULL,
            last_batch TEXT,
            reason TEXT,
            updated_at TEXT DEFAULT (datetime('now','localtime')),
            UNIQUE(date, machine_id)
        );
        CREATE TABLE IF NOT EXISTS abnormality (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            machine_id TEXT NOT NULL,
            shift TEXT,
            description TEXT,
            cause TEXT,
            countermeasure TEXT,
            status TEXT DEFAULT 'open',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
    ''')
    conn.commit()
    conn.close()

def save_entry(date, machine_id, shift, part_no, lot_no, submitted_by, notes,
               visual_items, eol_items, dim_items):
    conn = get_db()
    try:
        # Upsert entry
        conn.execute('''
            INSERT INTO qc_entries (date, machine_id, shift, part_no, lot_no, submitted_by, notes)
            VALUES (?,?,?,?,?,?,?)
            ON CONFLICT(date, machine_id, shift) DO UPDATE SET
                part_no=excluded.part_no, lot_no=excluded.lot_no,
                submitted_by=excluded.submitted_by, notes=excluded.notes,
                created_at=datetime('now','localtime')
        ''', (date, machine_id, shift, part_no, lot_no, submitted_by, notes))
        conn.commit()

        entry = conn.execute(
            'SELECT id FROM qc_entries WHERE date=? AND machine_id=? AND shift=?',
            (date, machine_id, shift)
        ).fetchone()
        entry_id = entry['id']

        for tbl in ['visual_results','eol_results','dim_results']:
            conn.execute(f'DELETE FROM {tbl} WHERE entry_id=?', (entry_id,))

        for i, item in enumerate(visual_items):
            conn.execute('INSERT INTO visual_results(entry_id,item_index,item_name,result,rejected_lot) VALUES(?,?,?,?,?)',
                        (entry_id, i, item['name'], item.get('result',''), item.get('rejected_lot','')))
        for i, item in enumerate(eol_items):
            conn.execute('INSERT INTO eol_results(entry_id,item_index,item_name,result,rejected_lot) VALUES(?,?,?,?,?)',
                        (entry_id, i, item['name'], item.get('result',''), item.get('rejected_lot','')))
        for i, item in enumerate(dim_items):
            conn.execute('INSERT INTO dim_results(entry_id,item_index,item_name,result) VALUES(?,?,?,?)',
                        (entry_id, i, item['name'], item.get('result','')))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def get_entry(date, machine_id, shift):
    conn = get_db()
    entry = conn.execute(
        'SELECT * FROM qc_entries WHERE date=? AND machine_id=? AND shift=?',
        (date, machine_id, shift)
    ).fetchone()
    if not entry:
        conn.close()
        return None
    entry_id = entry['id']
    visual = conn.execute('SELECT * FROM visual_results WHERE entry_id=? ORDER BY item_index', (entry_id,)).fetchall()
    eol = conn.execute('SELECT * FROM eol_results WHERE entry_id=? ORDER BY item_index', (entry_id,)).fetchall()
    dims = conn.execute('SELECT * FROM dim_results WHERE entry_id=? ORDER BY item_index', (entry_id,)).fetchall()
    conn.close()
    return {'entry': dict(entry), 'visual': [dict(r) for r in visual],
            'eol': [dict(r) for r in eol], 'dims': [dict(r) for r in dims]}

def get_daily_status(date):
    conn = get_db()
    entries = conn.execute('SELECT machine_id, shift FROM qc_entries WHERE date=?', (date,)).fetchall()
    handovers = conn.execute('SELECT * FROM machine_handover WHERE date=?', (date,)).fetchall()
    abnormalities = conn.execute('SELECT * FROM abnormality WHERE date=? ORDER BY created_at DESC', (date,)).fetchall()
    conn.close()
    return {
        'entries': [dict(e) for e in entries],
        'handovers': [dict(h) for h in handovers],
        'abnormalities': [dict(a) for a in abnormalities]
    }

def save_handover(date, machine_id, last_batch, reason):
    conn = get_db()
    conn.execute('''
        INSERT INTO machine_handover(date, machine_id, last_batch, reason)
        VALUES(?,?,?,?)
        ON CONFLICT(date, machine_id) DO UPDATE SET
            last_batch=excluded.last_batch, reason=excluded.reason,
            updated_at=datetime('now','localtime')
    ''', (date, machine_id, last_batch, reason))
    conn.commit()
    conn.close()

def save_abnormality(date, machine_id, shift, description, cause, countermeasure):
    conn = get_db()
    conn.execute('''
        INSERT INTO abnormality(date, machine_id, shift, description, cause, countermeasure)
        VALUES(?,?,?,?,?,?)
    ''', (date, machine_id, shift, description, cause, countermeasure))
    conn.commit()
    conn.close()

def get_history_dates(limit=30):
    conn = get_db()
    rows = conn.execute(
        'SELECT DISTINCT date FROM qc_entries ORDER BY date DESC LIMIT ?', (limit,)
    ).fetchall()
    conn.close()
    return [r['date'] for r in rows]

def get_all_entries_for_date(date):
    conn = get_db()
    entries = conn.execute('SELECT * FROM qc_entries WHERE date=? ORDER BY machine_id, shift', (date,)).fetchall()
    result = []
    for entry in entries:
        entry_id = entry['id']
        visual = conn.execute('SELECT * FROM visual_results WHERE entry_id=? ORDER BY item_index', (entry_id,)).fetchall()
        eol = conn.execute('SELECT * FROM eol_results WHERE entry_id=? ORDER BY item_index', (entry_id,)).fetchall()
        dims = conn.execute('SELECT * FROM dim_results WHERE entry_id=? ORDER BY item_index', (entry_id,)).fetchall()
        result.append({
            'entry': dict(entry),
            'visual': [dict(r) for r in visual],
            'eol': [dict(r) for r in eol],
            'dims': [dict(r) for r in dims]
        })
    handovers = conn.execute('SELECT * FROM machine_handover WHERE date=? ORDER BY machine_id', (date,)).fetchall()
    abnormalities = conn.execute('SELECT * FROM abnormality WHERE date=? ORDER BY machine_id', (date,)).fetchall()
    conn.close()
    return {
        'entries': result,
        'handovers': [dict(h) for h in handovers],
        'abnormalities': [dict(a) for a in abnormalities]
    }
