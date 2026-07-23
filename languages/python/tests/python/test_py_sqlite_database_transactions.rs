use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: SQLite Database Transactions — sqlite3.connect, Row, executemany, commit, rollback, create_function, iterdump
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_sqlite_in_memory_crud_operations() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()

cur.execute("CREATE TABLE users (id INT PRIMARY KEY, name TEXT)")
cur.execute("INSERT INTO users VALUES (1, 'Alice')")
cur.execute("INSERT INTO users VALUES (2, 'Bob')")
conn.commit()

cur.execute("SELECT name FROM users ORDER BY id")
print([r[0] for r in cur.fetchall()])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["['Alice', 'Bob']"]);
}

#[test]
fn test_py_sqlite_executemany_parameter_binding() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE scores (name TEXT, score INT)")

items = [("Alice", 90), ("Bob", 85), ("Charlie", 95)]
cur.executemany("INSERT INTO scores VALUES (?, ?)", items)
conn.commit()

cur.execute("SELECT SUM(score) FROM scores")
print(cur.fetchone()[0])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["270"]);
}

#[test]
fn test_py_sqlite_row_factory_dict_indexing() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
conn.row_factory = sqlite3.Row
cur = conn.cursor()

cur.execute("CREATE TABLE config (k TEXT, v TEXT)")
cur.execute("INSERT INTO config VALUES ('theme', 'dark')")

cur.execute("SELECT * FROM config")
row = cur.fetchone()
print(row["k"], row["v"])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["theme dark"]);
}

#[test]
fn test_py_sqlite_transaction_context_manager_rollback() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE items (val INT)")
cur.execute("INSERT INTO items VALUES (10)")
conn.commit()

try:
    with conn:
        cur.execute("INSERT INTO items VALUES (20)")
        raise RuntimeError("Transaction fail")
except RuntimeError:
    pass

cur.execute("SELECT COUNT(*) FROM items")
print(cur.fetchone()[0])  # rolled back to 1
conn.close()
"#;
    assert_eq!(run_python(src), vec!["1"]);
}

#[test]
fn test_py_sqlite_custom_python_function_in_sql() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
conn.create_function("double_val", 1, lambda x: x * 2)

cur = conn.cursor()
cur.execute("SELECT double_val(21)")
print(cur.fetchone()[0])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["42"]);
}

#[test]
fn test_py_sqlite_iterdump_sql_dump() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE test (id INT)")
cur.execute("INSERT INTO test VALUES (1)")

dump = "\n".join(conn.iterdump())
print("CREATE TABLE test" in dump)
print("INSERT INTO \"test\" VALUES(1)" in dump)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sqlite_lastrowid_auto_increment() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE logs (id INTEGER PRIMARY KEY AUTOINCREMENT, msg TEXT)")

cur.execute("INSERT INTO logs (msg) VALUES ('m1')")
print(cur.lastrowid)

cur.execute("INSERT INTO logs (msg) VALUES ('m2')")
print(cur.lastrowid)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["1", "2"]);
}

#[test]
fn test_py_sqlite_pragma_table_info() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE widget (id INT, name TEXT)")

cur.execute("PRAGMA table_info(widget)")
cols = [r[1] for r in cur.fetchall()]
print(cols)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["['id', 'name']"]);
}

#[test]
fn test_py_sqlite_blob_data_roundtrip() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE files (data BLOB)")

raw = b"\x00\xff\xfe\xfd"
cur.execute("INSERT INTO files VALUES (?)", (sqlite3.Binary(raw),))

cur.execute("SELECT data FROM files")
res = cur.fetchone()[0]
print(res == raw)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_sqlite_custom_aggregate_class() {
    let src = r#"
import sqlite3

class ProductAggregate:
    def __init__(self):
        self.product = 1
    def step(self, value):
        self.product *= value
    def finalize(self):
        return self.product

conn = sqlite3.connect(":memory:")
conn.create_aggregate("prod", 1, ProductAggregate)

cur = conn.cursor()
cur.execute("CREATE TABLE nums (v INT)")
cur.executemany("INSERT INTO nums VALUES (?)", [(2,), (3,), (4,)])

cur.execute("SELECT prod(v) FROM nums")
print(cur.fetchone()[0])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["24"]);
}
