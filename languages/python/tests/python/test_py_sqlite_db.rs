use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: sqlite3 — in-memory database, tables, insert, select, transactions, custom functions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_sqlite_create_table_insert_select() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()

cursor.execute("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT, age INTEGER)")
cursor.execute("INSERT INTO users (name, age) VALUES ('Alice', 30)")
cursor.execute("INSERT INTO users (name, age) VALUES ('Bob', 25)")
conn.commit()

cursor.execute("SELECT name, age FROM users ORDER BY age")
rows = cursor.fetchall()
print(rows)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["[('Bob', 25), ('Alice', 30)]"]);
}

#[test]
fn test_py_sqlite_executemany_parameterized() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()
cursor.execute("CREATE TABLE products (id INT, name TEXT, price REAL)")

items = [(1, "laptop", 999.99), (2, "mouse", 25.50), (3, "keyboard", 75.00)]
cursor.executemany("INSERT INTO products VALUES (?, ?, ?)", items)
conn.commit()

cursor.execute("SELECT name FROM products WHERE price > ?", (50.0,))
print([r[0] for r in cursor.fetchall()])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["['laptop', 'keyboard']"]);
}

#[test]
fn test_py_sqlite_row_factory_dict_access() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

cursor.execute("CREATE TABLE config (key TEXT, val TEXT)")
cursor.execute("INSERT INTO config VALUES ('theme', 'dark')")

cursor.execute("SELECT * FROM config WHERE key = ?", ('theme',))
row = cursor.fetchone()
print(row['key'])
print(row['val'])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["theme", "dark"]);
}

#[test]
fn test_py_sqlite_transaction_rollback() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()
cursor.execute("CREATE TABLE account (id INT, balance REAL)")
cursor.execute("INSERT INTO account VALUES (1, 100.0)")
conn.commit()

try:
    with conn:
        cursor.execute("UPDATE account SET balance = balance - 50 WHERE id = 1")
        raise RuntimeError("simulated error during transaction")
except RuntimeError:
    pass

cursor.execute("SELECT balance FROM account WHERE id = 1")
print(cursor.fetchone()[0])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["100.0"]);
}

#[test]
fn test_py_sqlite_create_function() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
conn.create_function("square", 1, lambda x: x * x)

cursor = conn.cursor()
cursor.execute("SELECT square(5)")
print(cursor.fetchone()[0])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["25"]);
}

#[test]
fn test_py_sqlite_aggregates() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()
cursor.execute("CREATE TABLE scores (val INT)")
cursor.executemany("INSERT INTO scores VALUES (?)", [(10,), (20,), (30,)])

cursor.execute("SELECT COUNT(*), SUM(val), AVG(val), MIN(val), MAX(val) FROM scores")
row = cursor.fetchone()
print(row[0])
print(row[1])
print(row[2])
print(row[3])
print(row[4])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["3", "60", "20.0", "10", "30"]);
}

#[test]
fn test_py_sqlite_blob_storage() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()
cursor.execute("CREATE TABLE files (name TEXT, data BLOB)")

binary_data = b"\x00\x01\x02\xff"
cursor.execute("INSERT INTO files VALUES (?, ?)", ("test.bin", sqlite3.Binary(binary_data)))
conn.commit()

cursor.execute("SELECT data FROM files WHERE name = ?", ("test.bin",))
blob = cursor.fetchone()[0]
print(blob == binary_data)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_sqlite_iterdump() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()
cursor.execute("CREATE TABLE t (id INT, val TEXT)")
cursor.execute("INSERT INTO t VALUES (1, 'hello')")

dump = "\n".join(conn.iterdump())
print("CREATE TABLE t" in dump)
print("INSERT INTO \"t\" VALUES(1,'hello')" in dump)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sqlite_lastrowid_rowcount() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()
cursor.execute("CREATE TABLE items (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT)")

cursor.execute("INSERT INTO items (name) VALUES ('item1')")
print(cursor.lastrowid)

cursor.execute("INSERT INTO items (name) VALUES ('item2')")
print(cursor.lastrowid)

cursor.execute("UPDATE items SET name = 'updated'")
print(cursor.rowcount)
conn.close()
"#;
    assert_eq!(run_python(src), vec!["1", "2", "2"]);
}

#[test]
fn test_py_sqlite_pragma_queries() {
    let src = r#"
import sqlite3

conn = sqlite3.connect(":memory:")
cursor = conn.cursor()
cursor.execute("PRAGMA user_version = 42")
cursor.execute("PRAGMA user_version")
print(cursor.fetchone()[0])
conn.close()
"#;
    assert_eq!(run_python(src), vec!["42"]);
}
