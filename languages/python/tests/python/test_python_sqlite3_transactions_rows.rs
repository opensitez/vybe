use super::helpers::run_python;

// sqlite3 — executescript, executemany, Row dict access, create_function, create_aggregate, transaction context manager, parameter binding (qmark and named)

#[test]
fn test_sqlite3_executescript_schema_and_data() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.executescript("""
CREATE TABLE users (id INT, name TEXT);
INSERT INTO users VALUES (1, 'Alice');
INSERT INTO users VALUES (2, 'Bob');
""")
cur.execute("SELECT COUNT(*) FROM users")
print(cur.fetchone()[0])
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_sqlite3_executemany_batch_insert() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE items (name TEXT, qty INT)")
items = [("apple", 10), ("banana", 20), ("cherry", 30)]
cur.executemany("INSERT INTO items VALUES (?, ?)", items)
cur.execute("SELECT SUM(qty) FROM items")
print(cur.fetchone()[0])
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_sqlite3_row_factory_dict_access() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
conn.row_factory = sqlite3.Row
cur = conn.cursor()
cur.execute("SELECT 1 AS id, 'Alice' AS name, 3.14 AS score")
row = cur.fetchone()
print(row["id"])
print(row["name"])
print(row["score"])
print(list(row.keys()))
"#,
    );
    assert_eq!(out, vec!["1", "Alice", "3.14", "['id', 'name', 'score']"]);
}

#[test]
fn test_sqlite3_custom_python_function() {
    let out = run_python(
        r#"
import sqlite3
def reverse_str(s):
    return s[::-1] if s else None

conn = sqlite3.connect(":memory:")
conn.create_function("rev", 1, reverse_str)
cur = conn.cursor()
cur.execute("SELECT rev('hello world')")
print(cur.fetchone()[0])
"#,
    );
    assert_eq!(out, vec!["dlrow olleh"]);
}

#[test]
fn test_sqlite3_custom_aggregate_function() {
    let out = run_python(
        r#"
import sqlite3
class ProductAgg:
    def __init__(self):
        self.prod = 1
    def step(self, value):
        if value is not None:
            self.prod *= value
    def finalize(self):
        return self.prod

conn = sqlite3.connect(":memory:")
conn.create_aggregate("product", 1, ProductAgg)
cur = conn.cursor()
cur.execute("CREATE TABLE nums (val INT)")
cur.executemany("INSERT INTO nums VALUES (?)", [(2,), (3,), (4,)])
cur.execute("SELECT product(val) FROM nums")
print(cur.fetchone()[0])
"#,
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn test_sqlite3_named_parameter_binding() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE t (a INT, b TEXT)")
cur.execute("INSERT INTO t VALUES (:val_a, :val_b)", {"val_a": 100, "val_b": "test_str"})
cur.execute("SELECT * FROM t WHERE a = :target", {"target": 100})
print(cur.fetchone())
"#,
    );
    assert_eq!(out, vec!["(100, 'test_str')"]);
}

#[test]
fn test_sqlite3_transaction_context_manager_commit() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
conn.execute("CREATE TABLE accounts (id INT, balance INT)")

with conn:
    conn.execute("INSERT INTO accounts VALUES (1, 500)")

cur = conn.cursor()
cur.execute("SELECT balance FROM accounts WHERE id = 1")
print(cur.fetchone()[0])
"#,
    );
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_sqlite3_transaction_context_manager_rollback() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
conn.execute("CREATE TABLE accounts (id INT UNIQUE, balance INT)")

try:
    with conn:
        conn.execute("INSERT INTO accounts VALUES (1, 100)")
        conn.execute("INSERT INTO accounts VALUES (1, 200)")  # duplicate key error
except sqlite3.IntegrityError:
    pass

cur = conn.cursor()
cur.execute("SELECT COUNT(*) FROM accounts")
print(cur.fetchone()[0])
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_sqlite3_blob_binary_data() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE files (name TEXT, data BLOB)")
data_bytes = b"\x00\x01\x02\xff\xfe"
cur.execute("INSERT INTO files VALUES (?, ?)", ("binary.dat", sqlite3.Binary(data_bytes)))
cur.execute("SELECT data FROM files")
res = cur.fetchone()[0]
print(res == data_bytes)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sqlite3_total_changes_attribute() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
conn.execute("CREATE TABLE t (x INT)")
conn.execute("INSERT INTO t VALUES (1)")
conn.execute("INSERT INTO t VALUES (2)")
print(conn.total_changes)
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_sqlite3_iterdump_database_backup() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
conn.execute("CREATE TABLE info (k TEXT, v TEXT)")
conn.execute("INSERT INTO info VALUES ('a', 'b')")
dump = "\n".join(conn.iterdump())
print("CREATE TABLE info" in dump)
print("INSERT INTO \"info\"" in dump or "INSERT INTO info" in dump or "a" in dump)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sqlite3_fetchmany_size() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE nums (n INT)")
cur.executemany("INSERT INTO nums VALUES (?)", [(i,) for i in range(10)])
cur.execute("SELECT n FROM nums")
rows = cur.fetchmany(3)
print(len(rows))
print([r[0] for r in rows])
"#,
    );
    assert_eq!(out, vec!["3", "[0, 1, 2]"]);
}

#[test]
fn test_sqlite3_lastrowid_attribute() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE items (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT)")
cur.execute("INSERT INTO items (name) VALUES ('widget')")
print(cur.lastrowid)
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_sqlite3_description_cursor_metadata() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("SELECT 42 AS num, 'foo' AS txt")
col_names = [desc[0] for desc in cur.description]
print(col_names)
"#,
    );
    assert_eq!(out, vec!["['num', 'txt']"]);
}

#[test]
fn test_sqlite3_rowcount_attribute() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE t (val INT)")
cur.executemany("INSERT INTO t VALUES (?)", [(1,), (2,), (3,)])
cur.execute("DELETE FROM t WHERE val > 1")
print(cur.rowcount)
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_sqlite3_custom_collation() {
    let out = run_python(
        r#"
import sqlite3
def ignore_case_cmp(s1, s2):
    return (s1.lower() > s2.lower()) - (s1.lower() < s2.lower())

conn = sqlite3.connect(":memory:")
conn.create_collation("nocase_custom", ignore_case_cmp)
cur = conn.cursor()
cur.execute("CREATE TABLE names (n TEXT)")
cur.executemany("INSERT INTO names VALUES (?)", [("b",), ("A",), ("c",)])
cur.execute("SELECT n FROM names ORDER BY n COLLATE nocase_custom")
print([r[0] for r in cur.fetchall()])
"#,
    );
    assert_eq!(out, vec!["['A', 'b', 'c']"]);
}

#[test]
fn test_sqlite3_sqlite_version() {
    let out = run_python(
        r#"
import sqlite3
print(isinstance(sqlite3.sqlite_version, str))
print(len(sqlite3.sqlite_version) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sqlite3_cursor_iterator() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE t (v INT)")
cur.executemany("INSERT INTO t VALUES (?)", [(10,), (20,)])
cur.execute("SELECT v FROM t")
vals = [r[0] for r in cur]
print(vals)
"#,
    );
    assert_eq!(out, vec!["[10, 20]"]);
}

#[test]
fn test_sqlite3_isolation_level_autocommit() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:", isolation_level=None)
print(conn.isolation_level is None)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sqlite3_in_memory_shared_cache() {
    let out = run_python(
        r#"
import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("SELECT sqlite_version()")
res = cur.fetchone()[0]
print(isinstance(res, str))
"#,
    );
    assert_eq!(out, vec!["True"]);
}
