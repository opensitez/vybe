# vybe-test: python/python_sqlite3_transactions_rows/test_sqlite3_cursor_iterator
# origin: languages/python/tests/python/test_python_sqlite3_transactions_rows.rs

import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE t (v INT)")
cur.executemany("INSERT INTO t VALUES (?)", [(10,), (20,)])
cur.execute("SELECT v FROM t")
vals = [r[0] for r in cur]
print(vals)
