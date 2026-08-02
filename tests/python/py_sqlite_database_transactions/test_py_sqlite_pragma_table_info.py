# vybe-test: python/py_sqlite_database_transactions/test_py_sqlite_pragma_table_info
# origin: languages/python/tests/python/test_py_sqlite_database_transactions.rs

import sqlite3

conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE widget (id INT, name TEXT)")

cur.execute("PRAGMA table_info(widget)")
cols = [r[1] for r in cur.fetchall()]
print(cols)
conn.close()
