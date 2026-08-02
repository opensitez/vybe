# vybe-test: python/python_sqlite3_transactions_rows/test_sqlite3_fetchmany_size
# origin: languages/python/tests/python/test_python_sqlite3_transactions_rows.rs

import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("CREATE TABLE nums (n INT)")
cur.executemany("INSERT INTO nums VALUES (?)", [(i,) for i in range(10)])
cur.execute("SELECT n FROM nums")
rows = cur.fetchmany(3)
print(len(rows))
print([r[0] for r in rows])
