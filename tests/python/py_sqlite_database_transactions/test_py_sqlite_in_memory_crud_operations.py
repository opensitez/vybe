# vybe-test: python/py_sqlite_database_transactions/test_py_sqlite_in_memory_crud_operations
# origin: languages/python/tests/python/test_py_sqlite_database_transactions.rs

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
