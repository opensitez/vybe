# vybe-test: python/py_sqlite_db/test_py_sqlite_executemany_parameterized
# origin: languages/python/tests/python/test_py_sqlite_db.rs

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
