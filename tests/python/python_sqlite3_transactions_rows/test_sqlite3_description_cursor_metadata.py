# vybe-test: python/python_sqlite3_transactions_rows/test_sqlite3_description_cursor_metadata
# origin: languages/python/tests/python/test_python_sqlite3_transactions_rows.rs

import sqlite3
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
cur.execute("SELECT 42 AS num, 'foo' AS txt")
col_names = [desc[0] for desc in cur.description]
print(col_names)
