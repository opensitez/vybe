# vybe-test: python/python_sqlite3_transactions_rows/test_sqlite3_custom_collation
# origin: languages/python/tests/python/test_python_sqlite3_transactions_rows.rs

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
