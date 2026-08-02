# vybe-test: python/new_features/database_crud
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

import sqlite3
conn = sqlite3.connect('app.db')
cur = conn.cursor()
cur.execute('CREATE TABLE users (id INT, name TEXT)')
cur.execute('INSERT INTO users VALUES (1, "Alice")')
rows = cur.fetchall()
conn.close()
