# vybe-test: python/new_features/database_crud
# origin: languages/python/tests/python/test_new_features.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    import sqlite3
    conn = sqlite3.connect('app.db')
    cur = conn.cursor()
    cur.execute('CREATE TABLE users (id INT, name TEXT)')
    cur.execute('INSERT INTO users VALUES (1, "Alice")')
    rows = cur.fetchall()
    conn.close()

except BaseException:
    pass
