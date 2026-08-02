# vybe-test: python/stdlib_modules/sqlite3_connect
# origin: languages/python/tests/python/test_stdlib_modules.rs
# vybe-test-mode: compile

import sqlite3
conn = sqlite3.connect('test.db')
