# vybe-test: python/network_db_runtime/sqlite3_exception
# origin: languages/python/tests/python/test_network_db_runtime.rs

import sqlite3
try:
 sqlite3.connect(':memory:').execute('INVALID SQL')
 print('ok')
except sqlite3.Error:
 print('err')
