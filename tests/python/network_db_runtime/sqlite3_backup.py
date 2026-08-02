# vybe-test: python/network_db_runtime/sqlite3_backup
# origin: languages/python/tests/python/test_network_db_runtime.rs
# vybe-test-mode: compile

import sqlite3
src = sqlite3.connect(':memory:')
