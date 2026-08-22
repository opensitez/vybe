# vybe-test: python/network_db_runtime/socket_connect_ex
# origin: languages/python/tests/python/test_network_db_runtime.rs

import socket
s = socket.socket()
s.connect_ex(('127.0.0.1', 9))
