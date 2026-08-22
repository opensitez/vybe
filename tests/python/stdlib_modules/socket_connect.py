# vybe-test: python/stdlib_modules/socket_connect
# origin: languages/python/tests/python/test_stdlib_modules.rs

import socket
s = socket.create_connection(('localhost', 80))
