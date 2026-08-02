# vybe-test: python/python_selectors_io_mux/test_selectors_multiple_register_same_fd_raises
# origin: languages/python/tests/python/test_python_selectors_io_mux.rs

import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
sel.register(sock, selectors.EVENT_READ)
try:
    sel.register(sock, selectors.EVENT_WRITE)
    print("no error")
except KeyError:
    print("KeyError")
sel.close()
sock.close()
