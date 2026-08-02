# vybe-test: python/python_selectors_io_mux/test_selectors_unregister_removes_key
# origin: languages/python/tests/python/test_python_selectors_io_mux.rs

import selectors, socket
sel = selectors.DefaultSelector()
sock = socket.socket()
sock.setblocking(False)
sel.register(sock, selectors.EVENT_READ)
sel.unregister(sock)
try:
    sel.unregister(sock)
    print("no error")
except KeyError:
    print("KeyError")
sel.close()
sock.close()
