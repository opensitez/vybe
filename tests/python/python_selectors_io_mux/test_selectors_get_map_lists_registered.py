# vybe-test: python/python_selectors_io_mux/test_selectors_get_map_lists_registered
# origin: languages/python/tests/python/test_python_selectors_io_mux.rs

import selectors, socket
sel = selectors.DefaultSelector()
socks = [socket.socket() for _ in range(3)]
for s in socks:
    s.setblocking(False)
    sel.register(s, selectors.EVENT_READ)
print(len(sel.get_map()) == 3)
sel.close()
for s in socks: s.close()
