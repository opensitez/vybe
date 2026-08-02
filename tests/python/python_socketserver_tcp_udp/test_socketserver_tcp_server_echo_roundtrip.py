# vybe-test: python/python_socketserver_tcp_udp/test_socketserver_tcp_server_echo_roundtrip
# origin: languages/python/tests/python/test_python_socketserver_tcp_udp.rs

# Vybe test harness — Python.
#
# The Python counterpart of harness/go/check.go and harness/js/check.js: real
# source in the language under test, the way test262's assert.js is JavaScript.
#
# A test's verdict is its EXIT CODE. `__check` prints its diagnostic BEFORE
# raising, because an uncaught exception surfaces as `RuntimeError: [object]`
# and tells you nothing — 1,692 of testecma's 2,158 failures say exactly that.


def __line(*args):
    """One printed line, exactly as `print` composes it: str() each argument,
    joined by a single space.

    Written with a plain loop on purpose. `" ".join(str(a) for a in args)` —
    a GENERATOR EXPRESSION as a call argument — returns the empty string under
    Vybe while CPython gives "1 x", and a harness must not depend on anything
    the runtime under test gets wrong. A list comprehension works, but a loop
    needs nothing at all: no comprehension, no `join`, no `range`.
    """
    out = ""
    first = True
    for a in args:
        if not first:
            out += " "
        out += str(a)
        first = False
    return out


def __check(got, want):
    if got != want:
        print("FAIL: want [" + want + "] got [" + got + "]")
        raise Exception("assertion failed")

import socketserver, socket, threading

class EchoTCPHandler(socketserver.StreamRequestHandler):
    def handle(self):
        data = self.rfile.readline().strip()
        self.wfile.write(data + b"_ack\n")

server = socketserver.TCPServer(("127.0.0.1", 0), EchoTCPHandler)
ip, port = server.server_address

t = threading.Thread(target=server.handle_request)
t.start()

client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
client.connect((ip, port))
client.sendall(b"ping\n")
response = client.recv(1024).strip()
client.close()
server.server_close()
t.join()
__check(__line(response.decode()), "ping_ack")
