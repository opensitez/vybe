# vybe-test: python/python_http_client_connections/test_http_client_response_readinto_buffer
# origin: languages/python/tests/python/test_python_http_client_connections.rs

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

import http.client, io

class DummySocket:
    def __init__(self, data): self.file = io.BytesIO(data)
    def makefile(self, mode): return self.file
    def close(self): pass

resp = http.client.HTTPResponse(DummySocket(b"HTTP/1.1 200 OK\r\nContent-Length: 4\r\n\r\nABCD"))
resp.begin()
buf = bytearray(4)
n = resp.readinto(buf)
__check(__line(n), "4")
__check(__line(buf.decode()), "ABCD")
