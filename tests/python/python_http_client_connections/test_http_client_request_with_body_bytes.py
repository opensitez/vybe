# vybe-test: python/python_http_client_connections/test_http_client_request_with_body_bytes
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


# Output is COLLECTED, not paired. The emitter rewrites every `print(a, b)`
# into `__p(__line(a, b))`, appending here, and compares the whole buffer once
# at the end of the file. Pairing the i-th print with the i-th expected line
# cannot assert anything about a loop — 936 of Python's cases.
#
# `__p`/`__pr` take an ALREADY-JOINED string, not `*args` plus keyword-only
# terminator parameters. Those are broken under Vybe (measured: with a
# keyword-only sep/end after *args, the call appends nothing at all), while the
# plain `__line(*args)` above works. So the newline decision is made by WHICH
# helper the emitter calls.
#
# A comment in the FIRST position of an indented block used to be a parse error
# under Vybe — `def f():` followed by a comment line. Fixed in
# `languages/python/src/grammar.pest`: the preprocessor emits a comment-only
# line without an INDENT marker but still emits its NEWLINE, so `block` has to
# accept `":" NEWLINE NEWLINE* … INDENT`. Mid-block comments always worked,
# which is why this hid for so long.
__buf = ""


def __p(s):
    global __buf
    __buf += s + "\n"


def __pr(s):
    global __buf
    __buf += s


def __check(got, want):
    # The final print contributes a trailing newline that the expected line
    # vector never carried, so both forms are accepted.
    if got != want and got != want + "\n":
        print("FAIL: want [" + want + "] got [" + got + "]")
        raise Exception("assertion failed")

import http.client, io

class DummySocket:
    def __init__(self): self.buf = io.BytesIO()
    def sendall(self, data): self.buf.write(data)
    def close(self): pass

conn = http.client.HTTPConnection("dummy.host", 80)
conn.sock = DummySocket()
conn.request("POST", "/submit", body=b"payload_data", headers={"Content-Type": "text/plain"})
sent = conn.sock.buf.getvalue().decode()
__p(__line("Content-Length: 12" in sent))
__p(__line("payload_data" in sent))
conn.close()
__check(__buf, "True\nTrue")
