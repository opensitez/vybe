# vybe-test: python/python_logging_handlers_formatters/test_logging_custom_filter_class
# origin: languages/python/tests/python/test_python_logging_handlers_formatters.rs

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

import logging, io
class SecretFilter(logging.Filter):
    def filter(self, record):
        return "secret" not in record.getMessage()

stream = io.StringIO()
logger = logging.getLogger("test_filter")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
handler.addFilter(SecretFilter())
logger.addHandler(handler)

logger.info("normal log")
logger.info("this has secret data")
logger.info("another normal log")

logs = stream.getvalue().strip().split("\n")
__check(__line(len(logs)), "2")
__check(__line("secret" not in stream.getvalue()), "True")
