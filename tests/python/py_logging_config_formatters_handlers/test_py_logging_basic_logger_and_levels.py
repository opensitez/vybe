# vybe-test: python/py_logging_config_formatters_handlers/test_py_logging_basic_logger_and_levels
# origin: languages/python/tests/python/test_py_logging_config_formatters_handlers.rs

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

stream = io.StringIO()
logger = logging.getLogger("test_logger")
logger.setLevel(logging.INFO)

handler = logging.StreamHandler(stream)
handler.setFormatter(logging.Formatter("%(levelname)s:%(name)s:%(message)s"))
logger.addHandler(handler)

logger.debug("Debug msg ignored")
logger.info("Info msg logged")
logger.warning("Warning msg logged")

__check(__line(stream.getvalue().strip().splitlines()), "['INFO:test_logger:Info msg logged', 'WARNING:test_logger:Warning msg logged']")
