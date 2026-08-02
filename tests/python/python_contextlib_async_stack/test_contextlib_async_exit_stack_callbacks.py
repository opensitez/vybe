# vybe-test: python/python_contextlib_async_stack/test_contextlib_async_exit_stack_callbacks
# origin: languages/python/tests/python/test_python_contextlib_async_stack.rs

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

import asyncio, contextlib

async def fn():
    cleaned_up = []
    async with contextlib.AsyncExitStack() as stack:
        stack.callback(lambda: cleaned_up.append("sync_cb"))
        async def async_cb(): cleaned_up.append("async_cb")
        stack.push_async_callback(async_cb)
    __check(__line(cleaned_up), "['async_cb', 'sync_cb']")

asyncio.run(fn())
