# vybe-test: python/py_concurrency/test_py_threading_barrier
# origin: languages/python/tests/python/test_py_concurrency.rs

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

import threading

log = []
barrier = threading.Barrier(3)

def task(name):
    log.append(f"{name}_before")
    barrier.wait()
    log.append(f"{name}_after")

threads = [threading.Thread(target=task, args=(f"T{i}",)) for i in range(3)]
for t in threads:
    t.start()
for t in threads:
    t.join()

befores = sorted(x for x in log if "before" in x)
afters = sorted(x for x in log if "after" in x)
__p(__line(befores))
__p(__line(afters))
__check(__buf, "['T0_before', 'T1_before', 'T2_before']\n['T0_after', 'T1_after', 'T2_after']")
