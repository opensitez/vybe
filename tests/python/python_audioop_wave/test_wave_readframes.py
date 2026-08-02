# vybe-test: python/python_audioop_wave/test_wave_readframes
# origin: languages/python/tests/python/test_python_audioop_wave.rs

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

import wave, struct, tempfile, os

path = tempfile.mktemp(suffix='.wav')

values = [100, -100, 200, -200]
with wave.open(path, 'wb') as wf:
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(8000)
    data = struct.pack('<' + 'h' * len(values), *values)
    wf.writeframes(data)

with wave.open(path, 'rb') as wf:
    frames = wf.readframes(4)
    decoded = struct.unpack('<' + 'h' * 4, frames)
    __check(__line(list(decoded)), "[100, -100, 200, -200]")

os.unlink(path)
