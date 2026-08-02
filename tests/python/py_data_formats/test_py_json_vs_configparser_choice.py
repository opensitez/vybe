# vybe-test: python/py_data_formats/test_py_json_vs_configparser_choice
# origin: languages/python/tests/python/test_py_data_formats.rs

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

import json, configparser

# JSON preserves types
json_config = json.loads('{"port": 8080, "debug": true, "tags": ["a", "b"]}')
__check(__line(type(json_config["port"]).__name__), "int")
__check(__line(type(json_config["debug"]).__name__), "bool")
__check(__line(json_config["tags"]), "['a', 'b']")

# ConfigParser keeps everything as strings
cfg = configparser.ConfigParser()
cfg.read_string("[app]\nport=8080\ndebug=true")
__check(__line(type(cfg["app"]["port"]).__name__), "str")
__check(__line(cfg.getboolean("app", "debug")), "True")
