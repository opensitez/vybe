# vybe-test: python/py_shutil_tempfile_archive_management/test_py_shutil_make_archive_unpack_zip
# origin: languages/python/tests/python/test_py_shutil_tempfile_archive_management.rs

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

import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmpdir:
    data_dir = os.path.join(tmpdir, "data")
    os.mkdir(data_dir)
    with open(os.path.join(data_dir, "hello.txt"), "w") as f:
        f.write("archive test")

    archive_path = shutil.make_archive(os.path.join(tmpdir, "my_archive"), "zip", data_dir)
    __check(__line(os.path.exists(archive_path)), "True")

    extract_dir = os.path.join(tmpdir, "extracted")
    shutil.unpack_archive(archive_path, extract_dir)
    __check(__line(os.path.exists(os.path.join(extract_dir, "hello.txt"))), "True")
