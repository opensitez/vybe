# vybe-test: python/error_handling/try_finally_without_except
# origin: languages/python/tests/python/test_error_handling.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
import types as _t
f = _t.SimpleNamespace()
open('x', 'w').close()

try:
    f = open('x')
finally:
    f.close()
