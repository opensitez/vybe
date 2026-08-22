# vybe-test: python/builtins/with_basic
# origin: languages/python/tests/python/test_builtins.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('file.txt', 'w').close()

with open('file.txt') as f:
    data = f.read()
