# vybe-test: python/syntax/with_nested
# origin: languages/python/tests/python/test_syntax.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('a', 'w').close()
open('b', 'w').close()

with open('a') as f:
    with open('b') as g:
        pass
