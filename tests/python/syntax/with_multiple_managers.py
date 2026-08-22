# vybe-test: python/syntax/with_multiple_managers
# origin: languages/python/tests/python/test_syntax.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('a', 'w').close()
open('b', 'w').close()

with open('a') as f1, open('b') as f2:
    pass
