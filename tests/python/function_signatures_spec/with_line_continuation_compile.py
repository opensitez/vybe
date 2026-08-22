# vybe-test: python/function_signatures_spec/with_line_continuation_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('a', 'w').close()
open('b', 'w').close()

# EXTRACTION DAMAGE: a diff marker `+` leaked onto the continuation line.
with open('a') as f, \
     open('b') as g:
    pass
