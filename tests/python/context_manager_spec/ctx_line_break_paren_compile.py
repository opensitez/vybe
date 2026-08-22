# vybe-test: python/context_manager_spec/ctx_line_break_paren_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('a', 'w').close()
open('b', 'w').close()

with (
    open('a') as f,
    open('b') as g
):
    pass
