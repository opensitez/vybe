# vybe-test: python/context_manager_spec/ctx_in_loop_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('a', 'w').close()
open('b', 'w').close()

for name in ['a', 'b']:
    with open(name) as f:
        print(name)
