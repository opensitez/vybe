# vybe-test: python/context_manager_spec/ctx_in_try_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
def cleanup(*_a, **_k):
    return None
open('x', 'w').close()

try:
    with open('x') as f:
        pass
finally:
    cleanup()
