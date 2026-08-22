# vybe-test: python/context_manager_spec/ctx_manager_expr_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('x', 'w').close()

factory = open
with factory('x') as f:
    pass
