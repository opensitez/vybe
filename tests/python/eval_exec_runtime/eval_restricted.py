# vybe-test: python/eval_exec_runtime/eval_restricted
# origin: languages/python/tests/python/test_eval_exec_runtime.rs

eval('1', {'__builtins__': {}}, {})
