# vybe-test: python/eval_exec_runtime/exec_raise_caught
# origin: languages/python/tests/python/test_eval_exec_runtime.rs

ns = {}
try:
 exec('raise ValueError("e")', ns)
 print('ok')
except ValueError:
 print('err')
