# vybe-test: python/eval_exec_runtime/eval_name_error
# origin: languages/python/tests/python/test_eval_exec_runtime.rs

try:
 eval('undefined_xyz')
 print('ok')
except NameError:
 print('err')
