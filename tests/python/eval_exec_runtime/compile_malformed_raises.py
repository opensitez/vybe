# vybe-test: python/eval_exec_runtime/compile_malformed_raises
# origin: languages/python/tests/python/test_eval_exec_runtime.rs

try:
 compile('1 +', '<s>', 'eval')
 print('ok')
except SyntaxError:
 print('err')
