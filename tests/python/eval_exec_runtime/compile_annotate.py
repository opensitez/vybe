# vybe-test: python/eval_exec_runtime/compile_annotate
# origin: languages/python/tests/python/test_eval_exec_runtime.rs
# vybe-test-mode: compile

compile('x: int = 1', '<s>', 'exec')
