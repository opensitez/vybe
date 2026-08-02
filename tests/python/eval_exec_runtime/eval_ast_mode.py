# vybe-test: python/eval_exec_runtime/eval_ast_mode
# origin: languages/python/tests/python/test_eval_exec_runtime.rs
# vybe-test-mode: compile

import ast
code = compile('1+1', '<s>', 'eval', flags=ast.PyCF_ONLY_AST)
