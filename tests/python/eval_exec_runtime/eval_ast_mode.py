# vybe-test: python/eval_exec_runtime/eval_ast_mode
# origin: languages/python/tests/python/test_eval_exec_runtime.rs

import ast
code = compile('1+1', '<s>', 'eval', flags=ast.PyCF_ONLY_AST)
