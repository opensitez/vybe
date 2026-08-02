# vybe-test: python/inspect_dis_ast/ast_compile
# origin: languages/python/tests/python/test_inspect_dis_ast.rs
# vybe-test-mode: compile

import ast
ast.compile('1+1', '<s>', 'eval')
