# vybe-test: python/inspect_dis_ast/ast_compile
# origin: languages/python/tests/python/test_inspect_dis_ast.rs
# There is no `ast.compile` — compilation is the BUILTIN. `ast` provides
# `parse`/`dump`/`unparse`.
import ast
compile('1+1', '<s>', 'eval')
