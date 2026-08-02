# vybe-test: python/py_ast_bytecode_compiler_visitor/test_py_ast_parse_walk_node_types
# origin: languages/python/tests/python/test_py_ast_bytecode_compiler_visitor.rs

import ast

code = "def foo(a, b): return a + b"
tree = ast.parse(code)
node_types = [type(n).__name__ for n in ast.walk(tree)]

print("FunctionDef" in node_types)
print("Return" in node_types)
print("BinOp" in node_types)
