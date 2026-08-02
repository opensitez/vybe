# vybe-test: python/py_ast_dis/test_py_ast_parse_and_walk
# origin: languages/python/tests/python/test_py_ast_dis.rs

import ast

code = "x = 10 + 20"
tree = ast.parse(code)

nodes = [type(n).__name__ for n in ast.walk(tree)]
print("Module" in nodes)
print("Assign" in nodes)
print("BinOp" in nodes)
print("Add" in nodes)
