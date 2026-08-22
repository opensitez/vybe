# vybe-test: python/inspect_dis_ast/dis_dis_function
# origin: languages/python/tests/python/test_inspect_dis_ast.rs

import dis
def f(): return 1
dis.dis(f)
