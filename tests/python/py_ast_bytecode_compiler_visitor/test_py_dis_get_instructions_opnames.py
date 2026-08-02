# vybe-test: python/py_ast_bytecode_compiler_visitor/test_py_dis_get_instructions_opnames
# origin: languages/python/tests/python/test_py_ast_bytecode_compiler_visitor.rs

import dis

def add(x, y):
    return x + y

insts = list(dis.get_instructions(add))
opnames = [inst.opname for inst in insts]
print("LOAD_FAST" in opnames)
print("RETURN_VALUE" in opnames)
