# vybe-test: python/py_ast_dis/test_py_dis_get_instructions
# origin: languages/python/tests/python/test_py_ast_dis.rs

import dis

def multiply(a, b):
    return a * b

instructions = list(dis.get_instructions(multiply))
opnames = [inst.opname for inst in instructions]
print("LOAD_FAST" in opnames)
print("RETURN_VALUE" in opnames)
