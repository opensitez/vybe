# vybe-test: python/python_dis_disassemble_bytecode/test_dis_positions_attribute_in_311
# origin: languages/python/tests/python/test_python_dis_disassemble_bytecode.rs

import dis, sys

def g(x): return x + 1

inst = next(dis.get_instructions(g))
if sys.version_info >= (3, 11):
    print(hasattr(inst, "positions"))
else:
    print(True)
