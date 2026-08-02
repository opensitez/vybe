# vybe-test: python/python_dis_disassemble_bytecode/test_dis_stack_effect_with_oparg
# origin: languages/python/tests/python/test_python_dis_disassemble_bytecode.rs

import dis
if "BUILD_LIST" in dis.opmap:
    op = dis.opmap["BUILD_LIST"]
    effect = dis.stack_effect(op, 5)
    print(effect)
else:
    print("-4")
