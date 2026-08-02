# vybe-test: python/python_dis_disassemble_bytecode/test_dis_instruction_is_jump_target
# origin: languages/python/tests/python/test_python_dis_disassemble_bytecode.rs

import dis

def cond(x):
    if x:
        return 1
    return 0

bc = dis.Bytecode(cond)
has_jump = any(inst.is_jump_target for inst in bc)
print(isinstance(has_jump, bool))
