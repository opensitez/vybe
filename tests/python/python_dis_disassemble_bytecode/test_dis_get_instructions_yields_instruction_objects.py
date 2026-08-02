# vybe-test: python/python_dis_disassemble_bytecode/test_dis_get_instructions_yields_instruction_objects
# origin: languages/python/tests/python/test_python_dis_disassemble_bytecode.rs

import dis

def add(a, b):
    return a + b

instructions = list(dis.get_instructions(add))
opnames = [inst.opname for inst in instructions]
print("BINARY_ADD" in opnames or "BINARY_OP" in opnames)
print(any(inst.opname == "RETURN_VALUE" for inst in instructions))
