# vybe-test: python/python_dis_disassemble_bytecode/test_dis_bytecode_from_code_string
# origin: languages/python/tests/python/test_python_dis_disassemble_bytecode.rs

import dis
bc = dis.Bytecode("x = 1; y = 2")
opnames = [inst.opname for inst in bc]
print(any("STORE_NAME" in op or "STORE_FAST" in op for op in opnames))
