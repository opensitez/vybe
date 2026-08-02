# vybe-test: python/python_dis_disassemble_bytecode/test_dis_findlabels_on_bytecode
# origin: languages/python/tests/python/test_python_dis_disassemble_bytecode.rs

import dis

def loop_func(n):
    for i in range(n):
        if i > 5:
            break

labels = dis.findlabels(loop_func.__code__.co_code)
print(isinstance(labels, list))
