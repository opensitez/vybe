# vybe-test: python/py_control_flow_loops/test_py_pass_statement_no_op
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

class Empty:
    pass

def stub():
    pass

for _ in range(3):
    pass

print("pass works")
