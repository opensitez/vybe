# vybe-test: python/boolean_logic/or_short_circuit_skips_right_operand
# origin: languages/python/tests/python/test_boolean_logic.rs

def boom():
    print('boom')
    return True
print(True or boom())
