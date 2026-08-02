# vybe-test: python/boolean_logic/and_short_circuit_skips_right_operand
# origin: languages/python/tests/python/test_boolean_logic.rs

def boom():
    print('boom')
    return True
print(False and boom())
