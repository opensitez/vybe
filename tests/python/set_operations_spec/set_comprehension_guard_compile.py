# vybe-test: python/set_operations_spec/set_comprehension_guard_compile
# origin: languages/python/tests/python/test_set_operations_spec.rs

s = {x * 2 for x in range(10) if x % 2 == 0}
