# vybe-test: python/raise_assert_extended/raise_in_generator
# origin: languages/python/tests/python/test_raise_assert_extended.rs
# vybe-test-mode: compile

def g():
 yield 1
 raise ValueError()
list(g())
