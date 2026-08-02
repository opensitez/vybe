# vybe-test: python/syntax/walrus_in_list
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

results = [y := f(x), y**2, y**3]
