# vybe-test: python/comprehensions/comp_with_method_chain
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

result = [word.strip().lower() for word in lines]
