# vybe-test: python/comprehensions/comp_with_ternary
# origin: languages/python/tests/python/test_comprehensions.rs

result = ['even' if x % 2 == 0 else 'odd' for x in range(5)]
