# vybe-test: python/pythonic_idioms/join_with_genexp
# origin: languages/python/tests/python/test_pythonic_idioms.rs

print('-'.join(str(x) for x in range(3)))
