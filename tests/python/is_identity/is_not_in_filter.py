# vybe-test: python/is_identity/is_not_in_filter
# origin: languages/python/tests/python/test_is_identity.rs

a = [1, 2]
b = a
pairs = [(a, b), ([1], [1])]
print(sum(1 for x, y in pairs if x is y))
